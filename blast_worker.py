
import os
import time
import base64
import queue
import tempfile
import logging
import threading
from datetime import datetime, date, timedelta
from urllib.parse import urlparse, unquote, parse_qsl

import requests
import win32com.client as win32
import pythoncom
from dotenv import load_dotenv

load_dotenv()

SUPABASE_URL = os.environ["SUPABASE_URL"]
SERVICE_ROLE_KEY = os.environ["SUPABASE_SERVICE_ROLE_KEY"]
BAILEYS_LOCAL_URL = os.environ.get("BAILEYS_LOCAL_URL", "http://localhost:4000")
REPORT_OUTPUT_DIR = os.environ.get("REPORT_OUTPUT_DIR", tempfile.gettempdir())

# Berapa Excel.Application paralel yang boleh jalan bareng.
# Naikkan pelan-pelan sambil pantau RAM/CPU -- tiap instance itu 1 proses Excel penuh,
# apalagi kalau RefreshAll narik data eksternal yang berat.
MAX_WORKERS = 1  # DIPAKSA 1 KARENA KITA MENGGUNAKAN ms-excel: (menghindari rebutan instance)

# Timeout nunggu workbook kebuka lewat fallback protocol handler.
OPEN_TIMEOUT_SEC = int(os.environ.get("EXCEL_OPEN_TIMEOUT_SEC", "45"))

# Set FORCE_RUN_ALL=1 di env kalau lagi testing dan mau paksa semua job active
# jalan tanpa peduli jam_blast/tanggal_blast. JANGAN nyala di production.
FORCE_RUN_ALL = os.environ.get("FORCE_RUN_ALL", "0") == "1"

HEADERS = {
    "apikey": SERVICE_ROLE_KEY,
    "Authorization": f"Bearer {SERVICE_ROLE_KEY}",
    "Content-Type": "application/json",
}

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] [%(threadName)s] %(message)s",
    handlers=[
        logging.FileHandler("worker.log"),
        logging.StreamHandler(),
    ],
)
log = logging.getLogger(__name__)

# Clipboard itu resource shared di level OS. Walaupun tiap thread punya
# Excel.Application sendiri-sendiri, operasi "copy range -> paste" TETAP
# harus 1-per-1 di seluruh proses, kalau enggak gambar antar-job bisa ketuker.
CLIPBOARD_LOCK = threading.Lock()


# ---------------------------------------------------------------------------
# Supabase helpers
# ---------------------------------------------------------------------------

def get_active_jobs():
    """Ambil semua job berstatus active, lengkap sama sheets/ranges/groups-nya."""
    resp = requests.get(
        f"{SUPABASE_URL}/rest/v1/form_request",
        headers=HEADERS,
        params={
            "status": "eq.active",
            "select": (
                "id,nama_automation,link,jam_blast,tanggal_blast,caption,last_run_date,"
                "request_sheets(id,sheet_name,sheet_tables(cell_range)),"
                "request_groups(wa_groups(group_jid,group_name,wa_account_id))"
            ),
        },
    )
    resp.raise_for_status()
    return resp.json()


def mark_job_done(job_id: str):
    resp = requests.patch(
        f"{SUPABASE_URL}/rest/v1/form_request",
        headers={**HEADERS, "Prefer": "return=minimal"},
        params={"id": f"eq.{job_id}"},
        json={"last_run_date": str(date.today())},
    )
    resp.raise_for_status()


TOLERANCE_MINUTES = 15


def is_due(job, now):
    """
    Job dianggap due kalau:
    1. tanggal_blast cocok hari ini (format 'YYYY-MM-DD')
    2. waktu sekarang ada di window [jam_blast, jam_blast + TOLERANCE_MINUTES]
    3. belum pernah dijalanin hari ini (guard last_run_date)
    """
    tanggal_blast_str = job.get("tanggal_blast", "")
    hari_ini = str(now.day)
    
    # Cek apakah tanggal hari ini ada di dalam list (dipisahkan dengan koma)
    daftar_tanggal = [t.strip() for t in tanggal_blast_str.split(",") if t.strip()]
    if hari_ini not in daftar_tanggal:
        return False

    jam_str = job["jam_blast"]  # format "HH:MM:SS"
    h, m, s = map(int, jam_str.split(":"))
    scheduled = now.replace(hour=h, minute=m, second=s, microsecond=0)
    window_end = scheduled + timedelta(minutes=TOLERANCE_MINUTES)

    if not (scheduled <= now <= window_end):
        return False

    last_run = job.get("last_run_date")
    if last_run == str(date.today()):
        return False

    return True


# ---------------------------------------------------------------------------
# Excel instance management
# ---------------------------------------------------------------------------

def make_excel_instance():
    pythoncom.CoInitialize()
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = True  # KITA BUAT VISIBLE agar kelihatan apakah ada pop-up error dari SharePoint
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False
    return excel


def close_excel_app(excel):
    if excel:
        try:
            excel.Quit()
        except Exception:
            pass
    try:
        pythoncom.CoUninitialize()
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Workbook opening: langsung lewat COM, fallback ke protocol handler kalau perlu
# ---------------------------------------------------------------------------

def extract_filename_hint(link: str) -> str:
    parsed = urlparse(link)
    qs = dict(parse_qsl(parsed.query))
    for key in ("file", "sourcedoc", "SourceDoc", "originalPath"):
        if key in qs:
            return unquote(qs[key]).strip("{}").lower()
    return unquote(parsed.path.rsplit("/", 1)[-1]).lower()


def open_workbook(excel, link: str, job_name: str):
    # Dapatkan nama-nama workbook yang SEDANG terbuka sekarang
    known_names = {wb.Name.lower() for wb in excel.Workbooks}
    
    log.info(f"[{job_name}] Workbook yang sudah terbuka sebelum eksekusi: {known_names}")

    # Buka paksa via protocol handler (menjamin SharePoint auth & Open in Desktop jalan)
    safe_link = link.replace("|", "^|")
    os.system(f"start ms-excel:ofe^|u^|{safe_link}")

    log.info(f"[{job_name}] Menunggu Excel merespons via protocol handler (maks {OPEN_TIMEOUT_SEC}s)...")
    deadline = time.time() + OPEN_TIMEOUT_SEC
    while time.time() < deadline:
        time.sleep(2)
        try:
            for wb in excel.Workbooks:
                name_lower = wb.Name.lower()
                if name_lower in known_names:
                    continue  # Ini file yang sudah terbuka dari tadi, abaikan
                
                log.info(f"[{job_name}] Ditemukan workbook baru terbuka: '{name_lower}'")
                
                # Paksa Excel untuk memunculkan window-nya ke depan layar
                try:
                    excel.Visible = True
                    wb.Activate()
                except:
                    pass
                
                return wb
                
        except Exception as e:
            log.warning(f"[{job_name}] Error saat mengecek Workbooks: {e}")

    return None


# ---------------------------------------------------------------------------
# Screenshot range: lewat Chart.Export, bukan screen-grab
# ---------------------------------------------------------------------------

def capture_range_image_from_wb(wb, sheet_name: str, cell_range: str, job_name: str) -> str:
    """
    Dulu: CopyPicture -> clipboard OS -> ImageGrab.grabclipboard() (PIL).
    Masalahnya itu tergantung render layar & window focus, dan gampang salah
    ambil kalau ada operasi copy-paste lain (dari job lain) di waktu yang sama.

    Sekarang: CopyPicture -> paste ke ChartObject sementara -> Chart.Export(png).
    Ini murni lewat object model Excel (gak butuh window ke-render di layar,
    gak peduli window mana yang fokus). Tetap lewat clipboard secara internal,
    makanya operasi copy+paste-nya dibungkus CLIPBOARD_LOCK.
    """
    ws = wb.Sheets(sheet_name)
    rng = ws.Range(cell_range)

    safe_sheet = "".join(c for c in sheet_name if c.isalnum() or c in (" ", "_")).strip()
    safe_range = cell_range.replace(":", "_")
    safe_job = "".join(c for c in job_name if c.isalnum() or c in (" ", "_")).strip()
    filename = f"{safe_job}_{safe_sheet}_{safe_range}_{int(time.time() * 1000)}.png"

    os.makedirs(REPORT_OUTPUT_DIR, exist_ok=True)
    out_path = os.path.join(REPORT_OUTPUT_DIR, filename)

    with CLIPBOARD_LOCK:
        rng.CopyPicture(Appearance=1, Format=2)  # xlScreen, xlBitmap
        chart_obj = ws.ChartObjects().Add(0, 0, rng.Width, rng.Height)
        try:
            chart_obj.Chart.Paste()
            chart_obj.Chart.Export(out_path, "PNG")
        finally:
            chart_obj.Delete()

    if not os.path.exists(out_path):
        raise RuntimeError(f"Gagal export gambar untuk range {cell_range}")

    return out_path


def image_to_base64(path: str) -> str:
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode("utf-8")


def send_whatsapp(wa_account_id: str, group_jid: str, caption: str, image_paths: list):
    """Kirim ke Baileys Manager (proses Node.js lokal yang nyala terus di server ini)."""
    images_b64 = [image_to_base64(p) for p in image_paths]

    resp = requests.post(
        f"{BAILEYS_LOCAL_URL}/send",
        json={
            "waAccountId": wa_account_id,
            "groupJid": group_jid,
            "caption": caption,
            "images": images_b64,
        },
        timeout=60,
    )
    resp.raise_for_status()
    return resp.json()


# ---------------------------------------------------------------------------
# Job processing
# ---------------------------------------------------------------------------

def process_job(job, excel):
    job_name = job["nama_automation"]
    log.info(f"Memproses job: {job_name} ({job['id']})")
    wb = None
    try:
        wb = open_workbook(excel, job["link"], job_name)
        if not wb:
            raise RuntimeError(
                f"Gagal mendapatkan workbook untuk job {job['id']}. "
                f"Pastikan login Microsoft tidak terputus / link masih valid."
            )

        log.info(f"[{job_name}] Attached ke workbook: {wb.Name}")

        log.info(f"[{job_name}] Me-refresh data Excel (RefreshAll)...")
        try:
            wb.RefreshAll()
            time.sleep(5)  # kasih waktu rendering setelah refresh
        except Exception as e:
            log.warning(f"[{job_name}] Error saat RefreshAll (bisa diabaikan kalau gak ada external data): {e}")

        image_paths = []
        for sheet in job.get("request_sheets", []):
            sheet_name = sheet["sheet_name"]
            for tbl in sheet.get("sheet_tables", []):
                cell_range = tbl["cell_range"]
                img_path = capture_range_image_from_wb(wb, sheet_name, cell_range, job_name)
                image_paths.append(img_path)

        targets = [
            (rg["wa_groups"]["wa_account_id"], rg["wa_groups"]["group_jid"])
            for rg in job.get("request_groups", [])
            if rg.get("wa_groups")
        ]

        if not targets:
            log.warning(f"[{job_name}] Gak punya grup tujuan, skip kirim WA")
        else:
            for wa_account_id, group_jid in targets:
                send_whatsapp(wa_account_id, group_jid, job.get("caption") or "", image_paths)
                log.info(f"[{job_name}]  -> terkirim ke grup {group_jid} (akun {wa_account_id})")

        mark_job_done(job["id"])
        log.info(f"[{job_name}] Job {job['id']} selesai, last_run_date diupdate")
    finally:
        try:
            if wb:
                wb.Close(SaveChanges=False)
        except Exception:
            pass


# ---------------------------------------------------------------------------
# Worker thread
# ---------------------------------------------------------------------------

def worker_loop(job_queue: "queue.Queue"):
    excel = make_excel_instance()
    try:
        while True:
            try:
                job = job_queue.get_nowait()
            except queue.Empty:
                break
            try:
                process_job(job, excel)
            except Exception as e:
                log.error(f"Gagal proses job {job['id']} ({job['nama_automation']}): {e}")
                # sengaja lanjut ke job berikutnya di queue, gak stop semua
            finally:
                job_queue.task_done()
    finally:
        close_excel_app(excel)


def main():
    now = datetime.now()
    log.info(f"=== Worker jalan: {now.isoformat()} ===")

    try:
        jobs = get_active_jobs()
    except Exception as e:
        log.error(f"Gagal ambil job dari Supabase: {e}")
        return

    if FORCE_RUN_ALL:
        due_jobs = jobs
        log.warning(f"FORCE_RUN_ALL aktif -- semua {len(jobs)} job active dipaksa jalan, abaikan jam/tanggal_blast")
    else:
        due_jobs = [j for j in jobs if is_due(j, now)]

    log.info(f"Total job active: {len(jobs)}, yang due dijalankan: {len(due_jobs)}")

    if not due_jobs:
        log.info("Gak ada job due, worker selesai.")
        return

    job_queue = queue.Queue()
    for j in due_jobs:
        job_queue.put(j)

    n_workers = min(MAX_WORKERS, len(due_jobs))
    log.info(f"Menjalankan {n_workers} worker thread paralel (MAX_EXCEL_WORKERS={MAX_WORKERS})")

    threads = []
    for i in range(n_workers):
        t = threading.Thread(target=worker_loop, args=(job_queue,), name=f"ExcelWorker-{i+1}")
        t.start()
        threads.append(t)

    for t in threads:
        t.join()

    log.info("=== Worker selesai ===")


if __name__ == "__main__":
    main()
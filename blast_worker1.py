import json
import os
import re
import shutil
import subprocess
import sys
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timedelta
from pathlib import Path

import pythoncom
import requests
import urllib.parse

# Hapus cache gen_py secara paksa sebelum memanggil win32com jika ada corrupt
import win32com
try:
    if hasattr(win32com, '__gen_path__') and win32com.__gen_path__:
        import shutil
        if os.path.exists(win32com.__gen_path__):
            shutil.rmtree(win32com.__gen_path__, ignore_errors=True)
    
    _gen_py_path = os.path.join(os.path.dirname(win32com.__file__), 'gen_py')
    if os.path.exists(_gen_py_path):
        import shutil
        shutil.rmtree(_gen_py_path, ignore_errors=True)
except Exception:
    pass

import win32com.client as win32
import win32com.client.gencache as _gencache
from dotenv import load_dotenv
from PIL import ImageGrab

# Paksa win32com pakai dynamic dispatch untuk menghindari error gen_py cache corrupt
_gencache.is_readonly = True

load_dotenv()

SUPABASE_URL = os.environ.get("SUPABASE_URL", "").strip().rstrip("/")
SUPABASE_SERVICE_ROLE_KEY = os.environ.get("SUPABASE_SERVICE_ROLE_KEY", "").strip()
TOLERANCE_MINUTES = 15
MAX_AUTOMATIONS_PER_BATCH = 3
MAX_WORKERS = min(
    MAX_AUTOMATIONS_PER_BATCH,
    max(1, int(os.environ.get("MAX_WORKERS", str(MAX_AUTOMATIONS_PER_BATCH)))),
)
CHROME_EMAIL = os.environ.get("CHROME_EMAIL", "").strip()
CHROME_USER_DATA_DIR = os.path.expandvars(os.environ.get("CHROME_USER_DATA_DIR", r"%LOCALAPPDATA%\Google\Chrome\User Data"))
EXCEL_OPEN_DELAY_SEC = float(os.environ.get("EXCEL_OPEN_DELAY_SEC", "5"))
EXCEL_OPEN_TIMEOUT_SEC = float(os.environ.get("EXCEL_OPEN_TIMEOUT_SEC", "180"))
EXCEL_WORKBOOK_POLL_SEC = float(os.environ.get("EXCEL_WORKBOOK_POLL_SEC", "3"))
EXCEL_REFRESH_TIMEOUT_SEC = float(os.environ.get("EXCEL_REFRESH_TIMEOUT_SEC", "300"))
EXCEL_REFRESH_POLL_SEC = float(os.environ.get("EXCEL_REFRESH_POLL_SEC", "2"))
CLIPBOARD_IMAGE_TIMEOUT_SEC = float(os.environ.get("CLIPBOARD_IMAGE_TIMEOUT_SEC", "3"))
SCRIPT_DIR = Path(__file__).resolve().parent
SCREENSHOT_DIR = SCRIPT_DIR / "screenshots"
DUE_IDS_TXT = SCRIPT_DIR / "due_automation_ids.txt"
CHROME_HEADLESS = os.environ.get("CHROME_HEADLESS", "1") == "1"
CLIPBOARD_LOCK = threading.Lock()
CHROME_LAUNCH_LOCK = threading.Lock()
SCREENSHOT_WORK_LOCK = threading.Lock()


def supabase_headers() -> dict[str, str]:
    if not SUPABASE_URL or not SUPABASE_SERVICE_ROLE_KEY:
        raise ValueError("SUPABASE_URL dan SUPABASE_SERVICE_ROLE_KEY wajib diisi di .env.")
    return {
        "apikey": SUPABASE_SERVICE_ROLE_KEY,
        "Authorization": f"Bearer {SUPABASE_SERVICE_ROLE_KEY}",
        "Content-Type": "application/json",
        "Prefer": "return=minimal",
    }


def get_active_jobs() -> list[dict]:
    response = requests.get(
        f"{SUPABASE_URL}/rest/v1/form_request",
        headers=supabase_headers(),
        params={
            "status": "eq.active",
            "select": "id,nama_automation,link,nama_file,jam_blast,tanggal_blast,last_run_date,caption",
        },
        timeout=30,
    )
    response.raise_for_status()
    return response.json()


def load_job_sheets(job: dict) -> None:
    """Load sheets by form_request.id, then ranges by each request_sheets.id."""
    request_id = str(job["id"])
    sheets_response = requests.get(
        f"{SUPABASE_URL}/rest/v1/request_sheets",
        headers=supabase_headers(),
        params={"request_id": f"eq.{request_id}", "select": "id,sheet_name"},
        timeout=30,
    )
    sheets_response.raise_for_status()
    sheets = sheets_response.json()
    for sheet in sheets:
        tables_response = requests.get(
            f"{SUPABASE_URL}/rest/v1/sheet_tables",
            headers=supabase_headers(),
            params={"sheet_id": f"eq.{sheet['id']}", "select": "id,cell_range"},
            timeout=30,
        )
        tables_response.raise_for_status()
        sheet["sheet_tables"] = tables_response.json()
    job["request_sheets"] = sheets
    print(
        f"{job.get('nama_automation') or request_id}: "
        f"{len(sheets)} sheet, "
        f"{sum(len(sheet['sheet_tables']) for sheet in sheets)} range dibaca dari database."
    )


def load_job_groups(job: dict) -> None:
    request_id = str(job["id"])
    response = requests.get(
        f"{SUPABASE_URL}/rest/v1/request_groups",
        headers=supabase_headers(),
        params={
            "request_id": f"eq.{request_id}", 
            "select": "wa_group_id,wa_groups(id,group_name,group_jid)"
        },
        timeout=30,
    )
    response.raise_for_status()
    
    groups = []
    for row in response.json():
        if row.get("wa_groups"):
            groups.append(row["wa_groups"])
            
    job["wa_groups"] = groups
    print(
        f"{job.get('nama_automation') or request_id}: tujuan grup = "
        f"{[group['group_name'] for group in groups]}"
    )


def is_due(job: dict, now: datetime) -> bool:
    dates = {value.strip() for value in str(job.get("tanggal_blast") or "").split(",") if value.strip()}
    if str(now.day) not in dates:
        return False
    blast_time = str(job.get("jam_blast") or "").split(".", 1)[0]
    scheduled_time = None
    for time_format in ("%H:%M:%S", "%H:%M"):
        try:
            scheduled_time = datetime.strptime(blast_time, time_format).time()
            break
        except ValueError:
            pass
    if scheduled_time is None:
        print(f"Job {job.get('id')} dilewati: jam_blast tidak valid ({blast_time!r}).")
        return False
    scheduled = datetime.combine(now.date(), scheduled_time)
    return scheduled <= now <= scheduled + timedelta(minutes=TOLERANCE_MINUTES)


def find_chrome() -> str:
    candidates = [
        shutil.which("chrome.exe"),
        os.path.expandvars(r"%PROGRAMFILES%\Google\Chrome\Application\chrome.exe"),
        os.path.expandvars(r"%PROGRAMFILES(X86)%\Google\Chrome\Application\chrome.exe"),
        os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
    ]
    for candidate in candidates:
        if candidate and Path(candidate).is_file():
            return candidate
    raise FileNotFoundError("Google Chrome tidak ditemukan di komputer ini.")


def excel_is_running() -> bool:
    result = subprocess.run(
        ["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/NH"],
        capture_output=True, text=True, check=False,
    )
    return "excel.exe" in result.stdout.casefold()


def kill_stale_excel() -> None:
    """Matikan semua proses EXCEL.EXE yang tidak memiliki window XLMAIN.

    Excel sisa dari run sebelumnya (background/stale) akan menghambat deteksi
    workbook baru. Fungsi ini membersihkannya sebelum automation dimulai.
    """
    result = subprocess.run(
        ["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/NH", "/FO", "CSV"],
        capture_output=True, text=True, check=False,
    )
    excel_pids = []
    for line in result.stdout.splitlines():
        parts = line.strip('"').split('","')
        if len(parts) >= 2 and "excel" in parts[0].lower():
            try:
                excel_pids.append(int(parts[1].strip('"')))
            except ValueError:
                pass

    if not excel_pids:
        return

    # Cek apakah ada window XLMAIN (workbook terbuka) via tasklist window title
    # Kalau tidak ada XLMAIN, proses ini adalah stale — matikan
    wnd_check = subprocess.run(
        ["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/V", "/NH"],
        capture_output=True, text=True, check=False,
    )
    has_real_window = "xlmain" in wnd_check.stdout.casefold()

    # Cek via PowerShell apakah ada window XLMAIN
    ps_check = subprocess.run(
        ["powershell", "-NoProfile", "-NonInteractive", "-Command",
         "(Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object {$_.MainWindowTitle -ne ''}).Count"],
        capture_output=True, text=True, check=False, timeout=10,
    )
    has_main_window = ps_check.stdout.strip() not in ("", "0")

    if not has_real_window and not has_main_window:
        print(f"[CLEANUP] Ditemukan {len(excel_pids)} proses EXCEL.EXE stale (tanpa workbook), menghentikannya...")
        subprocess.run(["taskkill", "/F", "/IM", "EXCEL.EXE"], capture_output=True, check=False)
        time.sleep(2)
        print("[CLEANUP] Proses Excel stale sudah dihentikan.")
    else:
        print(f"[INFO] Excel sedang berjalan dengan workbook aktif, tidak dihentikan.")


def chrome_is_running() -> bool:
    result = subprocess.run(
        ["tasklist", "/FI", "IMAGENAME eq chrome.exe", "/NH"],
        capture_output=True, text=True, check=False,
    )
    return "chrome.exe" in result.stdout.casefold()


def find_profile_for_email() -> str:
    if not CHROME_EMAIL:
        raise ValueError("CHROME_EMAIL belum diisi di file .env.")
    local_state_path = Path(CHROME_USER_DATA_DIR) / "Local State"
    if not local_state_path.is_file():
        raise FileNotFoundError(f"File profil Chrome tidak ditemukan: {local_state_path}")
    with local_state_path.open("r", encoding="utf-8") as state_file:
        profiles = json.load(state_file).get("profile", {}).get("info_cache", {})
    target_email = CHROME_EMAIL.casefold()
    for profile_directory, profile_info in profiles.items():
        known_emails = {str(profile_info.get("user_name", "")).casefold(), str(profile_info.get("gaia_name", "")).casefold()}
        if target_email in known_emails:
            return profile_directory
    raise ValueError(f"Profil Chrome untuk {CHROME_EMAIL} tidak ditemukan.")


def normalize_workbook_name(name: str) -> str:
    """Hapus ekstensi, zero-width space, whitespace berlebih, url decode, lalu lowercase."""
    name_unquoted = urllib.parse.unquote(str(name))
    normalized = " ".join(name_unquoted.replace("\u200b", "").replace("\xa0", " ").split()).casefold()
    # Hapus ekstensi Excel dan suffix Excel seperti [1], [2] (copy protection)
    normalized = re.sub(r"\s*\[\d+\]$", "", normalized)
    normalized = re.sub(r"\.(xlsx|xlsb|xlsm|xls)$", "", normalized)
    return normalized.strip()


def workbook_candidate_names(workbook: object) -> set[str]:
    names = {str(workbook.Name)}
    try:
        names.add(Path(str(workbook.FullName)).name)
    except Exception:
        pass
    return {normalize_workbook_name(name) for name in names}


def workbook_matches_expected(workbook: object, expected_norm: str) -> bool:
    """Cek apakah workbook cocok dengan nama yang diharapkan.
    
    Strategi (urutan prioritas):
    1. Exact match setelah normalisasi
    2. Expected name adalah substring dari nama workbook (toleran untuk suffix tambahan)
    3. Nama workbook adalah substring dari expected name (toleran untuk nama terpotong)
    """
    if not expected_norm:
        return False
    candidates = workbook_candidate_names(workbook)
    for candidate in candidates:
        if candidate == expected_norm:
            return True
        if expected_norm in candidate:
            return True
        if candidate in expected_norm:
            return True
    return False


def get_all_open_workbooks() -> list[object]:
    """Ambil semua workbook dari Excel yang berjalan menggunakan Late Binding murni.

    Ini menghindari error win32com.gen_py cache corrupt secara permanen
    karena tidak membaca typelib sama sekali.
    """
    workbooks: list[object] = []
    seen_names: set[str] = set()

    try:
        # Panggil Excel via late binding murni
        excel = win32.dynamic.Dispatch("Excel.Application")
        count = excel.Workbooks.Count
        for i in range(1, count + 1):
            try:
                wb = excel.Workbooks(i)
                name = str(wb.Name)
                if name not in seen_names:
                    workbooks.append(wb)
                    seen_names.add(name)
            except Exception:
                pass
    except Exception as e:
        print(f"[WARN] Gagal mendapatkan workbooks via dynamic dispatch: {e}")

    return workbooks


def wait_for_workbook(expected_name: str) -> object:
    """Poll setiap EXCEL_WORKBOOK_POLL_SEC detik hingga workbook yang cocok ditemukan.
    
    Pencocokan menggunakan normalize_workbook_name + substring matching agar
    toleran terhadap perbedaan ekstensi, suffix [1], spasi, atau karakter tersembunyi.
    """
    expected_name_norm = normalize_workbook_name(Path(expected_name.strip()).name)
    deadline = time.time() + EXCEL_OPEN_TIMEOUT_SEC
    last_seen_names: list[str] | None = None
    print(f"[WAIT] Menunggu workbook: '{expected_name}' (normalized: '{expected_name_norm}')")
    
    while time.time() < deadline:
        workbooks = get_all_open_workbooks()
        current_names = [str(wb.Name) for wb in workbooks]
        
        if current_names != last_seen_names:
            print(f"[POLL] Workbook Excel terbuka saat ini: {current_names}")
            if current_names:
                normalized_open = [normalize_workbook_name(n) for n in current_names]
                print(f"[POLL] Setelah normalisasi: {normalized_open} | Dicari: '{expected_name_norm}'")
            last_seen_names = current_names
        
        for wb in workbooks:
            if workbook_matches_expected(wb, expected_name_norm):
                print(f"[OK] Workbook ditemukan: '{wb.Name}' cocok dengan DB: '{expected_name}'")
                return wb
        
        remaining = deadline - time.time()
        print(f"[WAIT] Belum ditemukan '{expected_name_norm}', cek lagi dalam {EXCEL_WORKBOOK_POLL_SEC:.0f}s... (sisa {remaining:.0f}s)")
        time.sleep(EXCEL_WORKBOOK_POLL_SEC)
        
    raise TimeoutError(
        f"Workbook '{expected_name}' (normalized: '{expected_name_norm}') tidak ditemukan dalam "
        f"{EXCEL_OPEN_TIMEOUT_SEC:g} detik. Pastikan nama_file di database cocok dengan nama file Excel."
    )


def save_due_job_ids(jobs: list[dict]) -> None:
    DUE_IDS_TXT.parent.mkdir(parents=True, exist_ok=True)
    DUE_IDS_TXT.write_text("".join(f"{job['id']}{os.linesep}" for job in jobs), encoding="utf-8")
    print(f"Snapshot ID automation disimpan ke {DUE_IDS_TXT}: {len(jobs)} ID.")


def save_job_captions(jobs: list[dict]) -> None:
    SCREENSHOT_DIR.mkdir(parents=True, exist_ok=True)
    for automation_number, job in enumerate(jobs, start=1):
        caption_path = SCREENSHOT_DIR / f"{automation_number}_caption.txt"
        caption = str(job.get("caption") or "").strip()
        if caption:
            caption_path.write_text(caption, encoding="utf-8")
            print(f"Caption automation {automation_number} disimpan: {caption_path}")
        elif caption_path.exists():
            caption_path.unlink()
            print(f"Caption automation {automation_number} kosong; file lama dihapus.")


def capture_range_image(
    workbook,
    sheet_name: str,
    cell_range: str,
    file_prefix: str,
    screenshot_number: int,
) -> Path:
    reversed_range = re.fullmatch(r"([A-Za-z]+)(\d+):([A-Za-z]+)(\d+)", cell_range)
    if reversed_range:
        start_column, start_row, end_column, end_row = reversed_range.groups()
        start_index = sum((ord(char) - 64) * 26 ** index for index, char in enumerate(reversed(start_column.upper())))
        end_index = sum((ord(char) - 64) * 26 ** index for index, char in enumerate(reversed(end_column.upper())))
        if start_index > end_index:
            cell_range = f"{end_column}{start_row}:{start_column}{end_row}"
            print(f"Range terbalik dinormalisasi: {sheet_name}!{reversed_range.group(0)} -> {cell_range}")
    worksheet = workbook.Worksheets(sheet_name)
    cell_range_object = worksheet.Range(cell_range)
    output_path = SCREENSHOT_DIR / f"{file_prefix}_{screenshot_number}.png"
    SCREENSHOT_DIR.mkdir(parents=True, exist_ok=True)
    with CLIPBOARD_LOCK:
        workbook.Activate()
        worksheet.Activate()
        cell_range_object.Select()
        cell_range_object.CopyPicture(Appearance=1, Format=2)
        clipboard_deadline = time.time() + CLIPBOARD_IMAGE_TIMEOUT_SEC
        clipboard_image = None
        while time.time() < clipboard_deadline:
            clipboard_image = ImageGrab.grabclipboard()
            if clipboard_image is not None and hasattr(clipboard_image, "save"):
                break
            time.sleep(0.1)
        if clipboard_image is not None and hasattr(clipboard_image, "save"):
            clipboard_image.save(str(output_path), "PNG")
        else:
            chart_object = worksheet.ChartObjects().Add(0, 0, cell_range_object.Width, cell_range_object.Height)
            try:
                chart_object.Chart.Paste()
                chart_object.Chart.Export(str(output_path), "PNG")
            finally:
                chart_object.Delete()
    if not output_path.is_file() or output_path.stat().st_size < 1024:
        raise RuntimeError(f"Screenshot gagal dibuat untuk {sheet_name}!{cell_range}.")
    return output_path


def call_with_retry(func, *args, max_retries=15, delay_sec=2.0, **kwargs):
    """Jalankan fungsi COM berulang kali jika Excel sedang busy (0x800ac472)."""
    for attempt in range(max_retries):
        try:
            return func(*args, **kwargs)
        except Exception as e:
            error_str = str(e).lower()
            if "800ac472" in error_str or "-2146777998" in error_str:
                print(f"[WAIT] Excel sedang sibuk/loading, mencoba lagi dalam {delay_sec}s... ({attempt + 1}/{max_retries})")
                time.sleep(delay_sec)
            else:
                raise
    raise RuntimeError(f"Gagal mengeksekusi aksi: Excel terus sibuk setelah {max_retries} percobaan.")


def disable_background_refresh(workbook: object) -> None:
    """Paksa query workbook menunggu sampai refresh selesai sebelum lanjut."""
    try:
        connections = workbook.Connections
        for index in range(1, connections.Count + 1):
            connection = connections(index)
            for property_name in ("OLEDBConnection", "ODBCConnection"):
                try:
                    setattr(getattr(connection, property_name), "BackgroundQuery", False)
                except Exception:
                    pass
    except Exception as error:
        print(f"[WARN] Tidak semua koneksi Excel bisa diatur sinkron: {error}")

    for worksheet_index in range(1, workbook.Worksheets.Count + 1):
        worksheet = workbook.Worksheets(worksheet_index)
        try:
            query_tables = worksheet.QueryTables
            for index in range(1, query_tables.Count + 1):
                try:
                    query_tables(index).BackgroundQuery = False
                except Exception:
                    pass
        except Exception:
            pass


def refresh_workbook(workbook: object) -> None:
    """Refresh workbook secara blocking dan tunggu kalkulasi Excel selesai."""
    disable_background_refresh(workbook)
    print(f"[REFRESH] Memperbarui data workbook (timeout {EXCEL_REFRESH_TIMEOUT_SEC:g}s)...")
    call_with_retry(workbook.RefreshAll, max_retries=30, delay_sec=2.0)

    deadline = time.time() + EXCEL_REFRESH_TIMEOUT_SEC
    while time.time() < deadline:
        try:
            call_with_retry(workbook.Application.CalculateUntilAsyncQueriesDone, max_retries=5, delay_sec=1.0)
        except Exception as error:
            print(f"[WARN] Menunggu query async Excel: {error}")

        try:
            calculation_state = int(workbook.Application.CalculationState)
        except Exception:
            calculation_state = 0
        if calculation_state == 0:
            print("[REFRESH] Refresh dan kalkulasi workbook selesai.")
            return
        time.sleep(EXCEL_REFRESH_POLL_SEC)

    raise TimeoutError(f"Refresh workbook belum selesai dalam {EXCEL_REFRESH_TIMEOUT_SEC:g} detik.")


def close_workbook(workbook: object) -> None:
    try:
        excel = workbook.Application
        workbook.Close(SaveChanges=False)
        if excel.Workbooks.Count == 0:
            excel.Quit()
        print("Workbook selesai diproses dan Excel ditutup.")
    except Exception as error:
        print(f"Gagal menutup workbook/Excel: {error}")


def capture_job_ranges(workbook, job: dict, automation_number: int, group: dict) -> None:
    # Gunakan group_jid sebagai identitas utama untuk mempermudah script JS mengirim pesan
    group_jid = str(group.get("group_jid") or "").strip() or "tanpa_grup"
    # Ganti karakter non-word menjadi underscore _kecuali_ @ dan . (yang aman dan penting untuk JID)
    safe_jid = re.sub(r"[^\w@.-]", "_", group_jid)
    
    file_prefix = f"{automation_number}_{safe_jid}"
    screenshot_number = 0
    for sheet in job.get("request_sheets") or []:
        sheet_name = str(sheet.get("sheet_name") or "").strip()
        for table in sheet.get("sheet_tables") or []:
            cell_range = str(table.get("cell_range") or "").strip()
            if not sheet_name or not cell_range:
                continue
            screenshot_number += 1
            try:
                screenshot_path = call_with_retry(
                    capture_range_image,
                    workbook,
                    sheet_name,
                    cell_range,
                    file_prefix,
                    screenshot_number,
                )
                print(f"Screenshot dibuat: {screenshot_path}")
            except Exception as error:
                error_msg = str(error).replace('\r', '').replace('\n', ' ')
                print(
                    f"Gagal screenshot {job.get('nama_automation') or job['id']} pada sheet {sheet_name!r}, "
                    f"range {cell_range!r}: {error_msg}"
                )
                # Lanjutkan ke range berikutnya agar tidak membatalkan grup lain



def open_one_sharepoint_in_excel(
    job: dict,
    job_number: int,
    job_count: int,
    chrome: str,
    chrome_profile: str,
) -> None:
    """Buka satu file SharePoint, screenshot, lalu tutup workbook-nya."""
    pythoncom.CoInitialize()
    try:
        expected_name = str(job.get("nama_file") or "").strip()
        automation_label = job.get('nama_automation') or job['id']
        if not expected_name:
            raise ValueError(f"nama_file belum diisi untuk automation {job['id']}.")
        sharepoint_url = str(job["link"]).strip()

        # --- Langkah 1: Pastikan Chrome berjalan (hanya satu instance) ---
        chrome_args = [
            chrome,
            f"--user-data-dir={CHROME_USER_DATA_DIR}",
            f"--profile-directory={chrome_profile}",
            sharepoint_url,
        ]
        if CHROME_HEADLESS:
            chrome_args[1:1] = ["--headless=new", "--disable-gpu", "--no-first-run"]
        else:
            chrome_args.insert(1, "--start-minimized")

        print(f"[{automation_label}] Membuka automation {job_number}/{job_count}...")
        with CHROME_LAUNCH_LOCK:
            if chrome_is_running():
                print(f"[{automation_label}] Chrome sudah berjalan; tidak membuat instance kedua.")
            else:
                subprocess.Popen(chrome_args)
                print(f"[{automation_label}] Chrome diluncurkan.")

        # --- Langkah 2: Beri jeda singkat lalu perintahkan Excel membuka file ---
        time.sleep(EXCEL_OPEN_DELAY_SEC)
        print(f"[{automation_label}] Memerintahkan Excel membuka: {expected_name}")
        os.startfile(f"ms-excel:ofe|u|{sharepoint_url}")

        # --- Langkah 3: Poll hingga workbook ditemukan ---
        workbook = wait_for_workbook(expected_name)

        # --- Langkah 4: Refresh data sebelum screenshot ---
        refresh_workbook(workbook)

        # --- Langkah 5: Screenshot (serialized via lock karena clipboard shared) ---
        print(f"[{automation_label}] File siap! Memulai screenshot...")
        try:
            with SCREENSHOT_WORK_LOCK:
                groups = job.get("wa_groups") or [{"group_jid": "tanpa_grup"}]
                for group in groups:
                    capture_job_ranges(
                        workbook,
                        job,
                        job_number,
                        group,
                    )
                print(f"[{automation_label}] Screenshot selesai untuk: {expected_name}")
        finally:
            close_workbook(workbook)
        
    finally:
        pythoncom.CoUninitialize()


def open_sharepoint_in_excel(jobs: list[dict]) -> None:
    if os.name != "nt":
        raise OSError("Script ini membutuhkan Windows untuk membuka Excel desktop.")
    kill_stale_excel()  # Bersihkan proses Excel stale sebelum mulai
    chrome = find_chrome()
    chrome_profile = find_profile_for_email()
    total_jobs = len(jobs)
    for batch_start in range(0, total_jobs, MAX_WORKERS):
        batch = jobs[batch_start:batch_start + MAX_WORKERS]
        batch_end = batch_start + len(batch)
        print(
            f"Memproses batch {batch_start + 1}-{batch_end} dari {total_jobs} "
            f"(maksimal {MAX_WORKERS} automation paralel)..."
        )

        with ThreadPoolExecutor(max_workers=len(batch)) as executor:
            futures = {
                executor.submit(
                    open_one_sharepoint_in_excel,
                    job,
                    batch_start + index + 1,
                    total_jobs,
                    chrome,
                    chrome_profile,
                ): batch_start + index + 1
                for index, job in enumerate(batch)
            }

            for future in as_completed(futures):
                automation_number = futures[future]
                try:
                    future.result()
                except Exception as error:
                    print(f"Automation {automation_number} gagal: {error}")

        print(f"Batch {batch_start + 1}-{batch_end} selesai.")


def send_screenshots() -> None:
    sender_path = SCRIPT_DIR / "sent.py"
    if not sender_path.is_file():
        raise FileNotFoundError(f"Script pengiriman tidak ditemukan: {sender_path}")
    print(f"Menjalankan pengiriman screenshot: {sender_path}")
    subprocess.run([sys.executable, str(sender_path)], cwd=str(SCRIPT_DIR), check=True)


def update_last_run_dates(jobs: list[dict], run_date: datetime.date) -> None:
    for job in jobs:
        request_id = str(job["id"])
        response = requests.patch(
            f"{SUPABASE_URL}/rest/v1/form_request",
            headers=supabase_headers(),
            params={"id": f"eq.{request_id}"},
            json={"last_run_date": run_date.isoformat()},
            timeout=30,
        )
        response.raise_for_status()
        print(f"last_run_date automation {request_id} diperbarui menjadi {run_date.isoformat()}.")


def main() -> None:
    now = datetime.now()
    jobs = get_active_jobs()
    due_jobs = [job for job in jobs if is_due(job, now) and job.get("link")]
    print(f"Automation aktif: {len(jobs)}; memenuhi jadwal saat ini: {len(due_jobs)}.")
    for job in due_jobs:
        load_job_sheets(job)
        load_job_groups(job)
    save_due_job_ids(due_jobs)
    if due_jobs:
        save_job_captions(due_jobs)
        open_sharepoint_in_excel(due_jobs)
        send_screenshots()
        update_last_run_dates(due_jobs, now.date())
    else:
        print("Tidak ada automation yang perlu dijalankan sekarang.")


if __name__ == "__main__":
    main()

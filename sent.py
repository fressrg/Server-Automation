import json
import os
import re
from collections import defaultdict
from pathlib import Path

import requests
from websocket import WebSocketException, create_connection
from dotenv import load_dotenv

load_dotenv()

SCRIPT_DIR = Path(__file__).resolve().parent
SCREENSHOT_DIR = SCRIPT_DIR / "screenshots"
SUPABASE_URL = os.environ.get("SUPABASE_URL", "").strip().rstrip("/")
SUPABASE_SERVICE_ROLE_KEY = os.environ.get("SUPABASE_SERVICE_ROLE_KEY", "").strip()

# ACCOUNT_ID default (fallback jika grup tidak punya wa_account_id)
DEFAULT_ACCOUNT_ID = os.environ.get("ACCOUNT_ID", "").strip()
WA_WS_URL = os.environ.get("WA_WS_URL", "").strip()
WA_WS_PORT = os.environ.get("WHATSAPP_PORT", "3001").strip()

# Cache ws_url per account_id supaya tidak query Supabase berulang kali dalam 1 run
_ws_url_cache: dict[str, str] = {}


def supabase_headers() -> dict[str, str]:
    if not SUPABASE_URL or not SUPABASE_SERVICE_ROLE_KEY:
        raise ValueError("SUPABASE_URL dan SUPABASE_SERVICE_ROLE_KEY wajib diisi di .env.")
    return {
        "apikey": SUPABASE_SERVICE_ROLE_KEY,
        "Authorization": f"Bearer {SUPABASE_SERVICE_ROLE_KEY}",
    }


def safe_group_key(value: str) -> str:
    return re.sub(r"[^\w@.-]", "_", str(value).strip())


def get_ws_url_for_account(account_id: str) -> str:
    """
    Gunakan WebSocket lokal secara langsung. Port tidak lagi dibaca dari wa_accounts
    karena kolom ws_port memang tidak ada di schema database.
    """
    if account_id in _ws_url_cache:
        return _ws_url_cache[account_id]

    url = WA_WS_URL or f"ws://127.0.0.1:{WA_WS_PORT}"
    _ws_url_cache[account_id] = url
    return url


def receive_command_response(websocket) -> dict:
    """Lewati event status/QR dan ambil balasan command yang memiliki field ok."""
    while True:
        response = json.loads(websocket.recv())
        if "ok" in response:
            return response


def load_groups() -> list[dict]:
    """
    Muat semua grup aktif dari semua akun WA.
    Routing ke akun yang benar ditangani lewat kolom wa_account_id per baris.
    """
    params = {
        "is_active": "eq.true",
        "select": "id,group_jid,group_name,wa_account_id",
        "order": "group_name.asc",
    }
    response = requests.get(
        f"{SUPABASE_URL}/rest/v1/wa_groups",
        headers=supabase_headers(),
        params=params,
        timeout=30,
    )
    response.raise_for_status()
    return response.json()


def find_group(group_key: str, groups: list[dict]) -> dict:
    for group in groups:
        if group_key in {str(group.get("group_jid", "")), str(group.get("group_name", ""))}:
            return group
    normalized_key = safe_group_key(group_key)
    for group in groups:
        if normalized_key in {
            safe_group_key(group.get("group_jid", "")),
            safe_group_key(group.get("group_name", "")),
        }:
            return group
    raise LookupError(f"Grup screenshot '{group_key}' tidak ditemukan di tabel wa_groups.")


def send_message(
    group_jid: str,
    account_id: str,
    image_path: Path | None = None,
    caption: str = "",
) -> None:
    """
    Kirim pesan ke grup via WebSocket bridge yang sesuai dengan akun WhatsApp-nya.
    URL WebSocket diambil dari WA_WS_URL atau WHATSAPP_PORT.
    """
    ws_url = get_ws_url_for_account(account_id)
    payload = {"action": "send", "groupId": group_jid}
    if image_path is not None:
        payload["imagePath"] = str(image_path)
    else:
        payload["caption"] = caption
    websocket = None
    try:
        websocket = create_connection(ws_url, timeout=10)
        websocket.send(json.dumps(payload))
        response = receive_command_response(websocket)
    except (OSError, WebSocketException) as error:
        raise RuntimeError(
            f"WebSocket WhatsApp tidak dapat dihubungi di {ws_url}. "
            "Pastikan node index.js sudah berjalan untuk akun ini."
        ) from error
    finally:
        if websocket is not None:
            websocket.close()
    if not response.get("ok"):
        raise RuntimeError(response.get("error", "Pengiriman WhatsApp gagal."))


def ensure_whatsapp_bridge(account_id: str) -> None:
    """Cek koneksi WebSocket bridge untuk akun tertentu."""
    ws_url = get_ws_url_for_account(account_id)
    websocket = None
    try:
        websocket = create_connection(ws_url, timeout=10)
        websocket.send('{"action":"health"}')
        response = receive_command_response(websocket)
    except (OSError, WebSocketException) as error:
        raise RuntimeError(
            f"WebSocket WhatsApp tidak aktif di {ws_url}. "
            "Jalankan node index.js untuk akun ini terlebih dahulu."
        ) from error
    finally:
        if websocket is not None:
            websocket.close()
    if not response.get("connected"):
        raise RuntimeError(
            f"WebSocket aktif di {ws_url}, tetapi Baileys belum terhubung. "
            "QR hanya ditangani oleh index.js."
        )


def send_screenshots() -> None:
    if not SCREENSHOT_DIR.is_dir():
        print("Folder screenshots tidak ditemukan.")
        return

    pattern = re.compile(r"^(\d+)_(.+)_(\d+)\.png$", re.IGNORECASE)
    jobs: dict[int, dict[str, list[tuple[int, Path]]]] = defaultdict(lambda: defaultdict(list))
    for image_path in SCREENSHOT_DIR.glob("*.png"):
        match = pattern.match(image_path.name)
        if match:
            automation_number = int(match.group(1))
            group_key = match.group(2)
            screenshot_number = int(match.group(3))
            jobs[automation_number][group_key].append((screenshot_number, image_path))

    account_id = DEFAULT_ACCOUNT_ID or "default"
    ensure_whatsapp_bridge(account_id)

    for automation_number in sorted(jobs):
        caption_path = SCREENSHOT_DIR / f"{automation_number}_caption.txt"
        caption = caption_path.read_text(encoding="utf-8").strip() if caption_path.is_file() else ""
        print(f"Automation {automation_number}: {len(jobs[automation_number])} grup.")

        for group_key in sorted(jobs[automation_number]):
            images = jobs[automation_number][group_key]
            group_jid = group_key
            print(f"  Mengirim ke {group_jid} via WebSocket WA...")
            for _, image_path in sorted(images, key=lambda item: item[0]):
                send_message(group_jid, account_id, image_path=image_path)
                print(f"    Terkirim: {image_path.name}")
            if caption:
                send_message(group_jid, account_id, caption=caption)
                print(f"    Caption {automation_number}_caption.txt terkirim.")


if __name__ == "__main__":
    send_screenshots()



#!/usr/bin/env python3
"""
Download CAISO fuelsource.csv for yesterday (ET) and save to a OneDrive-synced folder.

Placeholders to change are clearly marked with ">>> CHANGE_ME <<<" comments.
"""

import os
import sys
import time
import logging
from datetime import datetime, timedelta

import requests
import subprocess

try:
    # Python 3.9+
    from zoneinfo import ZoneInfo
except Exception:
    ZoneInfo = None

# -----------------------
# CONFIG — >>> CHANGE_ME <<<
# -----------------------
# Set this to your OneDrive folder for CAISO data dumps.
ONE_DRIVE_PATH = r"C:\Users\8234545\OneDrive - Standard Chartered Bank\automation\CAISO_data_dumps"  # >>> CHANGE_ME: set your OneDrive path

# Optional: filename template (do not change unless you understand it)
FILENAME_TEMPLATE = "fuelsource_{datekey}.csv"
# Constant filename that Power Automate / Excel will read from
CONSTANT_FILENAME = "fuelsource.csv"

# Number of HTTP download attempts and backoff behavior
RETRIES = 3
RETRY_DELAY = 5  # seconds, multiplied by attempt number

# Timezone used to compute "yesterday" (ET)
TIMEZONE = "America/New_York"  # >>> CHANGE_ME only if you want a different TZ

# Archive folder under ONE_DRIVE_PATH where copies are kept
ARCHIVE_DIR = "archive"

# Log file path (will be created inside ONE_DRIVE_PATH)
LOGFILE = os.path.join(ONE_DRIVE_PATH, "download_caiso.log")
# -----------------------

# Ensure log directory exists before configuring logging
os.makedirs(ONE_DRIVE_PATH, exist_ok=True)
logging.basicConfig(
    filename=LOGFILE,
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
)

def get_zoneinfo(tz_name: str):
    if ZoneInfo:
        try:
            return ZoneInfo(tz_name)
        except Exception:
            logging.exception("ZoneInfo failed for %s", tz_name)
            # Fallback: try dateutil.tz if available
            try:
                from dateutil import tz as dateutil_tz
                t = dateutil_tz.gettz(tz_name)
                if t:
                    logging.info("Falling back to dateutil tz for %s", tz_name)
                    return t
            except Exception:
                logging.debug("dateutil.tz not available or failed")
            return None
    return None


def get_ie_proxy_from_registry() -> dict | None:
    """Try to read IE/WinINet proxy from current user's registry settings.
    Returns a proxies dict compatible with requests, or None.
    """
    try:
        import winreg
    except Exception:
        return None

    try:
        # IE proxy settings live under this key for current user
        key_path = r"Software\Microsoft\Windows\CurrentVersion\Internet Settings"
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path) as key:
            try:
                proxy_enable, _ = winreg.QueryValueEx(key, "ProxyEnable")
            except FileNotFoundError:
                return None
            if not proxy_enable:
                return None
            proxy_server, _ = winreg.QueryValueEx(key, "ProxyServer")
            if not proxy_server:
                return None

            # ProxyServer can contain multiple protocols; take first host:port
            proxy = proxy_server.split(";")[0]
            if "=" in proxy:
                # entries like "http=host:port;https=host:port" -> pick https if present
                parts = dict(p.split("=") for p in proxy_server.split(";") if "=" in p)
                proxy = parts.get("https") or parts.get("http") or list(parts.values())[0]

            if not proxy.startswith(("http://", "https://")):
                proxy = "http://" + proxy

            return {"http": proxy, "https": proxy}
    except Exception:
        logging.exception("Failed to read IE proxy from registry")
        return None

# Resolve proxies at import time so download() can use them and scheduled tasks
PROXIES = get_ie_proxy_from_registry()
if PROXIES:
    try:
        os.environ.setdefault("HTTP_PROXY", PROXIES["http"])
        os.environ.setdefault("HTTPS_PROXY", PROXIES["https"])
        logging.info("Set HTTP(S)_PROXY from IE settings: %s", PROXIES.get("https"))
    except Exception:
        logging.exception("Failed to set proxy env vars")
def yesterday_datekey() -> str:
    tz = get_zoneinfo(TIMEZONE)
    if tz:
        now_et = datetime.now(tz)
    else:
        # Fallback: assume local time is ET — change TIMEZONE or install Python 3.9+
        now_et = datetime.now()
    y = now_et.date() - timedelta(days=1)
    return y.strftime("%Y%m%d")

def download(url: str) -> bytes | None:
    for attempt in range(1, RETRIES + 1):
        try:
            resp = requests.get(url, timeout=30)
            if resp.status_code == 200 and resp.content:
                return resp.content
            logging.warning("Attempt %d: HTTP %s for %s", attempt, resp.status_code, url)
        except Exception:
            logging.exception("Attempt %d exception while downloading %s", attempt, url)
        time.sleep(RETRY_DELAY * attempt)
    return None


def powershell_download_to_bytes(url: str, target_dir: str, datekey: str) -> bytes | None:
    """Fallback: use PowerShell Invoke-WebRequest (WinINet) to download the URL to a temp file,
    then read and return its bytes. This inherits system/IE proxy settings.
    """
    tmp = os.path.join(target_dir, f"pw_{datekey}.tmp")
    # Build a PowerShell inline command that downloads to the tmp file.
    ps_command = (
        "try { Invoke-WebRequest -Uri '" + url + "' -OutFile '" + tmp + "' -UseBasicParsing; exit 0 } "
        "catch { exit 2 }"
    )
    cmd = [
        "powershell",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-Command",
        ps_command,
    ]

    for attempt in range(1, RETRIES + 1):
        try:
            proc = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
            if proc.returncode == 0 and os.path.exists(tmp):
                with open(tmp, "rb") as f:
                    data = f.read()
                try:
                    os.remove(tmp)
                except Exception:
                    logging.debug("Failed to remove tmp file %s", tmp)
                logging.info("PowerShell download succeeded for %s", url)
                return data
            logging.warning("PowerShell attempt %d failed rc=%s stderr=%s", attempt, proc.returncode, proc.stderr.strip())
        except FileNotFoundError:
            # powershell binary not found on PATH
            logging.exception("PowerShell not found on PATH; cannot use fallback")
            break
        except Exception:
            logging.exception("PowerShell attempt %d raised", attempt)

        time.sleep(RETRY_DELAY * attempt)

    return None

def ensure_dir(path: str):
    if not os.path.exists(path):
        os.makedirs(path, exist_ok=True)

def atomic_write(path: str, content: bytes):
    temp_path = path + ".tmp"
    with open(temp_path, "wb") as f:
        f.write(content)
        f.flush()
        os.fsync(f.fileno())
    os.replace(temp_path, path)

def main() -> int:
    ensure_dir(ONE_DRIVE_PATH)
    ensure_dir(os.path.join(ONE_DRIVE_PATH, ARCHIVE_DIR))

    datekey = yesterday_datekey()
    url = f"https://www.caiso.com/outlook/history/{datekey}/fuelsource.csv"
    logging.info("Starting download for %s -> %s", datekey, url)

    content = download(url)
    if not content:
        logging.error("Failed to download %s after %d attempts (requests). Trying PowerShell fallback...", url, RETRIES)
        try:
            content = powershell_download_to_bytes(url, ONE_DRIVE_PATH, datekey)
        except Exception:
            logging.exception("PowerShell fallback raised an exception")

    if not content:
        logging.error("Failed to download %s after %d attempts", url, RETRIES)
        # Create a failure marker file so Power Automate can detect the failure if desired
        marker_path = os.path.join(ONE_DRIVE_PATH, f"download_failed_{datekey}.txt")
        try:
            with open(marker_path, "w", encoding="utf-8") as f:
                f.write(f"download failed for {datekey}\n")
        except Exception:
            logging.exception("Failed to write failure marker %s", marker_path)
        return 1

    fname = FILENAME_TEMPLATE.format(datekey=datekey)
    target = os.path.join(ONE_DRIVE_PATH, fname)
    archive_target = os.path.join(ONE_DRIVE_PATH, ARCHIVE_DIR, fname)
    constant_target = os.path.join(ONE_DRIVE_PATH, CONSTANT_FILENAME)

    try:
        # Overwrite the constant CSV that downstream Power Automate/Excel will read
        atomic_write(constant_target, content)
        # Also save dated file and archive
        atomic_write(target, content)
        atomic_write(archive_target, content)
        logging.info("Saved constant %s, saved %s and archived to %s", constant_target, target, archive_target)
    except Exception:
        logging.exception("Failed to save files to %s", ONE_DRIVE_PATH)
        return 2

    # Remove any previous failure marker for this date if present
    marker = os.path.join(ONE_DRIVE_PATH, f"download_failed_{datekey}.txt")
    try:
        if os.path.exists(marker):
            os.remove(marker)
    except Exception:
        logging.exception("Failed to remove marker %s", marker)

    # Optional: rotate old archives (keep last N days) — uncomment and change KEEP_DAYS if desired
    # KEEP_DAYS = 30  # >>> CHANGE_ME: how many days of archives to keep
    # rotate_archives(os.path.join(ONE_DRIVE_PATH, ARCHIVE_DIR), KEEP_DAYS))``

    return 0

if __name__ == "__main__":
    sys.exit(main())
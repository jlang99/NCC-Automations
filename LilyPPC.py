"""
LilyPPC.py — Lily Power Plant Controller mode monitor.

Captures the Lily SCADA page in Chrome every 15 minutes, locates the PPC
table via template matching, and emails an alert if the control mode differs
from the expected state.

Setup
-----
1. Install dependencies:
       pip install pytesseract pillow opencv-python pygetwindow numpy

2. Tesseract must be installed at the path in TESSERACT_PATH below.
   Download: https://github.com/UB-Mannheim/tesseract/wiki

3. Run:
       python LilyPPC.py

   To start automatically on login, create a Task Scheduler task that runs
   this script at logon with "Run only when user is logged on".
"""

import os
import sys
import json
import time
import logging
import smtplib
from datetime import datetime
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import ctypes

sys.path.insert(0, r"G:\Shared drives\O&M\NCC Automations")
from PythonTools import CREDS, EMAILS

sys.path.insert(0, r"G:\Shared drives\O&M\NCC Automations\RoboCaller")
from robocaller import make_calls

import cv2
import numpy as np
import pytesseract
from PIL import Image, ImageGrab

# ─── Configuration ────────────────────────────────────────────────────────────
# Set the title of the console window
ctypes.windll.kernel32.SetConsoleTitleW("Lily PPC Config Verification")

TESSERACT_PATH = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
TEMPLATE_PATH  = r"G:\Shared drives\O&M\NCC Automations\Lily Tools\Don't Delete\ppc_to_find.png"
SAVE_DIR       = r"G:\Shared drives\O&M\NCC Automations\Lily Tools\PPC Monitor Data"
STATE_FILE     = r"G:\Shared drives\O&M\NCC Automations\Lily Tools\PPC Monitor Data\ppc_state.json"

CHECK_INTERVAL_SEC = 15 * 60   # 15 minutes between checks
WINDOW_TITLE       = 'US LIL REGULATION'  # substring of the Chrome tab title
MATCH_THRESHOLD    = 0.60      # minimum template-match confidence (0–1)

# An alert fires when the detected state differs from these expected values.
EXPECTED_MODE = {
    'Q REACTIVE POWER': 'ENABLED',
    'PF POWER FACTOR':  'DISABLED',
}

# Alert recipients — change to any EMAILS key or add addresses directly.
EMAIL_TO = EMAILS['Administrators + NCC']

# ─── Initialisation ───────────────────────────────────────────────────────────

pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH
os.makedirs(SAVE_DIR, exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s  %(levelname)-8s  %(message)s',
    handlers=[
        logging.FileHandler(os.path.join(SAVE_DIR, 'ppc_monitor.log')),
        logging.StreamHandler(),
    ],
)
log = logging.getLogger(__name__)

# ─── Screen capture ───────────────────────────────────────────────────────────

AHK_ACTIVATE = r"G:\Shared drives\O&M\NCC Automations\Lily Tools\Activate Lily PPC.ahk"

def grab_window():
    """Activate the Lily SCADA window via AHK then take a full-screen screenshot."""
    try:
        os.startfile(AHK_ACTIVATE)
        time.sleep(1.5)
    except Exception as exc:
        log.warning(f"AHK activation failed: {exc}")
    return ImageGrab.grab()

# ─── PPC table location ───────────────────────────────────────────────────────

def find_ppc(img_bgr):
    """Template-match ppc_to_find.png inside the screenshot.

    Returns (loc, template_w, template_h, score).
    loc is None when the match score is below MATCH_THRESHOLD.
    """
    tmpl = cv2.imread(TEMPLATE_PATH)
    if tmpl is None:
        raise FileNotFoundError(f"Template image not found: {TEMPLATE_PATH}")
    th, tw = tmpl.shape[:2]
    result = cv2.matchTemplate(img_bgr, tmpl, cv2.TM_CCOEFF_NORMED)
    _, score, _, loc = cv2.minMaxLoc(result)
    return (loc if score >= MATCH_THRESHOLD else None), tw, th, score

# ─── OCR ──────────────────────────────────────────────────────────────────────

def _classify_color(pixel):
    """Return 'ENABLED', 'DISABLED', or None based on RGB pixel color.
    Green background = ENABLED, orange/red background = DISABLED."""
    r, g, b = pixel[:3]
    if g > r and g > 80:
        return 'ENABLED'
    if r > g and r > 80:
        return 'DISABLED'
    return None


def read_status(img, loc, tw, th):
    """Locate PPC row positions via OCR then sample pixel color at the STATUS
    column to determine ENABLED/DISABLED for each channel."""
    x, y = loc
    region = img.crop((x, y, x + tw, y + th))
    rw, rh = region.size

    # 2× upscale improves Tesseract accuracy on small table text
    big = region.resize((rw * 2, rh * 2), Image.LANCZOS)
    big.save(os.path.join(SAVE_DIR, 'debug_ocr_input.png'))

    data = pytesseract.image_to_data(
        big,
        output_type=pytesseract.Output.DICT,
        config='--psm 6 --oem 3',
    )

    words = []
    for i in range(len(data['text'])):
        conf = int(data['conf'][i])
        raw  = data['text'][i].strip()
        txt  = raw.upper()
        if not raw:
            continue
        cx = data['left'][i] + data['width'][i] // 2
        cy = data['top'][i] + data['height'][i] // 2
        log.debug(f"  OCR word: '{txt}'  conf={conf}  cx={cx}  cy={cy}")
        if conf < 30:
            continue
        words.append((txt, cx, cy))

    log.info(f"OCR words (conf>=30): {[(w[0], w[1], w[2]) for w in words]}")

    # Find STATUS column X from its header word
    status_cx = next((cx for txt, cx, cy in words if 'STATUS' in txt), int(rw * 2 * 0.59))
    log.info(f"STATUS column x={status_cx}")

    # Find row Y positions from OCR labels.
    # Q row: look for 'REACTIVE' (OCR often merges it into QREACTIVEPOWER)
    # PF row: label is often missed entirely — use '1.00' (PF set point) as fallback
    q_cy  = next((cy for txt, cx, cy in words if 'REACTIVE' in txt), None)
    pf_cy = next((cy for txt, cx, cy in words if 'FACTOR'   in txt), None)
    if pf_cy is None:
        pf_cy = next((cy for txt, cx, cy in words if '1.00' in txt or '1,00' in txt), None)
    if pf_cy is None and q_cy is not None:
        # Last resort: PF row is typically ~1 row-height below Q row
        row_h = int(rh * 2 * 0.18)
        pf_cy = q_cy + row_h
        log.warning(f"PF row not found in OCR — estimating y={pf_cy}")

    log.info(f"Row Y positions — Q REACTIVE: {q_cy}, PF POWER FACTOR: {pf_cy}")

    # Sample the STATUS cell pixel color to determine ENABLED / DISABLED
    results = {k: None for k in EXPECTED_MODE}
    for channel, row_cy in [('Q REACTIVE POWER', q_cy), ('PF POWER FACTOR', pf_cy)]:
        if row_cy is None:
            log.warning(f"Could not locate row for {channel}")
            continue
        pixel = big.getpixel((status_cx, row_cy))
        status = _classify_color(pixel)
        log.info(f"{channel} pixel RGB{pixel[:3]} -> {status}")
        results[channel] = status

    return results

# ─── Email alert ──────────────────────────────────────────────────────────────

def send_alert(detected, img_path, ts):
    mismatches = '\n'.join(
        f"  {k}: expected {EXPECTED_MODE[k]}, detected {detected.get(k) or 'NOT FOUND'}"
        for k in EXPECTED_MODE
        if detected.get(k) != EXPECTED_MODE[k]
    )
    body = (
        f"Lily PPC Control Mode Alert\n"
        f"Timestamp: {ts}\n\n"
        f"{mismatches}\n\n"
        f"Screenshot of the PPC table is attached."
    )

    sender = EMAILS['NCC Desk']
    to_str = ', '.join(EMAIL_TO) if isinstance(EMAIL_TO, list) else EMAIL_TO

    msg = MIMEMultipart()
    msg['From']    = sender
    msg['To']      = to_str
    msg['Subject'] = f'[ALERT] Lily PPC Mode Changed — {ts}'
    msg.attach(MIMEText(body))

    if os.path.exists(img_path):
        with open(img_path, 'rb') as f:
            att = MIMEImage(f.read())
            att.add_header('Content-Disposition', 'attachment',
                           filename=os.path.basename(img_path))
            msg.attach(att)

    try:
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(sender, CREDS['remoteMonitoring'])
            server.send_message(msg)
        log.info("Alert email sent.")
    except Exception as exc:
        log.error(f"Failed to send email: {exc}")

# ─── State persistence ────────────────────────────────────────────────────────

def load_state():
    """Return the last saved PPC status dict, or {} on first run."""
    try:
        with open(STATE_FILE) as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return {}


def save_state(detected, ts):
    """Persist the current detected status plus a timestamp to the state file."""
    with open(STATE_FILE, 'w') as f:
        json.dump({**detected, 'timestamp': ts}, f, indent=2)

# ─── Single check ─────────────────────────────────────────────────────────────

def check_once():
    ts = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
    log.info(f"-- PPC check {ts} --")

    img     = grab_window()
    img_bgr = cv2.cvtColor(np.array(img), cv2.COLOR_RGB2BGR)

    loc, tw, th, score = find_ppc(img_bgr)
    if loc is None:
        log.warning(
            f"PPC table not found (best score={score:.2f}). "
            "Make sure the PPC page is visible and not covered by other windows."
        )
        img.save(os.path.join(SAVE_DIR, f'debug_{ts}.png'))
        return

    log.info(f"PPC table matched at {loc} (score={score:.2f})")

    detected = read_status(img, loc, tw, th)
    log.info(f"Detected: {detected}")

    x, y      = loc
    crop_path = os.path.join(SAVE_DIR, f'ppc_{ts}.png')
    img.crop((x, y, x + tw, y + th)).save(crop_path)

    previous = load_state()
    log.info(f"Previous state: {previous}")

    make_calls(previous, detected)  # fires on any change vs previous saved state

    if any(detected.get(k) != v for k, v in EXPECTED_MODE.items()):
        log.warning(f"Mode mismatch — expected {EXPECTED_MODE}, got {detected}")
        send_alert(detected, crop_path, ts)
    else:
        log.info("Mode OK.")

    save_state(detected, ts)

# ─── Cleanup ──────────────────────────────────────────────────────────────────

def cleanup_old_screenshots():
    """Delete .png files in SAVE_DIR that are older than 2 days."""
    cutoff = time.time() - 2 * 24 * 60 * 60
    for fname in os.listdir(SAVE_DIR):
        if not fname.lower().endswith('.png'):
            continue
        fpath = os.path.join(SAVE_DIR, fname)
        if os.path.getmtime(fpath) < cutoff:
            try:
                os.remove(fpath)
                log.info(f"Deleted old screenshot: {fname}")
            except Exception as exc:
                log.warning(f"Could not delete {fname}: {exc}")

# ─── Entry point ──────────────────────────────────────────────────────────────

def main():
    log.info(
        f"Lily PPC Monitor started "
        f"(interval={CHECK_INTERVAL_SEC // 60} min, expected={EXPECTED_MODE})"
    )
    check_once()  # run immediately on start, then on the interval
    while True:
        time.sleep(CHECK_INTERVAL_SEC)
        try:
            check_once()
            cleanup_old_screenshots()
        except Exception as exc:
            log.error(f"Unhandled error: {exc}", exc_info=True)


if __name__ == '__main__':
    main()

"""
Send a WhatsApp message to a list of students from an Excel sheet.

For every student the script:
    1. opens a NEW WhatsApp Web tab with the message pre-filled
    2. waits for the page to load
    3. presses Enter to send
    4. closes that tab (Ctrl+W / Cmd+W)
    5. waits, then repeats for the next student

    
Requirements:
    pip install pandas openpyxl pyautogui

Before running:
    - Log in to WhatsApp Web in your DEFAULT browser.
    - Close any other WhatsApp Web tabs.
    - Don't touch the mouse or keyboard while it runs.
    - Emergency stop: move the mouse to a screen corner (pyautogui failsafe).

Excel format:
    | الاسم       | رقم الجوال     |
    |-------------|----------------|
    | Ahmed Ali   | 0501234567     |
    | Sara Noor   | +966559876543  |
"""

import sys
import time
import platform
import webbrowser
from urllib.parse import quote

import pandas as pd
import pyautogui


# ================= CONFIG =================

import os

EXCEL_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "students.xlsx")

NAME_COLUMN = "الاسم"
NUMBER_COLUMN = "رقم الجوال"

FAILED_REPORT_PATH = "failed_students.xlsx"


# ================= MESSAGE =================



MESSAGE_TEMPLATE = (
    "عزيزي الطالب/ {name}\n\n"

    "نذكرك بموعد الاختبار الفصلي الساعة الخامسة عصرًا.\n\n"

    "*للدخول للاختبار، يجب إحضار التالي:*\n"
    "1- ورقة حضور خاصة بالمؤسسة وموقعة.\n"
    "2- بطاقة دخول الاختبار.\n"
    "3- قلم.\n\n"

    "تنبيه هام:\n"
    "في حال عدم دخول الاختبار، سيتم إضافة 500 ريال إلى الرسوم "
    "بسبب إهمال التقيد بمواعيد الاختبار.\n\n"

    "إدارة معهد الشمال الثقافي العالي للتدريب"
)







# ================= TIMING (seconds) =================

FIRST_PAGE_LOAD = 30          # first tab: WhatsApp Web needs longer to sync
PAGE_LOAD = 18                # later tabs
AFTER_SEND = 5                # let the message actually leave before closing
AFTER_CLOSE = 2               # let the browser finish closing the tab
DELAY_BETWEEN_STUDENTS = 8    # keeps sending pace human-like
START_COUNTDOWN = 10          # time to switch away / get hands off keyboard


# ================= KEYBOARD =================

pyautogui.FAILSAFE = True  # mouse to a corner = abort

CLOSE_TAB_KEYS = ("command", "w") if platform.system() == "Darwin" else ("ctrl", "w")


# ================= LOAD EXCEL =================

def load_students(path: str) -> pd.DataFrame:
    try:
        # Read numbers as text so Excel doesn't drop the leading 0 or add ".0"
        df = pd.read_excel(path, dtype={NUMBER_COLUMN: str})
    except FileNotFoundError:
        sys.exit(f"Excel file not found:\n{path}")

    df.columns = [str(c).strip() for c in df.columns]

    missing = [c for c in (NAME_COLUMN, NUMBER_COLUMN) if c not in df.columns]
    if missing:
        sys.exit(
            f"Excel file is missing required column(s): {missing}\n"
            f"Found columns: {list(df.columns)}"
        )

    df = df.dropna(subset=[NAME_COLUMN, NUMBER_COLUMN])
    df = df[df[NUMBER_COLUMN].astype(str).str.strip() != ""]
    return df.reset_index(drop=True)


# ================= PHONE NUMBER =================

ARABIC_DIGITS = str.maketrans("٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹", "01234567890123456789")


def normalize_number(raw) -> str:
    """Return digits only in international format, e.g. 966501234567."""
    s = str(raw).strip().translate(ARABIC_DIGITS)

    if s.endswith(".0"):  # number stored as float in Excel
        s = s[:-2]

    digits = "".join(ch for ch in s if ch.isdigit())

    if digits.startswith("00"):                         # 00966...
        digits = digits[2:]
    if digits.startswith("05") and len(digits) == 10:   # 05xxxxxxxx
        digits = "966" + digits[1:]
    elif digits.startswith("5") and len(digits) == 9:   # 5xxxxxxxx (Excel dropped the 0)
        digits = "966" + digits

    if digits.startswith("966") and len(digits) != 12:
        raise ValueError(f"invalid Saudi number: {raw}")
    if not 8 <= len(digits) <= 15:
        raise ValueError(f"invalid number: {raw}")

    return digits


# ================= SEND ONE =================

def send_one(number: str, message: str, load_wait: int):
    url = f"https://web.whatsapp.com/send?phone={number}&text={quote(message)}"

    # Open a fresh tab for this student
    webbrowser.open_new_tab(url)
    time.sleep(load_wait)

    # Focus the browser window, then send
    width, height = pyautogui.size()
    pyautogui.click(width // 2, height // 2)
    time.sleep(1)
    pyautogui.press("enter")

    # Give it time to send, then close the tab
    time.sleep(AFTER_SEND)
    pyautogui.hotkey(*CLOSE_TAB_KEYS)
    time.sleep(AFTER_CLOSE)


# ================= SEND ALL =================

def send_messages(df: pd.DataFrame):
    total = len(df)
    sent, failed = [], []

    for pos, row in df.iterrows():
        name = str(row[NAME_COLUMN]).strip()
        raw_number = row[NUMBER_COLUMN]

        print(f"\n[{pos + 1}/{total}] {name} ({raw_number})")

        try:
            number = normalize_number(raw_number)
            message = MESSAGE_TEMPLATE.format(name=name)

            send_one(
                number,
                message,
                FIRST_PAGE_LOAD if pos == 0 else PAGE_LOAD,
            )

            sent.append(name)
            print(f"    -> sent to +{number}, tab closed")

        except pyautogui.FailSafeException:
            print("\nStopped by user (mouse moved to corner).")
            break

        except Exception as e:
            failed.append({NAME_COLUMN: name, NUMBER_COLUMN: raw_number, "error": str(e)})
            print(f"    -> FAILED: {e}")

        if pos < total - 1:
            print(f"    waiting {DELAY_BETWEEN_STUDENTS}s...")
            time.sleep(DELAY_BETWEEN_STUDENTS)

    # ================= SUMMARY =================

    print("\n" + "=" * 50)
    print("SUMMARY".center(50))
    print("=" * 50)
    print(f"Total students    : {total}")
    print(f"Sent              : {len(sent)}")
    print(f"Failed            : {len(failed)}")

    if failed:
        print("\nFailed students:")
        for f in failed:
            print(f"  - {f[NAME_COLUMN]}: {f['error']}")
        pd.DataFrame(failed).to_excel(FAILED_REPORT_PATH, index=False)
        print(f"\nSaved to {FAILED_REPORT_PATH} (you can re-run with it as EXCEL_PATH).")

    print("=" * 50)


# ================= MAIN =================

if __name__ == "__main__":
    students = load_students(EXCEL_PATH)
    print(f"Loaded {len(students)} students from {EXCEL_PATH}")

    print(
        "\nMake sure WhatsApp Web is logged in and no other WhatsApp tab is open."
        "\nDon't touch the keyboard/mouse. Move mouse to a screen corner to abort."
    )
    for s in range(START_COUNTDOWN, 0, -1):
        print(f"Starting in {s}...", end="\r")
        time.sleep(1)
    print()

    send_messages(students)
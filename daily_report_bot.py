import pyautogui
import pyperclip
import time
import re
from datetime import datetime
from pathlib import Path


# ============================================================
# SETTINGS
# ============================================================

pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.2

PROJECT_FOLDER = Path(__file__).resolve().parent

GOOGLE_FINANCE_URL = (
    "https://www.google.com/finance/quote/GOOG:NASDAQ"
)


# ============================================================
# HELPER
# ============================================================

def wait(seconds):
    time.sleep(seconds)


# ============================================================
# STEP 1 - OPEN CHROME
# ============================================================

def open_chrome():

    print("Opening Chrome...")

    pyautogui.hotkey("win", "r")
    wait(1)

    pyautogui.write(
        "chrome",
        interval=0.03
    )

    pyautogui.press("enter")

    wait(4)


# ============================================================
# STEP 2 - OPEN GOOGLE FINANCE
# ============================================================

def open_google_finance():

    print("Opening Google Finance...")

    pyautogui.hotkey("ctrl", "l")

    # Copy the URL into clipboard.
    pyperclip.copy(
        GOOGLE_FINANCE_URL
    )

    pyautogui.hotkey(
        "ctrl",
        "v"
    )

    pyautogui.press("enter")

    # Wait for Google Finance.
    wait(6)


# ============================================================
# STEP 3 - COPY GOOGLE FINANCE PAGE
# ============================================================

def copy_google_finance_page():

    print(
        "Copying Google Finance information..."
    )

    # Close any popup/dialog.
    pyautogui.press("esc")

    wait(0.5)

    # Click on a safe area of the webpage.
    # Avoid the Add to list button.
    pyautogui.click(
        700,
        500
    )

    wait(0.5)

    # Select webpage text.
    pyautogui.hotkey(
        "ctrl",
        "a"
    )

    wait(0.5)

    # Copy webpage text.
    pyautogui.hotkey(
        "ctrl",
        "c"
    )

    wait(1)

    page_text = pyperclip.paste()

    if not page_text:

        raise RuntimeError(
            "Could not copy text from Google Finance."
        )

    print(
        "Google Finance text copied successfully."
    )

    return page_text


# ============================================================
# STEP 4 - EXTRACT GOOG PRICE
# ============================================================

def extract_goog_price(page_text):

    print(
        "Extracting GOOG price..."
    )

    # Convert all whitespace/newlines to spaces.
    normalized_text = re.sub(
        r"\s+",
        " ",
        page_text
    )

    company_name = (
        "Alphabet Inc Class C"
    )

    # Find Alphabet Inc Class C.
    position = normalized_text.find(
        company_name
    )

    if position == -1:

        raise RuntimeError(
            "Could not find "
            "'Alphabet Inc Class C' "
            "in Google Finance page."
        )

    # Search near the company name.
    nearby_text = normalized_text[
        position:
        position + 500
    ]

    # Find a dollar value such as:
    # $343.34
    # $1,234.56

    match = re.search(
        r"\$([0-9,]+\.\d{2})",
        nearby_text
    )

    if not match:

        raise RuntimeError(
            "Could not find GOOG price "
            "on Google Finance."
        )

    # IMPORTANT:
    # This value comes from the copied webpage text.
    stock_price = (
        "$" + match.group(1)
    )

    print(
        f"GOOG price fetched from website: "
        f"{stock_price}"
    )

    return stock_price


# ============================================================
# STEP 5 - OPEN MICROSOFT EXCEL
# ============================================================

def open_excel():

    print(
        "Opening Microsoft Excel..."
    )

    # Close any Chrome popup.
    pyautogui.press("esc")

    wait(0.5)

    # Open Windows Start menu.
    pyautogui.press("win")

    wait(1.5)

    # Search for Excel.
    pyautogui.write(
        "Microsoft Excel",
        interval=0.03
    )

    wait(1.5)

    # Launch Excel.
    pyautogui.press("enter")

    # Give Excel time to launch.
    wait(6)

    print(
        "Excel opened."
    )


# ============================================================
# STEP 6 - CREATE NEW WORKBOOK
# ============================================================

def create_new_excel_workbook():

    print(
        "Creating new Excel workbook..."
    )

    pyautogui.hotkey(
        "ctrl",
        "n"
    )

    wait(3)

    print(
        "New workbook created."
    )


# ============================================================
# STEP 7 - ENTER REPORT INTO EXCEL
# ============================================================

def enter_report_into_excel(
    date_time,
    stock_price,
    comment
):

    print(
        "Entering report into Excel..."
    )

    # Make sure Excel is focused.
    pyautogui.click(
        700,
        500
    )

    wait(0.5)

    # Go to cell A1.
    pyautogui.hotkey(
        "ctrl",
        "home"
    )

    wait(0.5)

    # --------------------------------------------------------
    # HEADER ROW
    # --------------------------------------------------------

    pyautogui.write(
        "Date & Time",
        interval=0.03
    )

    pyautogui.press(
        "tab"
    )

    pyautogui.write(
        "GOOG Price",
        interval=0.03
    )

    pyautogui.press(
        "tab"
    )

    pyautogui.write(
        "Comment",
        interval=0.03
    )

    pyautogui.press(
        "enter"
    )

    # --------------------------------------------------------
    # DATA ROW
    # --------------------------------------------------------

    pyautogui.write(
        date_time,
        interval=0.02
    )

    pyautogui.press(
        "tab"
    )

    pyautogui.write(
        stock_price,
        interval=0.02
    )

    pyautogui.press(
        "tab"
    )

    pyautogui.write(
        comment,
        interval=0.02
    )

    pyautogui.press(
        "enter"
    )

    wait(1)

    print(
        "Report entered successfully."
    )


# ============================================================
# STEP 8 - QUICK FORMATTING
# ============================================================

def quick_format_excel():

    print(
        "Applying quick formatting..."
    )

    # Go to A1.
    pyautogui.hotkey(
        "ctrl",
        "home"
    )

    wait(0.3)

    # Select header row.
    pyautogui.hotkey(
        "shift",
        "space"
    )

    wait(0.2)

    # Bold headers.
    pyautogui.hotkey(
        "ctrl",
        "b"
    )

    wait(0.3)

    # Return to A1.
    pyautogui.hotkey(
        "ctrl",
        "home"
    )

    wait(0.3)

    # Select columns A-C.
    pyautogui.hotkey(
        "ctrl",
        "shift",
        "right"
    )

    wait(0.3)

    # Excel ribbon:
    # Alt + H -> O -> I
    # AutoFit column width.
    pyautogui.hotkey(
        "alt",
        "h"
    )

    wait(0.2)

    pyautogui.press(
        "o"
    )

    wait(0.2)

    pyautogui.press(
        "i"
    )

    wait(1)

    print(
        "Quick formatting completed."
    )


# ============================================================
# STEP 9 - SAVE EXCEL FILE
# ============================================================

def save_excel_file(today):

    filename = (
        f"daily_report_{today}.xlsx"
    )

    file_path = (
        PROJECT_FOLDER / filename
    )

    print()
    print(
        f"Saving Excel file: {filename}"
    )

    # --------------------------------------------------------
    # F12 = Save As in Excel
    # --------------------------------------------------------

    pyautogui.press(
        "f12"
    )

    # Give Save As dialog time to appear.
    wait(3)

    print(
        "Save As dialog opened."
    )

    # --------------------------------------------------------
    # Enter complete file path
    # --------------------------------------------------------

    pyperclip.copy(
        str(file_path)
    )

    pyautogui.hotkey(
        "ctrl",
        "a"
    )

    pyautogui.hotkey(
        "ctrl",
        "v"
    )

    wait(0.5)

    pyautogui.press(
        "enter"
    )

    # Wait for Excel to save.
    wait(4)

    # --------------------------------------------------------
    # Handle possible overwrite confirmation
    # --------------------------------------------------------

    pyautogui.press(
        "enter"
    )

    wait(2)

    # --------------------------------------------------------
    # VERIFY THE FILE REALLY EXISTS
    # --------------------------------------------------------

    if not file_path.exists():

        raise RuntimeError(
            "\nExcel file was NOT created.\n"
            f"Expected:\n{file_path}"
        )

    print(
        "Excel file successfully saved."
    )

    print(
        file_path
    )

    return file_path


# ============================================================
# STEP 10 - TAKE FINAL SCREENSHOT
# ============================================================

def take_screenshot(today):

    screenshot_name = (
        f"screenshot_{today}.png"
    )

    screenshot_path = (
        PROJECT_FOLDER /
        screenshot_name
    )

    print()
    print(
        "Preparing final Excel screenshot..."
    )

    # Make sure Excel has focus.
    pyautogui.click(
        700,
        500
    )

    wait(0.5)

    # Maximize Excel.
    pyautogui.hotkey(
        "alt",
        "space"
    )

    wait(0.3)

    pyautogui.press(
        "x"
    )

    wait(1.5)

    # Take screenshot.
    pyautogui.screenshot(
        str(screenshot_path)
    )

    wait(1)

    # Verify screenshot exists.
    if not screenshot_path.exists():

        raise RuntimeError(
            "\nScreenshot was NOT created.\n"
            f"Expected:\n{screenshot_path}"
        )

    print(
        "Screenshot successfully saved."
    )

    print(
        screenshot_path
    )

    return screenshot_path


# ============================================================
# MAIN PROGRAM
# ============================================================

def main():

    print()
    print("=" * 65)
    print(
        "              DAILY REPORT BOT"
    )
    print("=" * 65)
    print()

    # --------------------------------------------------------
    # GENERATE DATE AND TIME AT RUNTIME
    # --------------------------------------------------------

    now = datetime.now()

    date_time = now.strftime(
        "%Y-%m-%d %H:%M:%S"
    )

    today = now.strftime(
        "%Y-%m-%d"
    )

    print(
        f"Runtime date/time: {date_time}"
    )

    print()

    # --------------------------------------------------------
    # CHROME
    # --------------------------------------------------------

    open_chrome()

    open_google_finance()

    # --------------------------------------------------------
    # FETCH GOOG PRICE FROM WEBSITE
    # --------------------------------------------------------

    page_text = (
        copy_google_finance_page()
    )

    stock_price = (
        extract_goog_price(
            page_text
        )
    )

    # --------------------------------------------------------
    # COMMENT
    # --------------------------------------------------------

    comment = (
        "Google stock is being tracked "
        "for today's report."
    )

    print(
        f"Comment: {comment}"
    )

    print()

    # --------------------------------------------------------
    # EXCEL
    # --------------------------------------------------------

    open_excel()

    create_new_excel_workbook()

    # --------------------------------------------------------
    # ENTER REPORT
    # --------------------------------------------------------

    enter_report_into_excel(
        date_time,
        stock_price,
        comment
    )

    # --------------------------------------------------------
    # QUICK FORMATTING
    # --------------------------------------------------------

    quick_format_excel()

    # --------------------------------------------------------
    # SAVE EXCEL
    # --------------------------------------------------------

    excel_file = (
        save_excel_file(
            today
        )
    )

    # --------------------------------------------------------
    # SCREENSHOT
    # --------------------------------------------------------

    screenshot_file = (
        take_screenshot(
            today
        )
    )

    # --------------------------------------------------------
    # FINAL SUCCESS
    # --------------------------------------------------------

    print()
    print("=" * 65)
    print(
        "       AUTOMATION COMPLETED SUCCESSFULLY"
    )
    print("=" * 65)

    print()
    print(
        "Excel file:"
    )
    print(
        excel_file
    )

    print()
    print(
        "Screenshot:"
    )
    print(
        screenshot_file
    )

    print()
    print(
        "Both output files were verified."
    )

    print()
    print("=" * 65)


# ============================================================
# START PROGRAM
# ============================================================

if __name__ == "__main__":
    main()
import pyautogui
import pyperclip
import time
from datetime import datetime
from pathlib import Path

url = "https://www.accuweather.com/en/in/chennai/206671/hourly-weather-forecast/206671"
# Keep one timestamp for the workbook entry and its filename.
run_time = datetime.now()

# Configure PyAutoGUI's delay and emergency stop behavior.
pyautogui.PAUSE = 0.5
pyautogui.FAILSAFE = True

# Open Chrome and navigate to Chennai's hourly weather forecast.
pyautogui.hotkey("win", "r")
pyautogui.typewrite("chrome")
pyautogui.press("enter")
pyautogui.hotkey("ctrl", "t")
pyautogui.typewrite(url)
pyautogui.press("enter")
time.sleep(2)
# Copy the webpage's selectable text, then read it from the clipboard.
pyautogui.hotkey("ctrl", "a")
pyautogui.hotkey("ctrl", "c")

weather_text = pyperclip.paste().strip()

# Open Excel and create a blank workbook for the report.
pyautogui.hotkey("win", "r")
pyautogui.typewrite("excel")
pyautogui.press("enter")
time.sleep(2)
pyautogui.hotkey('ctrl','n')
time.sleep(2)

# Enter the collection timestamp in the first cell and size the column.
pyautogui.hotkey("ctrl", "home")
pyperclip.copy(run_time.strftime("%Y-%m-%d %H:%M:%S"))
pyautogui.hotkey("ctrl", "v")
pyautogui.hotkey("alt", "h", "o", "i")  # Home > Format > AutoFit Column Width
# Move to the next cell and paste the copied weather-page text.
pyautogui.press("right")
pyperclip.copy(weather_text)

pyautogui.hotkey("ctrl", "v")
pyautogui.hotkey("alt", "h", "o", "i")  # Home > Format > AutoFit Column Width

# Add a title in the next cell, then autofit all populated columns.
pyautogui.press("right")
pyperclip.copy("Next 5 Hours Chennai Weather")
pyautogui.hotkey("ctrl", "v")
pyautogui.hotkey("ctrl", "home")
pyautogui.hotkey("ctrl", "shift", "right")
pyautogui.hotkey("alt", "h", "o", "i")  # Home > Format > AutoFit Column Width
pyautogui.sleep(2)

# Create a timestamped output path and check whether it already exists.
report_path = (
    Path(__file__).resolve().parent
    / f"daily_report_{run_time:%Y-%m-%d_%H-%M-%S}.xlsx"
)
replace_existing = report_path.exists()
# Open Save As and enter the full workbook path.
pyautogui.press("f12")  # Save As
time.sleep(3)
# Focus the Save As dialog's File name field and enter the full .xlsx path.
pyautogui.hotkey("alt", "n")
pyautogui.hotkey("ctrl", "a")
pyperclip.copy(str(report_path))
pyautogui.hotkey("ctrl", "v")
# Save the workbook, confirming replacement if the same file already exists.
print("Pasted the copied Chennai weather page text into a new Excel workbook.")
print("Review the workbook and save it in Excel when you are ready.")
pyautogui.press("enter")
if replace_existing:
    time.sleep(2)
    # Accept Excel's confirmation when replacing this day's report.
    pyautogui.hotkey("alt", "y")
    time.sleep(3)
# Stop with a clear error if Excel did not create the workbook.
for _ in range(60):
    if report_path.is_file():
        break
    time.sleep(0.5)

if not report_path.is_file():
    raise RuntimeError(
        f"Excel did not create the expected .xlsx file: {report_path}. "
        "Check the Save As dialog and selected file type."
    )
print(f"Saved the workbook as {report_path}")

# Capture a screenshot of the completed workbook window.
time.sleep(2)
screenshot_path = report_path.with_name(f"{report_path.stem}_screenshot.png")
pyautogui.screenshot().save(screenshot_path)
print(f"Saved the Excel screenshot as {screenshot_path}")
pyautogui.alert(
        text=f"Process Complete Successfully!",
        title="Chennai weather report complete",
        button="OK",
    )
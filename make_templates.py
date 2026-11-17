"""
make_templates.py - creates the 4 crop images that print_gui.py needs.

Run:  python make_templates.py

For each item it tells you what to put on screen. You press Enter in this window,
switch to the screen (you have 5 seconds), then a screenshot opens: drag a small
rectangle around ONLY the button/box, press Enter (or Space) to save, c to cancel.

Tip: crop tight and include a little of the text/icon, e.g. just the printer icon.
"""
import time
from pathlib import Path

import cv2
import numpy as np
import pyautogui

OUT = Path(__file__).with_name("gui_templates")
OUT.mkdir(exist_ok=True)

ITEMS = [
    ("1_print_icon.png",
     "Open the invoice PDF viewer (after clicking ADIDAS FINANCIAL INVOICE PDF)\n"
     "  so the printer icon is visible in the viewer toolbar."),
    ("2_destination.png",
     "Click the printer icon so Chrome's Print window is open.\n"
     "  The 'Destination' drop-down must be visible (closed)."),
    ("3_save_as_pdf.png",
     "In the Print window, click the Destination drop-down so the list is OPEN.\n"
     "  The 'Save as PDF' row must be visible."),
    ("4_save.png",
     "Select 'Save as PDF' so the blue 'Save' button is visible at the bottom."),
]


def grab_and_crop(filename):
    for s in range(5, 0, -1):
        print(f"  capturing in {s}...", end="\r")
        time.sleep(1)
    shot = pyautogui.screenshot()
    img = cv2.cvtColor(np.array(shot), cv2.COLOR_RGB2BGR)
    h, w = img.shape[:2]
    scale = min(1.0, 1400 / w, 800 / h)
    view = cv2.resize(img, (int(w * scale), int(h * scale))) if scale < 1 else img
    x, y, rw, rh = cv2.selectROI("Drag a box around the item, then press Enter", view, showCrosshair=False)
    cv2.destroyAllWindows()
    if rw == 0 or rh == 0:
        print("  skipped")
        return
    x, y, rw, rh = (int(v / scale) for v in (x, y, rw, rh))
    cv2.imwrite(str(OUT / filename), img[y:y + rh, x:x + rw])
    print(f"  saved {OUT / filename}")


def main():
    for filename, instruction in ITEMS:
        print(f"\n== {filename} ==\n  {instruction}")
        input("  Press Enter here when the screen is ready, then switch to it... ")
        grab_and_crop(filename)
    print("\nDone. Now run the main script with PRINT_MODE=gui in .env")


if __name__ == "__main__":
    main()

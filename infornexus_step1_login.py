"""
Step 1 - Infor Nexus RPA login (Playwright)

Setup (once):
    pip install playwright python-dotenv
    playwright install chromium

.env (same folder, never commit it):
    NEXUS_URL=https://network.infornexus.com
    NEXUS_USER=your_username
    NEXUS_PASS=your_password
    NEXUS_ACCESS_CODE=          # leave empty if not required
    HEADLESS=false

Run:
    python infornexus_step1_login.py

Result: nexus_session.json (cookies/storage) that later steps reuse
so they don't have to log in again.
"""
import os
import sys
from pathlib import Path

from dotenv import load_dotenv
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

load_dotenv(Path(__file__).resolve().parent / ".env")

URL = os.getenv("NEXUS_URL", "https://network.infornexus.com")
USER = os.getenv("NEXUS_USER")
PASSWORD = os.getenv("NEXUS_PASS")
ACCESS_CODE = os.getenv("NEXUS_ACCESS_CODE", "").strip()
HEADLESS = os.getenv("HEADLESS", "false").lower() == "true"
SESSION_FILE = Path(__file__).with_name("nexus_session.json")


def find_login_scope(page, timeout=30_000):
    """Return the page or the iframe that actually holds the Username box."""
    page.wait_for_load_state("domcontentloaded")
    deadline = timeout / 1000
    import time
    t0 = time.time()
    while time.time() - t0 < deadline:
        for scope in [page] + list(page.frames):
            box = scope.get_by_placeholder("Username")
            if box.count() and box.first.is_visible():
                return scope
        page.wait_for_timeout(500)
    page.screenshot(path="debug_login_not_found.png", full_page=True)
    raise RuntimeError("Username box not found (see debug_login_not_found.png)")


def type_into(box, text):
    """Click, clear, then type key by key so the site's JS sees the input."""
    box.click()
    box.fill("")
    box.press_sequentially(text, delay=40)
    if box.input_value() != text:          # fallback if typing was swallowed
        box.fill(text)
    if box.input_value() != text:
        raise RuntimeError("Could not write into field: " + (box.get_attribute("placeholder") or ""))


def login(page):
    if not USER or not PASSWORD:
        raise RuntimeError("NEXUS_USER / NEXUS_PASS are empty - check the .env file next to the scripts")
    page.goto(URL, wait_until="domcontentloaded")
    scope = find_login_scope(page)

    type_into(scope.get_by_placeholder("Username").first, USER)
    type_into(scope.get_by_placeholder("Password").first, PASSWORD)
    if ACCESS_CODE:
        type_into(scope.get_by_placeholder("Access Code (If Required)").first, ACCESS_CODE)

    page.screenshot(path="debug_before_click.png")   # password is masked on screen
    scope.get_by_role("button", name="Log In", exact=True).click()

    # Success = the password box is gone
    try:
        scope.get_by_placeholder("Password").first.wait_for(state="detached", timeout=30_000)
    except PWTimeout:
        page.screenshot(path="debug_login_failed.png", full_page=True)
        err = scope.locator("[class*='error'], [role='alert']").first
        msg = err.inner_text() if err.count() else "Login form still visible"
        raise RuntimeError(f"Login failed: {msg} (see debug_login_failed.png)")

    page.wait_for_load_state("networkidle", timeout=30_000)


def main():
    if not USER or not PASSWORD:
        sys.exit("Missing NEXUS_USER / NEXUS_PASS in .env")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=HEADLESS)
        context = browser.new_context(viewport={"width": 1600, "height": 900})
        page = context.new_page()

        login(page)
        print("Logged in:", page.url)

        context.storage_state(path=str(SESSION_FILE))
        print("Session saved to", SESSION_FILE)

        browser.close()


if __name__ == "__main__":
    main()

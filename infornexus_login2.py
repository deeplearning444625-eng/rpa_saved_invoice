"""
infornexus_login2 - log in to Infor Nexus with the SECOND account and keep its own saved session.

.env (next to this file):
    NEXUS_USER_2=...            <- second account user name
    NEXUS_PASS_2=...            <- second account password
    NEXUS_ACCESS_CODE_2=...     <- only if this account needs an Access Code (optional)
    NEXUS_URL=...               <- optional, same login address as the first account
    HEADLESS=true               <- optional, hide the browser

The session is saved in nexus_session_2.json, so it never overwrites the first account's session.

Run a test login (PowerShell, from the project folder):
    python infornexus_login2.py            <- reuse the saved session, log in only if needed
    python infornexus_login2.py --fresh    <- ignore the saved session and log in again

Use it from another script:
    from infornexus_login2 import open_session_2
    browser, context, page = open_session_2(p)       # p = the sync_playwright() object
"""
import os
import re
import sys
from pathlib import Path

from dotenv import load_dotenv
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

BASE = Path(__file__).parent
load_dotenv(BASE / ".env")

LOGIN_URL = os.getenv("NEXUS_URL", "https://network.infornexus.com/")
SESSION_FILE_2 = BASE / "nexus_session_2.json"
HEADLESS = os.getenv("HEADLESS", "false").strip().lower() == "true"


def get_credentials_2():
    """(user, password, access_code) of the second account, read from the environment / .env."""
    user = os.getenv("NEXUS_USER_2") or os.getenv("NEXUS_USER2")
    password = os.getenv("NEXUS_PASS_2") or os.getenv("NEXUS_PASS2")
    code = os.getenv("NEXUS_ACCESS_CODE_2") or os.getenv("NEXUS_ACCESS_CODE2") or ""
    if not user or not password:
        sys.exit("NEXUS_USER_2 / NEXUS_PASS_2 are missing in .env")
    return user.strip(), password, code.strip()


def is_logged_in(page, timeout=15_000):
    """True when the top menu (APPLICATIONS) is showing."""
    try:
        page.get_by_text(re.compile(r"^\s*applications\s*$", re.I)).first.wait_for(
            state="visible", timeout=timeout)
        return True
    except PWTimeout:
        return False


def login(page, user, password, access_code=""):
    """Fill the Nexus 'Log In' form and wait until the top menu shows."""
    page.goto(LOGIN_URL, wait_until="domcontentloaded")
    if is_logged_in(page, 5_000):                       # already logged in (saved session still valid)
        return

    pass_box = page.locator("input[type=password]:visible").first
    pass_box.wait_for(state="visible", timeout=30_000)
    user_box = page.locator("input[type=text]:visible, input:not([type]):visible").first

    user_box.click()
    user_box.fill(user)
    pass_box.click()
    pass_box.fill(password)
    if user_box.input_value() != user:                  # fallback if fill was swallowed
        user_box.press_sequentially(user, delay=40)

    if access_code:
        code_box = page.get_by_placeholder(re.compile(r"access code", re.I))
        if code_box.count():
            code_box.first.fill(access_code)

    button = page.get_by_role("button", name=re.compile(r"^\s*log\s*in\s*$", re.I))
    if button.count():
        button.first.click()
    else:
        pass_box.press("Enter")

    # optional: the 'session timed out / re-activate' page of the first account's login module
    try:
        from infornexus_step1_login import reactivate_if_needed
        page.wait_for_load_state("domcontentloaded")
        reactivate_if_needed(page)
    except Exception:
        pass

    if not is_logged_in(page, 60_000):
        page.screenshot(path="debug_login2_failed.png", full_page=True)
        raise RuntimeError("Login with account 2 failed - see debug_login2_failed.png "
                           "(wrong user / password, Access Code needed, or the page changed)")


def open_session_2(p, fresh=False):
    """Start the browser with account 2's saved session; log in again if needed.
    Returns (browser, context, page)."""
    browser = p.chromium.launch(headless=HEADLESS, args=["--start-maximized"])
    kwargs = {"no_viewport": True}
    if SESSION_FILE_2.exists() and not fresh:
        kwargs["storage_state"] = str(SESSION_FILE_2)
    context = browser.new_context(**kwargs)
    context.on("page", lambda pg: pg.on("dialog", lambda d: d.accept()))
    page = context.new_page()
    page.on("dialog", lambda d: d.accept())

    page.goto(LOGIN_URL, wait_until="domcontentloaded")
    if not is_logged_in(page, 10_000):
        print("Account 2: no valid saved session - logging in...")
        user, password, code = get_credentials_2()
        try:
            login(page, user, password, code)
        except Exception:
            browser.close()
            raise
    context.storage_state(path=str(SESSION_FILE_2))
    print("Account 2: logged in. Session saved to", SESSION_FILE_2.name)
    return browser, context, page


def main():
    fresh = "--fresh" in sys.argv
    with sync_playwright() as p:
        browser, context, page = open_session_2(p, fresh=fresh)
        print("Page title:", page.title())
        print("URL       :", page.url)
        if os.getenv("NONINTERACTIVE") != "1" and sys.stdin and sys.stdin.isatty():
            try:
                input("Press Enter to close the browser...")
            except EOFError:
                pass
        browser.close()


if __name__ == "__main__":
    main()

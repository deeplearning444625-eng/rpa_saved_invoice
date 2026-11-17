"""
Dispatched_Step - Infor Nexus process (Playwright)
    1  click APPLICATIONS (top menu)
    2  click Dispatches (FULFILLMENT column)
    3  type the PO number into the 'PO Number(s) contains' filter
    4  click Apply and wait for the list to load
    5  click the blue Goods Dispatch Ref Number link in the result row (opens the dispatch)
    6  type the invoice number into the 'Invoice Number' field (skipped if already filled)
    7  click Preview in the left-hand step menu
    8  click Validate (Goods Dispatch - Preview) and wait until it turns into Approve
    9  click Approve (opens the Sign dialog)
    10 click Ok in the Sign dialog (notes left empty)
    (more steps will be added below as they are defined)

Reuses the login / saved session / home-page logic of infornexus_step2_menu.py,
so keep these files in the same folder:
    infornexus_step1_login.py, infornexus_step2_menu.py, dispatched_step.py, .env

Run (PowerShell, from the project folder):
    python dispatched_step.py 4502406661 SC260489     <- PO number, then invoice number
    (the invoice number is only needed when the field on the page is empty)
    (or PO_NUMBER=... and INVOICE_NUMBER=... in .env; PO defaults to the test PO 4502406661)

On a problem it saves debug_dispatched_failed.png and rebuilds the process (back to home, start
again from step 1) up to MAX_ATTEMPTS times (default 3, set in .env). It never restarts after Ok
in the Sign dialog was clicked, so a dispatch is not signed twice. Success: debug_dispatched_done.png.
"""
import os
import re
import sys
import time

from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

import infornexus_step2_menu as s2          # login check, session file, home page
from infornexus_step1_login import reactivate_if_needed


class NoRetry(RuntimeError):
    """An error that restarting cannot fix (several dispatches match, validation failed ...)."""


class AlreadyDone(Exception):
    """The dispatch is already approved in Nexus (finished Preview page): skip it, never re-run it."""


PROGRESS = {"sign_clicked": False, "note": ""}      # sign_clicked: set just before Ok in the Sign dialog (never restart after that); note: why skipped
MAX_ATTEMPTS = int(os.getenv("MAX_ATTEMPTS", "3"))


def dispatches_link(page):
    """The visible 'Dispatches' entry (FULFILLMENT column of the APPLICATIONS menu)."""
    return page.get_by_text("Dispatches", exact=True).locator("visible=true").first


BAD_GATEWAY = re.compile(r"Bad gateway|Error code 50[234]|Gateway Time-?out|Service (Temporarily )?Unavailable", re.I)


def is_bad_gateway(page):
    """True if the page is a Cloudflare / server error page (502 Bad gateway, 503, 504)."""
    try:
        if re.search(r"\b50[234]\b", page.title() or ""):
            return True
        return bool(BAD_GATEWAY.search(page.locator("body").inner_text(timeout=3_000)[:1500]))
    except Exception:
        return False


def reload_if_bad_gateway(page, tries=8, pause_ms=5_000):
    """If a Bad gateway page is showing, wait and refresh until the real page loads.
    Returns the number of refreshes done; raises if it never recovers."""
    n = 0
    while is_bad_gateway(page):
        n += 1
        if n > tries:
            raise RuntimeError("Bad gateway still showing after %d refreshes" % tries)
        print("  Bad gateway - refreshing (%d/%d)..." % (n, tries))
        page.wait_for_timeout(pause_ms)
        try:
            page.reload(wait_until="domcontentloaded")
        except PWTimeout:
            pass
        try:
            page.wait_for_load_state("networkidle", timeout=30_000)
        except PWTimeout:
            pass
        if reactivate_if_needed(page):
            print("  session had timed out - re-activated")
    if n:
        print("  page loaded after %d refresh(es)" % n)
    return n


def step_1_open_applications(page):
    """1 - open the APPLICATIONS mega menu."""
    reactivate_if_needed(page)                 # in case the session timed out meanwhile
    menu = page.get_by_text(re.compile(r"^\s*applications\s*$", re.I)).first
    menu.wait_for(state="visible", timeout=30_000)
    menu.click()
    dispatches_link(page).wait_for(state="visible", timeout=10_000)   # menu is open
    print("1 APPLICATIONS opened")


def step_2_click_dispatches(page):
    """2 - click Dispatches and wait for the page to load."""
    before = page.url
    link = dispatches_link(page)
    link.wait_for(state="visible", timeout=10_000)
    link.click()
    page.wait_for_load_state("domcontentloaded")
    page.wait_for_load_state("networkidle", timeout=30_000)
    if reactivate_if_needed(page):             # 'session timed out' page can appear here
        print("  session had timed out - re-activated")
    reload_if_bad_gateway(page)
    print("2 Dispatches clicked")
    print("  page title:", page.title())
    print("  url       :", page.url)
    if page.url == before:
        print("  (url did not change - check debug_dispatched_done.png to confirm the page opened)")


def step_3_fill_po_number(page, po_no):
    """3 - type the PO number into the 'PO Number(s) contains' filter."""
    page.get_by_text("Please apply filter criteria").first.wait_for(state="visible", timeout=30_000)
    label = page.get_by_text(re.compile(r"^\s*PO Number\(s\)\s*(contains)?\s*$", re.I)).first
    label.wait_for(state="visible", timeout=15_000)
    box = label.locator("xpath=following::input[not(@readonly) and not(@type='hidden')][1]")
    box.click()
    box.press("Control+A")
    box.press("Delete")
    box.press_sequentially(po_no, delay=40)
    if box.input_value() != po_no:                      # fallback if typing was swallowed
        box.fill(po_no)
    if box.input_value() != po_no:
        raise PWTimeout("Could not type the PO number into the filter")
    print("3 PO Number filled:", po_no)


SEARCH_WAIT_MS = 6000                           # extra wait after the search finishes


def step_4_click_apply(page):
    """4 - click Apply and wait until the 'Please apply filter criteria' message is gone."""
    page.get_by_role("button", name="Apply", exact=True).first.click()
    page.get_by_text("Please apply filter criteria").first.wait_for(state="hidden", timeout=30_000)
    try:
        page.wait_for_load_state("networkidle", timeout=30_000)
    except PWTimeout:
        pass
    # the result grid shows "Searching..." while it loads: wait for it to go, then wait 6000 ms more
    try:
        page.get_by_text("Searching...").first.wait_for(state="hidden", timeout=60_000)
    except PWTimeout:
        pass
    page.wait_for_timeout(SEARCH_WAIT_MS)
    print("4 Apply done - list loaded")


REF_LINK = re.compile(r"^\s*\d{8,}\s*$")          # the blue Goods Dispatch Ref Number (digits)


def step_5_open_dispatch(page, po_no):
    """5 - click the Goods Dispatch Ref Number link of the result row.
    Returns (page that shows the dispatch, ref number). Fails if the PO matches more than
    one dispatch, so the wrong one is never opened by guess."""
    page.get_by_role("link", name=REF_LINK).first.wait_for(state="visible", timeout=60_000)
    rows = page.locator("tr", has_text=po_no).get_by_role("link", name=REF_LINK)
    links = rows if rows.count() else page.get_by_role("link", name=REF_LINK)
    n = links.count()
    refs = [links.nth(i).inner_text().strip() for i in range(n)]
    if n > 1:
        raise NoRetry(f"{n} dispatches match PO {po_no}: {', '.join(refs)} - not opening any by guess")
    link, ref = links.first, refs[0]

    context = page.context
    try:
        with context.expect_page(timeout=4_000) as new_page_info:
            link.click()
        target = new_page_info.value               # dispatch opened in a new tab
    except PWTimeout:
        target = page                              # dispatch opened in the same tab
    target.wait_for_load_state("domcontentloaded")
    try:
        target.wait_for_load_state("networkidle", timeout=30_000)
    except PWTimeout:
        pass
    if reactivate_if_needed(target):
        print("  session had timed out - re-activated")
    reload_if_bad_gateway(target)
    print("5 Dispatch opened:", ref)
    print("  url:", target.url)
    return target, ref


def find_invoice_box(page, timeout_ms=30_000):
    """Find the input on the same row as the 'Invoice Number' label (to its right), in any frame."""
    deadline = time.time() + timeout_ms / 1000
    while time.time() < deadline:
        for frame in page.frames:
            try:
                lab = frame.get_by_text(re.compile(r"Invoice\s*Number", re.I)).locator("visible=true").first
                if not lab.count():
                    continue
                lb = lab.bounding_box()
                if not lb:
                    continue
                lcy = lb["y"] + lb["height"] / 2
                inputs = frame.locator("input:visible:not([type=hidden]):not([type=checkbox]):not([readonly])")
                best, best_dx = None, 1e9
                for i in range(inputs.count()):
                    ib = inputs.nth(i).bounding_box()
                    if not ib:
                        continue
                    icy = ib["y"] + ib["height"] / 2
                    dx = ib["x"] - (lb["x"] + lb["width"])
                    if abs(icy - lcy) < 20 and dx > -5 and dx < best_dx:
                        best, best_dx = inputs.nth(i), dx
                if best is not None:
                    return best
            except Exception:
                pass
        page.wait_for_timeout(500)
    return None


def step_6_fill_invoice_number(page, invoice_no):
    """6 - type the invoice number into the 'Invoice Number' field of Goods Dispatch - Summary."""
    reload_if_bad_gateway(page)
    box = find_invoice_box(page)
    if box is None:
        print("  frames checked:")
        for f in page.frames:
            print("   -", f.url)
            try:
                ins = f.locator("input:visible")
                for i in range(min(ins.count(), 40)):
                    el = ins.nth(i)
                    print("      input id=%r name=%r type=%r value=%r" % (
                        el.get_attribute("id"), el.get_attribute("name"),
                        el.get_attribute("type"), el.input_value()))
            except Exception:
                pass
        raise RuntimeError("Invoice Number field not found")
    box.scroll_into_view_if_needed()
    current = box.input_value().strip()
    if current:                                         # already filled -> skip this step
        note = "" if current == invoice_no else " (differs from the given %r - left unchanged)" % invoice_no
        print("6 Invoice Number already filled: %s - skipped%s" % (current, note))
        return
    if not invoice_no:
        raise NoRetry("Invoice Number field is empty and no invoice number was given")
    box.click()
    box.press("Control+A")
    box.press("Delete")
    box.press_sequentially(invoice_no, delay=40)
    if box.input_value() != invoice_no:                 # fallback if typing was swallowed
        box.fill(invoice_no)
    if box.input_value() != invoice_no:
        raise PWTimeout("Could not type the invoice number into the field")
    box.press("Tab")                                    # leave the field so the page registers the value
    print("6 Invoice Number filled:", invoice_no)


def step_7_click_preview(page):
    """7 - click 'Preview' in the left-hand step menu and wait for it to load."""
    link = page.get_by_text("Preview", exact=True).locator("visible=true").first
    link.wait_for(state="visible", timeout=15_000)
    link.click()
    page.wait_for_load_state("domcontentloaded")
    try:
        page.wait_for_load_state("networkidle", timeout=30_000)
    except PWTimeout:
        pass
    if reactivate_if_needed(page):
        print("  session had timed out - re-activated")
    reload_if_bad_gateway(page)
    print("7 Preview opened")
    print("  url:", page.url)


def action_button(page, label):
    """The visible button/link with exactly this label (Validate, Approve, Ok ...)."""
    return (page.get_by_role("button", name=label, exact=True)
            .or_(page.get_by_text(label, exact=True))
            .locator("visible=true").first)


def raise_if_already_done(page):
    """A 'Goods Dispatch - Preview' page with no Validate / Approve / Edit / Cancel button
    (only Download, Excel, Print/PDF, Display Document Versioning) is a dispatch that is already
    approved. Nothing to do: raise AlreadyDone so the robot skips it and does not re-run it."""
    preview = page.get_by_text("Goods Dispatch - Preview").locator("visible=true").first
    summary = page.get_by_text("Goods Dispatch - Summary").locator("visible=true").first
    try:
        preview.or_(summary).first.wait_for(state="visible", timeout=30_000)
    except PWTimeout:
        return                                          # unknown page: let the normal steps deal with it
    if not preview.is_visible():
        return                                          # Summary page = still being prepared
    page.wait_for_timeout(3_000)                        # let every button render before deciding
    if any(action_button(page, b).is_visible() for b in ("Validate", "Approve", "Edit", "Cancel")):
        return
    raise AlreadyDone("already approved in Nexus (Preview page has no Validate / Approve button)")


def wait_page_settled(page):
    try:
        page.wait_for_load_state("networkidle", timeout=30_000)
    except PWTimeout:
        pass
    if reactivate_if_needed(page):
        print("  session had timed out - re-activated")
    reload_if_bad_gateway(page)


def step_8_click_validate(page):
    """8 - click Validate on Goods Dispatch - Preview, then wait until the Approve button shows."""
    raise_if_already_done(page)
    btn = action_button(page, "Validate")
    btn.wait_for(state="visible", timeout=30_000)
    btn.click()
    wait_page_settled(page)
    try:
        action_button(page, "Approve").wait_for(state="visible", timeout=60_000)
    except PWTimeout:
        raise NoRetry("Approve button did not appear after Validate - "
                           "check the validation messages in debug_dispatched_failed.png")
    print("8 Validate done - Approve is available")


def step_9_click_approve(page):
    """9 - click Approve and wait for the Sign dialog."""
    btn = action_button(page, "Approve")
    btn.wait_for(state="visible", timeout=15_000)
    btn.click()
    page.get_by_text("ACKNOWLEDGE SIGNATURE").first.wait_for(state="visible", timeout=30_000)
    print("9 Approve clicked - Sign dialog open")


def step_10_confirm_sign(page):
    """10 - click Ok in the Sign dialog (Notes are optional, left empty)."""
    btn = action_button(page, "Ok")
    btn.wait_for(state="visible", timeout=15_000)
    PROGRESS["sign_clicked"] = True                     # from here on the process is never restarted
    btn.click()
    page.get_by_text("ACKNOWLEDGE SIGNATURE").first.wait_for(state="hidden", timeout=30_000)
    wait_page_settled(page)
    try:                                                # confirmation: Approve is gone once signed
        action_button(page, "Approve").wait_for(state="hidden", timeout=30_000)
        print("10 Signed - Approve button gone")
    except PWTimeout:
        print("10 Ok clicked - Approve button still visible, check debug_dispatched_done.png")


def run_dispatched_step(page, po_no, invoice_no):
    """Runs all steps and returns the page where the process ended."""
    step_1_open_applications(page)
    step_2_click_dispatches(page)
    step_3_fill_po_number(page, po_no)
    step_4_click_apply(page)
    page, ref = step_5_open_dispatch(page, po_no)
    raise_if_already_done(page)                # finished dispatch -> skip, don't re-run
    step_6_fill_invoice_number(page, invoice_no)
    step_7_click_preview(page)
    step_8_click_validate(page)
    step_9_click_approve(page)
    step_10_confirm_sign(page)
    return page
    # next steps go here: step_11_...


def recover_to_home(context, page):
    """Close extra tabs, go back to the Nexus home page (log in again if needed). Returns the page to use."""
    for pg in list(context.pages):
        if pg is not page and not pg.is_closed():
            try:
                pg.close()
            except Exception:
                pass
    if page.is_closed():
        page = context.new_page()
        page.on("dialog", lambda d: d.accept())
    page.goto(s2.STATE["home"], wait_until="domcontentloaded")
    reload_if_bad_gateway(page)
    if not s2.is_logged_in(page):
        from infornexus_step1_login import login
        login(page)
        if not s2.is_logged_in(page, timeout=30_000):
            raise RuntimeError("Could not log in again")
    context.storage_state(path=str(s2.SESSION_FILE))
    return page


def run_with_recovery(context, page, po_no, invoice_no, max_attempts=MAX_ATTEMPTS):
    """Run the whole process; on an error rebuild it (back to home, start again from step 1)
    up to max_attempts times. Returns (exit_code, final_page).
    exit_code: 0 = done, 1 = failed, 2 = Ok in the Sign dialog was clicked but the result is uncertain,
    3 = skipped because the dispatch was already approved."""
    for attempt in range(1, max_attempts + 1):
        print("\n=== Attempt %d/%d ===" % (attempt, max_attempts))
        try:
            final_page = run_dispatched_step(page, po_no, invoice_no)
            final_page.screenshot(path="debug_dispatched_done.png")
            print("\nDispatched_Step finished.")
            return 0, final_page
        except AlreadyDone as e:
            PROGRESS["note"] = str(e)
            print("SKIPPED: %s - not re-running this dispatch." % e)
            try:
                context.pages[-1].screenshot(path="debug_dispatched_skipped.png")
            except Exception:
                pass
            return 3, None
        except Exception as e:
            text = str(e).splitlines()[0][:200] if str(e) else ""
            print("FAILED (attempt %d): %s %s" % (attempt, type(e).__name__, text))
            try:
                last = context.pages[-1]
                last.screenshot(path="debug_dispatched_failed_attempt%d.png" % attempt)
                last.screenshot(path="debug_dispatched_failed.png")
                print("Screenshot: debug_dispatched_failed_attempt%d.png" % attempt)
            except Exception:
                pass
            if PROGRESS["sign_clicked"]:
                print("Ok in the Sign dialog was already clicked - NOT restarting (it could sign twice).")
                print("Please check this dispatch in Nexus.")
                return 2, None
            if isinstance(e, NoRetry):
                print("Restarting cannot fix this - stopping.")
                return 1, None
            if attempt == max_attempts:
                break
            wait = 5 * attempt
            print("Rebuilding: waiting %ds, back to home, starting again from step 1..." % wait)
            time.sleep(wait)
            try:
                page = recover_to_home(context, page)
            except Exception as e2:
                print("  recovery problem:", str(e2).splitlines()[0][:200] if str(e2) else type(e2).__name__)
    print("Gave up after %d attempts." % max_attempts)
    return 1, None


def open_session(p):
    """Start the browser with the saved session; log in again if needed."""
    browser = p.chromium.launch(headless=s2.HEADLESS, args=["--start-maximized"])
    kwargs = {"no_viewport": True}
    if s2.SESSION_FILE.exists():
        kwargs["storage_state"] = str(s2.SESSION_FILE)
    context = browser.new_context(**kwargs)
    context.on("page", lambda pg: pg.on("dialog", lambda d: d.accept()))
    page = context.new_page()
    page.on("dialog", lambda d: d.accept())

    s2.STATE["home"] = s2.load_home_url()
    page.goto(s2.STATE["home"], wait_until="domcontentloaded")
    if not s2.is_logged_in(page):
        print("Saved session not valid - logging in again...")
        from infornexus_step1_login import login
        login(page)
        if not s2.is_logged_in(page, timeout=30_000):
            page.screenshot(path="debug_dispatched_nomenu.png", full_page=True)
            browser.close()
            sys.exit("Logged in but top menu not found - see debug_dispatched_nomenu.png")
        context.storage_state(path=str(s2.SESSION_FILE))
    s2.remember_home(page)
    context.storage_state(path=str(s2.SESSION_FILE))     # keep the saved cookies fresh
    return browser, context, page


PO_NUMBER = (sys.argv[1] if len(sys.argv) > 1 else os.getenv("PO_NUMBER", "4502406661")).strip().strip('"')
INVOICE_NUMBER = (sys.argv[2] if len(sys.argv) > 2 else os.getenv("INVOICE_NUMBER", "")).strip().strip('"')


def main():
    print("PO number:", PO_NUMBER, "| Invoice:", INVOICE_NUMBER)
    exit_code = 0
    with sync_playwright() as p:
        browser, context, page = open_session(p)
        exit_code, _ = run_with_recovery(context, page, PO_NUMBER, INVOICE_NUMBER)
        if os.getenv("NONINTERACTIVE") != "1" and sys.stdin and sys.stdin.isatty():
            try:
                input("Press Enter to close the browser...")
            except EOFError:
                pass
        browser.close()
    sys.exit(exit_code)


if __name__ == "__main__":
    main()

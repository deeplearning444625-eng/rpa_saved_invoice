"""
Orders_Step - Infor Nexus process for the SECOND account (Playwright)
    login  log in with account 2 (infornexus_login2.py, its own saved session)
    1      click APPLICATIONS (top menu)
    2      click Orders (SOURCING column)
    3      type the order number into the 'Order' field (Search Orders)
    4      click Apply and wait for the result list to load
    5      click the folder icon in the first column of the order's row (opens the order)
    6      in 'Plans & Shipments' click the Plan/Shipment number of the row whose Type is 'Packing List'
    7      on the Packing List page click the blue 'Reopen' button
           (add --dry-run to the command to stop just before the click, for testing)
    8      in the 'Confirm Reopen' pop-up click the blue 'Reopen' button (not 'Do NOT Reopen')
    9      type the HAWB number into 'HAWB Number (Sample Only)' (the page is editable after Reopen)
    10     click 'Preview' in the left step menu (Summary > Item Packing > Equipment > Preview)
    11     click 'Validate' (the button then turns into 'Approve')
    12     click 'Approve'   (add --no-approve to stop just before this click)
    13     in the 'Sign' pop-up click 'Ok' (notes are optional and left empty; add --no-sign to stop before this click)
           then it is VERIFIED: the Sign pop-up must close and the Approve button must be gone (else status CHECK;
           set VERIFY_SIGN=0 to switch the check off, VERIFY_TIMEOUT_MS to change the 60 s wait)
    (more steps will be added below as they are defined)

Keep these files in the same folder:
    orders_step.py, infornexus_login2.py, .env (NEXUS_USER_2 / NEXUS_PASS_2)
    (infornexus_step1_login.py is used too, only to re-activate a timed-out session, if present)

Run (PowerShell, from the project folder):
    python orders_step.py 4502408413 6634365323   <- order number to search, then the HAWB number to type
    (or ORDER_NUMBER=... / HAWB_NUMBER=... in .env; with neither it uses the test values 4502408413 / 6634365323)
    python orders_step.py 4502408413 6634365323 --dry-run   <- stops before the Reopen click
    python orders_step.py 4502408413 6634365323 --no-approve <- does everything except the Approve click
    python orders_step.py 4502408413 6634365323 --no-sign    <- clicks Approve, stops at the Sign pop-up before Ok

BATCH (many orders, 1 order = 1 HAWB number):
    python orders_step.py --batch                       <- reads orders_list.csv in the same folder
    python orders_step.py --batch my_orders.xlsx        <- or any .csv / .txt / .xlsx file
    The list has one row per order:   order number, HAWB number   (a header row is fine; .xlsx = columns A and B)
    Each order starts clean from the home page. Every result is added to orders_results.csv
    (DONE / FAILED / CHECK / NO HAWB). Started again, the batch skips orders already DONE (--redo runs them again).
    If an order fails AFTER Approve was clicked it is NOT repeated (status CHECK - look at it by hand).
    The test flags (--dry-run / --no-approve / --no-sign) work in batch mode too; results are then marked TEST.

On a problem it saves debug_orders_failed_attemptN.png and rebuilds the process (back to the
home page, start again from step 1) up to MAX_ATTEMPTS times (default 3, set in .env).
On success it saves debug_orders_done.png.

FIX (step 3): the old locator took the first <input> after 'Search Orders', which on this page
is a type="button" input, so fill() failed ('Input of type "button" cannot be filled').
Now only real text/search inputs are accepted, starting from the 'Order' label.
"""
import csv
import os
import re
import sys
from datetime import datetime

from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

import infornexus_login2 as l2

MAX_ATTEMPTS = int(os.getenv("MAX_ATTEMPTS", "3"))
SEARCH_WAIT_MS = 6000                                  # extra wait after Apply, for the result list
VERIFY_SIGN = os.getenv("VERIFY_SIGN", "1") != "0"     # after Ok: check the approval really went through (0 = off)
VERIFY_TIMEOUT_MS = int(os.getenv("VERIFY_TIMEOUT_MS", "60000"))
_argv = list(sys.argv[1:])
BATCH_FILE = None                                      # --batch [list file]: process many orders one after the other
if "--batch" in _argv:
    _i = _argv.index("--batch")
    if _i + 1 < len(_argv) and not _argv[_i + 1].startswith("--"):
        BATCH_FILE = _argv[_i + 1]
        del _argv[_i:_i + 2]
    else:
        BATCH_FILE = "orders_list.csv"
        del _argv[_i]
_ARGS = [a for a in _argv if not a.startswith("--")]
REDO = "--redo" in sys.argv[1:]                        # batch: also run orders already marked DONE
RESULTS_FILE = os.getenv("RESULTS_FILE", "orders_results.csv")
STATE = {"approved": False}                            # set when Approve was clicked (point of no return)
NO_SIGN = "--no-sign" in sys.argv[1:]                  # stop at the Sign pop-up, before clicking Ok
NO_APPROVE = "--no-approve" in sys.argv[1:]            # stop just before the final Approve click
DRY_RUN = "--dry-run" in sys.argv[1:]                  # only when you type --dry-run (no environment variable)
ORDER_NUMBER = (_ARGS[0] if _ARGS else os.getenv("ORDER_NUMBER", "4502408413")).strip().strip('"')
HAWB_NUMBER = (_ARGS[1] if len(_ARGS) > 1 else os.getenv("HAWB_NUMBER", "6634365323")).strip().strip('"')

# XPath filter: a real, editable text box (never a button / submit / checkbox input)
TEXT_INPUT = ("input[(not(@type) or @type='text' or @type='search') "
              "and not(@readonly) and not(@disabled)]")


# ---------------------------------------------------------------- helpers
def reactivate(page):
    """Re-activate a timed-out session (uses the first login module if it is in the folder)."""
    try:
        from infornexus_step1_login import reactivate_if_needed
    except Exception:
        return False
    try:
        return bool(reactivate_if_needed(page))
    except Exception:
        return False


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
    """If a Bad gateway page is showing, wait and refresh until the real page loads."""
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
        wait_idle(page)
        if reactivate(page):
            print("  session had timed out - re-activated")
    if n:
        print("  page loaded after %d refresh(es)" % n)
    return n


def wait_idle(page, timeout=30_000):
    try:
        page.wait_for_load_state("networkidle", timeout=timeout)
    except PWTimeout:
        pass


# ---------------------------------------------------------------- locators
def applications_menu(page):
    """The APPLICATIONS entry of the top menu."""
    return page.get_by_text(re.compile(r"^\s*applications\s*$", re.I)).locator("visible=true").first


def orders_link(page):
    """The 'Orders' entry under SOURCING in the open APPLICATIONS menu."""
    heading = page.get_by_text(re.compile(r"^\s*sourcing\s*$", re.I)).locator("visible=true").first
    if heading.count():
        link = heading.locator("xpath=following::*[not(*) and normalize-space(.)='Orders'][1]")
        if link.count():
            return link
    return page.get_by_text("Orders", exact=True).locator("visible=true").first


def is_text_box(loc):
    """True if the locator is a real <input> that accepts typing (not a button/submit/checkbox)."""
    try:
        return bool(loc.evaluate(
            "e => e.tagName === 'INPUT' && !['button','submit','reset','checkbox','radio',"
            "'image','file','hidden'].includes((e.type || 'text').toLowerCase())"
        ))
    except Exception:
        return False


def order_row(page, order_no):
    """The result-list row that contains the order number (table row or ARIA grid row)."""
    candidates = [
        page.get_by_role("row").filter(has_text=order_no),
        page.locator("tr").filter(has_text=order_no),
        page.get_by_text(order_no, exact=True).locator("visible=true").first
            .locator("xpath=ancestor::*[self::tr or @role='row'][1]"),
    ]
    for loc in candidates:
        try:
            row = loc.locator("visible=true").last          # last = the data row, not a wrapper
            row.wait_for(state="visible", timeout=5_000)
            return row
        except Exception:
            continue
    raise PWTimeout("Result row for order %s not found" % order_no)


def folder_icon(row):
    """The folder icon in the first column of the row (falls back to the first cell itself)."""
    first_cell = row.locator("td, [role='gridcell'], [role='cell']").locator("visible=true").first
    icon = first_cell.locator("a, button, [role='button'], img, svg, i, span").locator("visible=true").first
    return icon if icon.count() else first_cell


def packing_list_link(page, type_text="Packing List"):
    """The Plan/Shipment number link (first column) of the 'Plans & Shipments' row whose Type is 'Packing List'."""
    heading = page.get_by_text("Plans & Shipments", exact=True).locator("visible=true").first
    heading.wait_for(state="visible", timeout=30_000)
    rows = [
        # row whose Type cell is exactly 'Packing List', searched after the 'Plans & Shipments' heading
        heading.locator("xpath=following::td[normalize-space(.)='%s'][1]/ancestor::tr[1]" % type_text),
        # fallback: any visible element with that exact text anywhere on the page
        page.get_by_text(type_text, exact=True).locator("visible=true").first
            .locator("xpath=ancestor::tr[1]"),
    ]
    for row in rows:
        try:
            row.wait_for(state="visible", timeout=5_000)
            link = row.locator("td").first.locator("a").first
            if not link.count():
                link = row.locator("a").first
            link.wait_for(state="visible", timeout=5_000)
            return link
        except Exception:
            continue
    raise PWTimeout("No '%s' row with a Plan/Shipment link found" % type_text)


def order_field(page):
    """The text box under the 'Order' label inside 'Search Orders' (never a button input)."""
    heading = page.get_by_text("Search Orders", exact=True).locator("visible=true").first
    heading.wait_for(state="visible", timeout=30_000)

    candidates = [
        # 1) the 'Order' label that comes after the 'Search Orders' heading -> next real text input
        heading.locator("xpath=following::*[not(*) and normalize-space(.)='Order'][1]"
                        "/following::" + TEXT_INPUT + "[1]"),
        # 2) first real text input after the 'Search Orders' heading
        heading.locator("xpath=following::" + TEXT_INPUT + "[1]"),
        # 3) accessible label, if the page links label and input
        page.get_by_label("Order", exact=True).first,
    ]
    for i, loc in enumerate(candidates, 1):
        try:
            loc.wait_for(state="visible", timeout=5_000)
            if is_text_box(loc):
                return loc
            print("  Order field candidate %d is not a text box - skipped" % i)
        except Exception:
            continue
    raise PWTimeout("Order input field not found")


# ---------------------------------------------------------------- steps
def step_1_open_applications(page):
    """1 - open the APPLICATIONS mega menu."""
    reactivate(page)                                   # in case the session timed out meanwhile
    reload_if_bad_gateway(page)
    menu = applications_menu(page)
    menu.wait_for(state="visible", timeout=30_000)
    menu.click()
    orders_link(page).wait_for(state="visible", timeout=10_000)       # menu is open
    print("1 APPLICATIONS opened")


def step_2_click_orders(page):
    """2 - click Orders (SOURCING column) and wait for the page to load."""
    before = page.url
    link = orders_link(page)
    link.wait_for(state="visible", timeout=10_000)
    link.click()
    page.wait_for_load_state("domcontentloaded")
    wait_idle(page)
    if reactivate(page):
        print("  session had timed out - re-activated")
    reload_if_bad_gateway(page)
    print("2 Orders clicked")
    print("  page title:", page.title())
    print("  url       :", page.url)
    if page.url == before:
        print("  (url did not change - check debug_orders_done.png to confirm the page opened)")


def step_3_fill_order_number(page, order_no):
    """3 - type the order number into the 'Order' field under 'Search Orders'."""
    box = order_field(page)
    box.click()
    box.press("Control+A")
    box.press("Delete")
    box.press_sequentially(order_no, delay=40)
    if box.input_value() != order_no:                   # fallback if typing was swallowed
        box.fill(order_no)
    if box.input_value() != order_no:
        raise PWTimeout("Could not type the order number into the Order field")
    print("3 Order filled:", order_no)


def step_4_click_apply(page):
    """4 - click Apply and wait for the result list."""
    btn = (page.get_by_role("button", name="Apply", exact=True)
           .or_(page.get_by_text("Apply", exact=True)).locator("visible=true").first)
    btn.wait_for(state="visible", timeout=15_000)
    btn.click()
    wait_idle(page)
    try:                                                # grid text while it loads, if the page shows one
        page.get_by_text("Searching...").first.wait_for(state="hidden", timeout=60_000)
    except PWTimeout:
        pass
    page.wait_for_timeout(SEARCH_WAIT_MS)
    reload_if_bad_gateway(page)
    print("4 Apply done - list loaded")


def step_5_click_folder_icon(page, order_no):
    """5 - click the folder icon in the first column of the order's row.
    Returns the page to continue on (a new tab if the click opened one, else the same page)."""
    row = order_row(page, order_no)
    icon = folder_icon(row)
    icon.wait_for(state="visible", timeout=10_000)
    before = page.url
    try:                                                # the click may open a new tab / popup
        with page.context.expect_page(timeout=5_000) as new_page_info:
            icon.click()
        target = new_page_info.value
        target.wait_for_load_state("domcontentloaded")
        print("  (opened in a new tab)")
    except PWTimeout:                                   # no new tab -> same page (or a panel) updated
        target = page
    wait_idle(target)
    if reactivate(target):
        print("  session had timed out - re-activated")
    reload_if_bad_gateway(target)
    print("5 Folder icon clicked")
    print("  page title:", target.title())
    print("  url       :", target.url)
    if target is page and target.url == before:
        print("  (url did not change - check debug_orders_done.png to see what opened)")
    return target


def step_6_open_packing_list(page):
    """6 - in 'Plans & Shipments' click the number of the 'Packing List' row.
    Returns the page to continue on (a new tab if the click opened one, else the same page)."""
    link = packing_list_link(page)
    shipment_no = (link.inner_text() or "").strip()
    print("  Packing List shipment:", shipment_no)
    before = page.url
    try:                                                # the click may open a new tab / popup
        with page.context.expect_page(timeout=5_000) as new_page_info:
            link.click()
        target = new_page_info.value
        target.wait_for_load_state("domcontentloaded")
        print("  (opened in a new tab)")
    except PWTimeout:                                   # no new tab -> same page
        target = page
    wait_idle(target)
    if reactivate(target):
        print("  session had timed out - re-activated")
    reload_if_bad_gateway(target)
    print("6 Packing List opened:", shipment_no)
    print("  page title:", target.title())
    print("  url       :", target.url)
    if target is page and target.url == before:
        print("  (url did not change - check debug_orders_done.png to see what opened)")
    return target


def step_7_click_reopen(page):
    """7 - click the 'Reopen' button on the Packing List page.
    Returns 'clicked', 'dry' (--dry-run, nothing clicked) or 'already' (no Reopen button because an earlier
    attempt already reopened it and the HAWB field is editable)."""
    btn = (page.get_by_role("button", name="Reopen", exact=True)
           .or_(page.get_by_text("Reopen", exact=True)).locator("visible=true").first)
    try:
        btn.wait_for(state="visible", timeout=30_000)
    except PWTimeout:
        try:
            hawb_field(page, timeout_ms=3_000)
        except PWTimeout:
            raise PWTimeout("Reopen button not found and the page is not editable")
        print("7 Reopen button not shown but the page is already editable - already reopened, skipping 7-8")
        return "already"
    if DRY_RUN:
        print("7 --dry-run - Reopen button found, NOT clicked (steps 8-9 skipped)")
        return "dry"
    btn.click()
    wait_idle(page)
    if reactivate(page):
        print("  session had timed out - re-activated")
    reload_if_bad_gateway(page)
    print("7 Reopen clicked")
    return "clicked"


CONFIRM_TITLE = re.compile(r"Confirm\s+Reopen", re.I)


def find_confirm_popup(page, timeout_ms=30_000, title=CONFIRM_TITLE):
    """Wait for a pop-up whose title/heading text matches `title` (default: 'Confirm Reopen'). It may live in the main page OR in an iframe, so every frame is checked.
    Returns (frame, heading_locator)."""
    waited = 0
    while waited <= timeout_ms:
        for fr in page.frames:
            try:
                heading = fr.get_by_text(title).locator("visible=true").first
                if heading.count():
                    return fr, heading
            except Exception:
                continue
        page.wait_for_timeout(500)
        waited += 500
    print("  frames on the page:")
    for fr in page.frames:
        print("   -", fr.url[:150])
    raise PWTimeout("pop-up '%s' not found in the page or any iframe" % title.pattern)


def confirm_reopen_button(page):
    """The blue 'Reopen' button inside the 'Confirm Reopen' pop-up (never 'Do NOT Reopen', never the page button).
    Returns (frame, heading, button, how_found)."""
    fr, heading = find_confirm_popup(page)
    where = "main page" if fr is page.main_frame else "iframe " + fr.url[:80]
    print("  pop-up found in:", where)
    # the pop-up box = nearest ancestor of the heading that also holds the 'Do NOT Reopen' button
    box = heading.locator("xpath=ancestor::*[.//*[normalize-space(.)='Do NOT Reopen'] "
                          "or .//input[@value='Do NOT Reopen']][1]")
    candidates = [
        ("pop-up box / button role", box.get_by_role("button", name="Reopen", exact=True)),
        ("pop-up box / exact text",  box.get_by_text("Reopen", exact=True).locator("visible=true").first),
        ("pop-up box / input value", box.locator("input[value='Reopen']")),
        ("frame dialog role",        fr.get_by_role("dialog").get_by_role("button", name="Reopen", exact=True)),
        ("frame last visible Reopen", fr.get_by_role("button", name="Reopen", exact=True).locator("visible=true").last),
        ("frame last exact text",    fr.get_by_text("Reopen", exact=True).locator("visible=true").last),
    ]
    for name, loc in candidates:
        try:
            loc.first.wait_for(state="visible", timeout=4_000)
            return fr, heading, loc.first, name
        except Exception:
            continue
    raise PWTimeout("'Reopen' button not found in the Confirm Reopen pop-up")


def _mouse_click(page, loc):
    """Real mouse click at the centre of the element (for pages that ignore synthetic clicks)."""
    loc.scroll_into_view_if_needed()
    box = loc.bounding_box()
    if not box:
        raise PWTimeout("button has no position")
    x, y = box["x"] + box["width"] / 2, box["y"] + box["height"] / 2
    page.mouse.move(x, y)
    page.wait_for_timeout(200)
    page.mouse.click(x, y, delay=120)


def popup_gone(fr):
    """True if the pop-up frame/window has been closed or removed."""
    try:
        return fr.is_detached() or fr.page.is_closed()
    except Exception:
        return True


def click_confirm(page, fr, heading, btn):
    """Try several ways to click the pop-up button; stop as soon as the pop-up closes (never double-click).
    This button runs 'submitUserAction(...); window.close()', so the pop-up closing during the click is normal."""
    strategies = [
        ("normal click",                lambda: btn.click(timeout=5_000)),
        ("mouse click on button centre", lambda: _mouse_click(page, btn)),
        ("JavaScript click",            lambda: btn.evaluate("e => e.click()")),
        ("Enter key",                   lambda: (btn.focus(), btn.press("Enter"))),
    ]
    for name, action in strategies:
        try:
            action()
        except Exception as e:
            page.wait_for_timeout(1_000)
            if popup_gone(fr):                          # the click worked: the pop-up closed itself
                print("  pop-up closed after: %s (the click closed it)" % name)
                return True
            print("  %s failed: %s" % (name, str(e).splitlines()[0][:120] if str(e) else type(e).__name__))
            continue
        try:                                            # give the site time to react before trying the next way
            heading.wait_for(state="hidden", timeout=10_000)
            print("  pop-up closed after: %s" % name)
            return True
        except PWTimeout:
            print("  pop-up still showing after: %s" % name)
        except Exception:                               # frame detached / page navigated = pop-up is gone
            print("  pop-up closed after: %s (frame gone)" % name)
            return True
    return False


def step_8_confirm_reopen(page):
    """8 - confirm the 'Reopen' pop-up."""
    fr, heading, btn, how = confirm_reopen_button(page)
    print("  button found via:", how)
    try:
        print("  button html:", btn.evaluate("e => e.outerHTML.slice(0, 200)"))
    except Exception:
        pass
    try:                                                # Goods Dispatch number(s) listed in the pop-up
        linked = heading.locator("xpath=following::a[1]").inner_text(timeout=3_000).strip()
        print("  Goods Dispatch affected:", linked)
    except Exception:
        pass
    if not click_confirm(page, fr, heading, btn):
        raise RuntimeError("Confirm Reopen pop-up did not close after all click methods")
    wait_idle(page)
    if reactivate(page):
        print("  session had timed out - re-activated")
    reload_if_bad_gateway(page)
    print("8 Reopen confirmed")
    print("  page title:", page.title())
    print("  url       :", page.url)


def hawb_field(page, timeout_ms=30_000):
    """The text box next to the 'HAWB Number (Sample Only)' label (main page or any iframe; never a button input)."""
    waited = 0
    while waited <= timeout_ms:
        for fr in page.frames:
            try:
                label = fr.get_by_text(re.compile(r"HAWB\s+Number", re.I)).locator("visible=true").first
                if not label.count():
                    continue
                for loc in (label.locator("xpath=following::" + TEXT_INPUT + "[1]"),
                            fr.get_by_label(re.compile(r"HAWB\s+Number", re.I)).first):
                    try:
                        loc.wait_for(state="visible", timeout=3_000)
                        if is_text_box(loc):
                            return loc
                    except Exception:
                        continue
            except Exception:
                continue
        page.wait_for_timeout(500)
        waited += 500
    raise PWTimeout("HAWB Number input field not found")


def step_9_fill_hawb(page, hawb_no):
    """9 - type the HAWB number into 'HAWB Number (Sample Only)'."""
    box = hawb_field(page)
    box.scroll_into_view_if_needed()
    box.click()
    box.press("Control+A")
    box.press("Delete")
    box.press_sequentially(hawb_no, delay=40)
    if box.input_value() != hawb_no:                    # fallback if typing was swallowed
        box.fill(hawb_no)
    if box.input_value() != hawb_no:
        raise PWTimeout("Could not type the HAWB number into the HAWB field")
    box.press("Tab")                                    # leave the field so the page registers the value
    print("9 HAWB filled:", hawb_no)


def preview_nav(page, timeout_ms=30_000):
    """The 'Preview' entry of the left step menu (main page or any iframe)."""
    waited = 0
    while waited <= timeout_ms:
        for fr in page.frames:
            try:
                loc = fr.get_by_text(re.compile(r"^\s*Preview\s*$")).locator("visible=true").first
                if loc.count():
                    return loc
            except Exception:
                continue
        page.wait_for_timeout(500)
        waited += 500
    raise PWTimeout("'Preview' entry of the left menu not found")


def step_10_click_preview(page):
    """10 - click 'Preview' in the left step menu."""
    nav = preview_nav(page)
    nav.scroll_into_view_if_needed()
    before = page.url
    nav.click()
    page.wait_for_load_state("domcontentloaded")
    wait_idle(page)
    if reactivate(page):
        print("  session had timed out - re-activated")
    reload_if_bad_gateway(page)
    print("10 Preview clicked")
    print("  page title:", page.title())
    print("  url       :", page.url)
    if page.url == before:
        print("  (url did not change - check debug_orders_done.png to see what opened)")


def find_button(page, name, timeout_ms=30_000):
    """A visible button/link whose text is exactly `name`, in the main page or any iframe."""
    waited = 0
    while waited <= timeout_ms:
        for fr in page.frames:
            try:
                loc = (fr.get_by_role("button", name=name, exact=True)
                       .or_(fr.get_by_text(name, exact=True))).locator("visible=true").first
                if loc.count():
                    return loc
            except Exception:
                continue
        page.wait_for_timeout(500)
        waited += 500
    raise PWTimeout("'%s' button not found" % name)


def step_11_validate(page):
    """11 - click 'Validate'; the button turns into 'Approve' when the validation passed."""
    btn = find_button(page, "Validate")
    btn.scroll_into_view_if_needed()
    btn.click()
    wait_idle(page)
    if reactivate(page):
        print("  session had timed out - re-activated")
    reload_if_bad_gateway(page)
    print("11 Validate clicked")
    try:
        find_button(page, "Approve", timeout_ms=60_000)
    except PWTimeout:
        raise PWTimeout("'Approve' did not appear after Validate - the validation probably found errors "
                        "(see the failed-attempt screenshot)")
    print("  validation passed - 'Approve' button is showing")


def step_12_approve(page):
    """12 - click 'Approve'. Returns True if clicked (False with --no-approve). A 'Sign' pop-up follows."""
    btn = find_button(page, "Approve")
    if NO_APPROVE:
        print("12 --no-approve - Approve button found, NOT clicked")
        return False
    btn.scroll_into_view_if_needed()
    btn.click()
    STATE["approved"] = True                            # from here on a failed order is never repeated
    wait_idle(page)
    if reactivate(page):
        print("  session had timed out - re-activated")
    print("12 Approve clicked")
    return True


SIGN_TITLE = re.compile(r"ACKNOWLEDGE\s+SIGNATURE", re.I)
OK_NAME = re.compile(r"^\s*ok\s*$", re.I)


def step_13_sign_ok(page):
    """13 - in the 'Sign' pop-up click 'Ok' (the optional Notes box stays empty)."""
    fr, heading = find_confirm_popup(page, title=SIGN_TITLE)
    where = "main page" if fr is page.main_frame else "iframe " + fr.url[:80]
    print("  Sign pop-up found in:", where)
    box = heading.locator("xpath=ancestor::*[.//*[normalize-space(.)='Cancel'] or .//input[@value='Cancel']][1]")
    candidates = [
        ("pop-up box / button role", box.get_by_role("button", name=OK_NAME)),
        ("pop-up box / exact text",  box.get_by_text(OK_NAME).locator("visible=true").first),
        ("pop-up box / input value", box.locator("input[value='Ok' i]")),
        ("frame button role",        fr.get_by_role("button", name=OK_NAME).locator("visible=true").last),
        ("frame exact text",         fr.get_by_text(OK_NAME).locator("visible=true").last),
    ]
    btn = how = None
    for name, loc in candidates:
        try:
            loc.first.wait_for(state="visible", timeout=4_000)
            btn, how = loc.first, name
            break
        except Exception:
            continue
    if btn is None:
        raise PWTimeout("'Ok' button not found in the Sign pop-up")
    print("  Ok button found via:", how)
    try:
        print("  button html:", btn.evaluate("e => e.outerHTML.slice(0, 200)"))
    except Exception:
        pass
    if NO_SIGN:
        print("13 --no-sign - Sign pop-up found, Ok NOT clicked")
        return
    if not click_confirm(page, fr, heading, btn):
        raise RuntimeError("Sign pop-up did not close after all click methods")
    wait_idle(page)
    if reactivate(page):
        print("  session had timed out - re-activated")
    reload_if_bad_gateway(page)
    print("13 Ok clicked")
    print("  page title:", page.title())
    print("  url       :", page.url)
    if VERIFY_SIGN:
        verify_signed(page)
        print("  verified: Sign pop-up closed and the Approve button is gone - the approval went through")
    else:
        print("  (VERIFY_SIGN=0 - the result of the signature was not checked)")


def _visible_anywhere(page, make_locator):
    """True if make_locator(frame) finds a visible element in the main page or any iframe."""
    for fr in page.frames:
        try:
            if make_locator(fr).locator("visible=true").count():
                return True
        except Exception:
            continue
    return False


def verify_signed(page, timeout_ms=VERIFY_TIMEOUT_MS):
    """After 'Ok': wait until the Sign pop-up is gone AND the Approve button no longer shows (two checks in a row).
    If that does not happen the order is NOT reported as DONE: it is raised here and, because Approve was
    already clicked, the order ends as CHECK (look at it by hand in Nexus)."""
    waited, stable = 0, 0
    while waited <= timeout_ms:
        sign_showing = _visible_anywhere(page, lambda f: f.get_by_text(SIGN_TITLE))
        approve_showing = _visible_anywhere(page, lambda f: f.get_by_role("button", name="Approve", exact=True)
                                            .or_(f.get_by_text("Approve", exact=True)))
        stable = stable + 1 if not (sign_showing or approve_showing) else 0
        if stable >= 2:
            return
        page.wait_for_timeout(1_000)
        waited += 1_000
    raise PWTimeout("after Ok the %s still showing after %ds - signature not confirmed"
                    % ("Sign pop-up is" if sign_showing else "Approve button is", timeout_ms // 1000))


def run_orders_step(page, order_no, hawb_no):
    """Runs all steps and returns the page where the process ended."""
    step_1_open_applications(page)
    step_2_click_orders(page)
    step_3_fill_order_number(page, order_no)
    step_4_click_apply(page)
    page = step_5_click_folder_icon(page, order_no)
    page = step_6_open_packing_list(page)
    state = step_7_click_reopen(page)
    if state == "clicked":
        step_8_confirm_reopen(page)
    if state in ("clicked", "already"):
        step_9_fill_hawb(page, hawb_no)
        step_10_click_preview(page)
        step_11_validate(page)
        if step_12_approve(page):                       # False with --no-approve
            step_13_sign_ok(page)
    return page
    # next steps go here: step_14_...


# ---------------------------------------------------------------- run with rebuild
def recover_to_home(context, page, home_url):
    """Close extra tabs, go back to the home page (log in again if needed). Returns the page to use."""
    for pg in list(context.pages):
        if pg is not page and not pg.is_closed():
            try:
                pg.close()
            except Exception:
                pass
    if page.is_closed():
        page = context.new_page()
        page.on("dialog", lambda d: d.accept())
    page.goto(home_url, wait_until="domcontentloaded")
    reload_if_bad_gateway(page)
    if not l2.is_logged_in(page, 10_000):
        user, password, code = l2.get_credentials_2()
        l2.login(page, user, password, code)
    context.storage_state(path=str(l2.SESSION_FILE_2))
    return page


def pick_page(context):
    """Any open tab of the browser (or a new one if every tab was closed)."""
    for pg in context.pages:
        if not pg.is_closed():
            return pg
    pg = context.new_page()
    pg.on("dialog", lambda d: d.accept())
    return pg


def run_with_recovery(context, page, home_url, order_no, hawb_no, max_attempts=MAX_ATTEMPTS):
    """Run the process for ONE order; on an error rebuild it and start again.
    Returns (code, note): 0 = done, 1 = failed, 2 = failed AFTER Approve was clicked (not repeated, check by hand)."""
    STATE["approved"] = False
    tag = "_" + order_no if BATCH_FILE else ""          # batch mode: one screenshot set per order
    last = ""
    for attempt in range(1, max_attempts + 1):
        print("\n=== Attempt %d/%d ===" % (attempt, max_attempts))
        try:
            final_page = run_orders_step(page, order_no, hawb_no)
            final_page.screenshot(path="debug_orders_done%s.png" % tag)
            print("\nOrders_Step finished.")
            return 0, ""
        except Exception as e:
            text = str(e).splitlines()[0][:200] if str(e) else ""
            last = ("%s %s" % (type(e).__name__, text)).strip()
            print("FAILED (attempt %d): %s" % (attempt, last))
            shot = "debug_orders_failed%s_attempt%d.png" % (tag, attempt)
            try:
                context.pages[-1].screenshot(path=shot)
                print("Screenshot:", shot)
            except Exception:
                pass
            if STATE["approved"]:                       # point of no return: never click Approve/Sign twice
                print("Approve was already clicked - NOT starting again. Check this order by hand.")
                return 2, "failed after Approve: " + last
            if attempt == max_attempts:
                break
            wait = 5 * attempt
            print("Rebuilding: waiting %ds, back to home, starting again from step 1..." % wait)
            page.wait_for_timeout(wait * 1000)
            try:
                page = recover_to_home(context, page, home_url)
            except Exception as e2:
                print("  recovery problem:", str(e2).splitlines()[0][:200] if str(e2) else type(e2).__name__)
    print("Gave up after %d attempts." % max_attempts)
    return 1, last


# ---------------------------------------------------------------- batch: order list <-> HAWB list, one to one
def _clean(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():         # Excel numbers come as 4502408413.0
        v = int(v)
    return str(v).strip().strip('"').strip("'")


def read_pairs(path):
    """Read the list file: one order + its HAWB number per row (1 order = 1 HAWB).
    .csv / .txt  -> 'order,hawb' (comma, semicolon, tab or space);
    .xlsx -> first sheet: the columns whose header contains 'Order' and 'HAWB', else column A and column B.
    A header row, empty lines and lines starting with # are ignored."""
    if not os.path.exists(path):
        raise SystemExit("List file not found: %s" % path)
    rows = []
    if path.lower().endswith((".xlsx", ".xlsm")):
        try:
            import openpyxl
        except ImportError:
            raise SystemExit("Reading .xlsx needs:  pip install openpyxl   (or save the list as .csv)")
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        rows_all = [[_clean(c) for c in r] for r in wb.worksheets[0].iter_rows(values_only=True)]
        oi, hi = 0, 1                                   # default: column A = order, column B = HAWB
        for cells in rows_all:                          # a header row with 'Order...' and 'HAWB...' columns wins
            if not any(cells):
                continue
            o_idx = [i for i, c in enumerate(cells) if re.search(r"order", c, re.I)]
            h_idx = [i for i, c in enumerate(cells) if re.search(r"hawb|house\s*awb", c, re.I)]
            if o_idx and h_idx:
                oi, hi = o_idx[0], h_idx[0]
            break
        for cells in rows_all:
            rows.append([cells[oi] if oi < len(cells) else "", cells[hi] if hi < len(cells) else ""])
    else:
        with open(path, encoding="utf-8-sig") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#"):
                    continue
                rows.append([_clean(c) for c in re.split(r"[,;\t ]+", line)[:2]])
    pairs, seen = [], set()
    for cells in rows:
        order = cells[0] if cells else ""
        hawb = cells[1] if len(cells) > 1 else ""
        if not re.search(r"\d", order):                 # header row / blank row
            continue
        if order in seen:
            print("  duplicate order in the list, ignored:", order)
            continue
        seen.add(order)
        pairs.append((order, hawb))
    return pairs


def log_result(order, hawb, status, note=""):
    """Append one line to the results file (orders_results.csv)."""
    try:
        new = not os.path.exists(RESULTS_FILE)
        with open(RESULTS_FILE, "a", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            if new:
                w.writerow(["time", "order", "hawb", "status", "note"])
            w.writerow([datetime.now().strftime("%Y-%m-%d %H:%M:%S"), order, hawb, status, note])
    except OSError as e:                                # e.g. the file is open in Excel
        print("  (could not write %s: %s - close it in Excel)" % (RESULTS_FILE, e))


def load_done():
    """Orders already finished (status DONE) in the results file - skipped when the batch is started again."""
    done = set()
    if os.path.exists(RESULTS_FILE):
        try:
            with open(RESULTS_FILE, newline="", encoding="utf-8-sig") as f:
                for row in csv.DictReader(f):
                    if (row.get("status") or "").strip() == "DONE":
                        done.add((row.get("order") or "").strip())
        except OSError:
            pass
    return done


def run_batch(context, page, home_url, pairs):
    """Process every (order, HAWB) pair of the list, one after the other. Returns the exit code."""
    test_mode = DRY_RUN or NO_APPROVE or NO_SIGN
    done = set() if (REDO or test_mode) else load_done()
    results = []
    total = len(pairs)
    for n, (order, hawb) in enumerate(pairs, 1):
        print("\n" + "#" * 72)
        print("# Order %d/%d : %s     HAWB : %s" % (n, total, order, hawb or "(missing)"))
        print("#" * 72)
        if order in done:
            print("  already DONE in %s - skipped (use --redo to run it again)" % RESULTS_FILE)
            results.append((order, hawb, "SKIPPED", "already DONE"))
            continue
        if not hawb:
            print("  no HAWB number in the list for this order - skipped")
            log_result(order, hawb, "NO HAWB", "no HAWB number in the list")
            results.append((order, hawb, "NO HAWB", "no HAWB number in the list"))
            continue
        try:                                            # every order starts clean from the home page
            page = recover_to_home(context, pick_page(context), home_url)
        except Exception as e:
            note = "could not open the home page: " + (str(e).splitlines()[0][:150] if str(e) else type(e).__name__)
            print(" ", note)
            log_result(order, hawb, "FAILED", note)
            results.append((order, hawb, "FAILED", note))
            continue
        code, note = run_with_recovery(context, page, home_url, order, hawb)
        status = ("TEST" if test_mode else "DONE") if code == 0 else ("CHECK" if code == 2 else "FAILED")
        log_result(order, hawb, status, note)
        results.append((order, hawb, status, note))
        page = pick_page(context)
        page.wait_for_timeout(2_000)

    print("\n" + "=" * 72)
    print("BATCH FINISHED - %d order(s)" % total)
    for status in ("DONE", "TEST", "SKIPPED", "NO HAWB", "CHECK", "FAILED"):
        rows = [r for r in results if r[2] == status]
        if rows:
            print("  %-8s %d  %s" % (status, len(rows), ", ".join(r[0] for r in rows)))
    print("  results file:", RESULTS_FILE)
    if any(r[2] == "CHECK" for r in results):
        print("  CHECK = failed after Approve was clicked: open that order in Infor Nexus and look at it by hand.")
    return 1 if any(r[2] in ("FAILED", "CHECK", "NO HAWB") for r in results) else 0


def main():
    pairs = read_pairs(BATCH_FILE) if BATCH_FILE else None     # read the list first: a bad file stops here
    if pairs is not None:
        print("List file:", BATCH_FILE, "-", len(pairs), "order(s)")
    else:
        print("Order number:", ORDER_NUMBER)
        print("HAWB number :", HAWB_NUMBER)
    exit_code = 0
    with sync_playwright() as p:
        browser, context, page = l2.open_session_2(p)       # login with account 2
        home_url = page.url
        if pairs is not None:
            exit_code = run_batch(context, page, home_url, pairs)
        else:
            code, _ = run_with_recovery(context, page, home_url, ORDER_NUMBER, HAWB_NUMBER)
            exit_code = 0 if code == 0 else 1
        if os.getenv("NONINTERACTIVE") != "1" and sys.stdin and sys.stdin.isatty():
            try:
                input("Press Enter to close the browser...")
            except EOFError:
                pass
        browser.close()
    sys.exit(exit_code)


if __name__ == "__main__":
    main()

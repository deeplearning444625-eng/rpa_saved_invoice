"""
Step 2 - Infor Nexus menu navigation (Playwright)
    2.1  click APPLICATIONS (top menu)
    2.2  click Invoices (FINANCE column)
    2.3  type the invoice number in the 'Invoice Number' filter
    2.4  click Apply and wait for the result row
    2.5  click the invoice number link in the result row (opens the invoice)
    2.6  set 2 dates on the invoice from the Invoice Issue Date: Document Date (2),
         Cargo handover/ FCR date (3). Expected Shipment Date is NOT touched.
    2.7  click 'Preview' in the left step menu (under Parties)
    3.1  click 'Print / PDF'                      (marker 1)
    3.2  click 'ADIDAS FINANCIAL INVOICE PDF'     (marker 2)
    3.3  save the PDF into output_pdfs/           (markers 3-6: print icon ->
         Save as PDF -> Save. Chrome's print dialog cannot be clicked by a
         script, so the PDF file is captured/downloaded directly instead.)

Needs nexus_session.json from step 1 (same folder).

Run one invoice:      python infornexus_step2_menu.py SC260475
Run an Excel list:    python infornexus_step2_menu.py invoices.xlsx
                      (or INVOICE_FILE=... / INVOICE_NO=... in .env)

Excel: the invoice numbers are read from the column headed "INV.NO." (falls back to
column H). Blank cells and repeated numbers are skipped, so each invoice runs once.
Optional .env:  SHEET=Sheet1   SKIP_EXISTING=true|false (default true)
Results of every run are written to batch_results.csv.
"""
import csv
import json
import os
import re
import sys
import time
from datetime import date, datetime, timedelta
from pathlib import Path

from dotenv import load_dotenv
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

load_dotenv(Path(__file__).resolve().parent / ".env")

URL = os.getenv("NEXUS_URL", "https://network.infornexus.com")
HEADLESS = os.getenv("HEADLESS", "false").lower() == "true"
SESSION_FILE = Path(__file__).with_name("nexus_session.json")
HOME_FILE = Path(__file__).with_name("nexus_home.txt")   # the page we land on after login
STATE = {"home": URL}
# Fallback date if Invoice Issue Date is empty (DOC_DATE=2026-11-17 in .env). Default: today.
_d = os.getenv("DOC_DATE", "").strip()
DOC_DATE = datetime.strptime(_d, "%Y-%m-%d").date() if _d else date.today()
# extra wait after clicking Preview so the preview page finishes rendering
PREVIEW_WAIT_MS = int(os.getenv("PREVIEW_WAIT_MS", "6000"))
INPUT_ARG = (sys.argv[1] if len(sys.argv) > 1 else os.getenv("INVOICE_FILE") or os.getenv("INVOICE_NO", "SC260475")).strip().strip('"')


def step_2_1_open_applications(page):
    """2.1 - open the APPLICATIONS mega menu."""
    menu = page.get_by_text(re.compile(r"^\s*applications\s*$", re.I)).first
    menu.wait_for(state="visible", timeout=30_000)
    menu.click()
    # menu is open when the FINANCE > Invoices entry becomes visible
    page.get_by_text("Invoices", exact=True).first.wait_for(state="visible", timeout=10_000)
    print("2.1 APPLICATIONS opened")


def step_2_2_click_invoices(page):
    """2.2 - click Invoices under FINANCE."""
    page.get_by_text("Invoices", exact=True).first.click()
    page.wait_for_load_state("networkidle", timeout=30_000)
    print("2.2 Invoices opened:", page.url)


def step_2_3_fill_invoice_number(page, invoice_no):
    """2.3 - fill the 'Invoice Number starts with' filter."""
    page.get_by_text("Please apply filter criteria").first.wait_for(state="visible", timeout=30_000)
    label = page.get_by_text(re.compile(r"^\s*Invoice Number\s*(starts with)?\s*$", re.I)).first
    label.wait_for(state="visible", timeout=15_000)
    box = label.locator("xpath=following::input[not(@readonly) and not(@type='hidden')][1]")
    box.click()
    box.press("Control+A")
    box.press("Delete")
    box.press_sequentially(invoice_no, delay=40)
    if box.input_value() != invoice_no:
        box.fill(invoice_no)
    if box.input_value() != invoice_no:
        raise PWTimeout("Could not type the invoice number into the filter")
    print("2.3 Invoice Number filled:", invoice_no)


def step_2_4_click_apply(page, invoice_no):
    """2.4 - click Apply and wait until the table shows the invoice."""
    page.get_by_role("button", name="Apply", exact=True).click()
    row = page.get_by_text(invoice_no, exact=False).last
    row.wait_for(state="visible", timeout=30_000)
    print("2.4 Apply done - result found:", invoice_no)


def step_2_5_open_invoice(page, invoice_no):
    """2.5 - click the blue invoice number link in the result row.
    Returns the page that shows the invoice (same tab, or a new tab if one opens)."""
    link = page.get_by_role("link", name=invoice_no, exact=True).first
    if not link.count():                      # fallback: the text inside the table row
        link = page.locator("tr", has_text=invoice_no).get_by_text(invoice_no, exact=True).first
    link.wait_for(state="visible", timeout=15_000)

    context = page.context
    try:
        with context.expect_page(timeout=4_000) as new_page_info:
            link.click()
        target = new_page_info.value           # invoice opened in a new tab
    except PWTimeout:
        target = page                          # invoice opened in the same tab
    target.wait_for_load_state("domcontentloaded")
    target.wait_for_load_state("networkidle", timeout=30_000)
    print("2.5 Invoice opened:", target.url)
    return target


# (marker number, label, regex). Expected Shipment Date (1) is NOT touched; 0 = Invoice Issue Date.
DATE_LABELS = [
    (0, "Invoice Issue Date",       r"^\s*Invoice Issue Date\s*$"),      # filled from the Excel date
    (2, "Document Date",            r"^\s*Document Date\s*$"),
    (3, "Cargo handover/ FCR date", r"^\s*Cargo handover/\s*FCR date\s*$"),
]


def _pick(select, wanted):
    """select_option by visible label; try alternatives (e.g. '5' vs '05')."""
    for w in wanted:
        try:
            select.select_option(label=w, timeout=3_000)
            return
        except Exception:
            continue
    raise PWTimeout(f"Option not found: {wanted}")


def _date_selects(page, label_pattern, name):
    label = page.get_by_text(re.compile(label_pattern, re.I)).first
    label.wait_for(state="visible", timeout=15_000)
    label.scroll_into_view_if_needed()
    selects = label.locator("xpath=following::select[position()<=3]")
    if selects.count() < 3:
        raise PWTimeout(f"Date drop-downs not found for: {name}")
    return selects


def _read_date(selects):
    """Return the shown (month, day, year) texts of a 3-drop-down date field."""
    return tuple(
        selects.nth(i).evaluate("e => e.options[e.selectedIndex] ? e.options[e.selectedIndex].text.trim() : ''")
        for i in range(3)
    )


def step_2_6_select_dates(page, fallback_date, only=None, excel_date=None):
    """2.6 - put the invoice date into the 3 date fields
    (Invoice Issue Date, Document Date, Cargo handover/ FCR date).
    Expected Shipment Date is left alone.
    The date comes from the uploaded Excel (excel_date). Only when no Excel date is given
    (single-invoice run) does it copy the 'Invoice Issue Date' shown on the page.
    Each field is verified after selecting and retried up to 3 times.
    only = 2 or 3 re-does just that field."""
    if excel_date:
        month, day, year = excel_date.strftime("%b"), str(excel_date.day), str(excel_date.year)
        print(f"2.6 Date from Excel: {month} / {day} / {year}")
    else:
        issue = _date_selects(page, r"^\s*Invoice Issue Date\s*$", "Invoice Issue Date")
        month, day, year = _read_date(issue)
        months = ("jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec")
        valid = (month[:3].lower() in months and day.isdigit() and 1 <= int(day) <= 31
                 and re.fullmatch(r"\d{4}", year) is not None)
        if valid:
            print(f"2.6 Source date (Invoice Issue Date on page): {month} / {day} / {year}")
        else:                                              # empty or a placeholder like 'Month'/'--'
            month, day, year = fallback_date.strftime("%b"), str(fallback_date.day), str(fallback_date.year)
            print(f"2.6 WARNING: Invoice Issue Date is not a real date (read {_read_date(issue)}) - "
                  f"using fallback date {month} / {day} / {year}")
    day_opts = [day, day.lstrip("0"), f"{int(day):02d}"]

    for n, name, pattern in DATE_LABELS:
        if only and n != only:
            continue
        if n == 0 and not excel_date:          # no Excel date: the page's own Issue Date is the source
            continue
        for attempt in range(1, 4):
            sel = _date_selects(page, pattern, name)
            _pick(sel.nth(0), [month])
            _pick(sel.nth(1), day_opts)
            _pick(sel.nth(2), [year])
            page.wait_for_timeout(300)
            got = _read_date(_date_selects(page, pattern, name))
            if got[0] == month and got[1] in day_opts and got[2] == year:
                print(f"2.6 [{n}] {name}: {month} / {day} / {year}  OK")
                break
            print(f"2.6 [{n}] {name}: attempt {attempt} shows {got}, retrying...")
        else:
            page.screenshot(path=f"debug_date{n}_failed.png", full_page=True)
            raise PWTimeout(f"Date [{n}] {name} did not stick (see debug_date{n}_failed.png)")


def step_2_7_click_preview(page):
    """2.7 - click 'Preview' in the left step menu. Returns the page showing the preview."""
    link = page.get_by_text("Parties", exact=True).first.locator(
        "xpath=following::*[normalize-space(text())='Preview'][1]")
    if not link.count():
        link = page.get_by_text("Preview", exact=True).first
    link.scroll_into_view_if_needed()
    link.wait_for(state="visible", timeout=15_000)

    try:
        with page.context.expect_page(timeout=4_000) as new_page_info:
            link.click()
        target = new_page_info.value          # preview opened in a new tab
    except PWTimeout:
        target = page                         # preview opened in the same tab
    target.wait_for_load_state("domcontentloaded")
    target.wait_for_load_state("networkidle", timeout=30_000)
    target.wait_for_timeout(PREVIEW_WAIT_MS)      # wait for the preview to render (6000 ms)
    print("2.7 Preview opened:", target.url)
    return target


OUTPUT_DIR = Path(__file__).with_name("output_pdfs")


def step_3_1_open_print_menu(page):
    """3.1 - click the 'Print / PDF' drop-down (marker 1)."""
    page.get_by_text("Print / PDF", exact=True).first.click()
    page.get_by_text("ADIDAS FINANCIAL INVOICE PDF", exact=True).first.wait_for(
        state="visible", timeout=10_000)
    print("3.1 Print / PDF menu opened")


def _pdf_candidates(page):
    """Collect every possible PDF address on the page (embed/iframe/object/frames)."""
    found = []
    js = """() => [...document.querySelectorAll('embed,iframe,object')].flatMap(e =>
              [e.src, e.data, e.getAttribute('original-url'), e.getAttribute('src')])"""
    for fr in page.frames:
        try:
            found += fr.evaluate(js)
        except Exception:
            pass
        found.append(fr.url)
    out = []
    for u in found:
        if not u or u.startswith("about:"):
            continue
        if u.startswith("chrome-extension://") and "?" in u:       # Chrome PDF viewer wrapper
            u = u.split("?", 1)[1]
            u = u[4:] if u.startswith("src=") else u
        if u.startswith(("http", "blob:", "data:")) and u not in out and u != page.url:
            out.append(u)
    return out


def _fetch_pdf(page, url):
    """Return PDF bytes for http(s)/blob/data URLs, or None."""
    try:
        if url.startswith("http"):
            data = page.context.request.get(url).body()      # uses the login cookies
        else:
            b64 = page.evaluate(
                """async u => { const b = await (await fetch(u)).arrayBuffer();
                   let s = ''; const a = new Uint8Array(b);
                   for (let i = 0; i < a.length; i += 8192) s += String.fromCharCode(...a.subarray(i, i + 8192));
                   return btoa(s); }""", url)
            import base64
            data = base64.b64decode(b64)
        return data if data[:4] == b"%PDF" else None
    except Exception:
        return None


def _valid_pdf(data):
    """Return the PDF bytes (junk before %PDF removed) or None if it is not a PDF."""
    if not data:
        return None
    i = data[:1024].find(b"%PDF")
    return data[i:] if i >= 0 else None


def step_3_2_and_3_save_financial_pdf(page, invoice_no):
    """3.2 click 'ADIDAS FINANCIAL INVOICE PDF' (marker 2), then
    3.3 save the PDF to output_pdfs/ (replaces markers 3-6 = print dialog).
    Every source is validated (%PDF header) before anything is written."""
    OUTPUT_DIR.mkdir(exist_ok=True)
    target = OUTPUT_DIR / f"{invoice_no}-INV.pdf"
    if target.exists():
        target.unlink()                      # remove a broken file from an earlier run
    ctx = page.context
    resps, dls, log = [], [], []

    def on_response(r):
        try:
            ct = (r.headers.get("content-type") or "").lower()
            if "pdf" in ct or r.url.lower().split("?")[0].endswith(".pdf"):
                resps.append(r)
        except Exception:
            pass

    def on_download(d):
        dls.append(d)

    ctx.on("response", on_response)
    page.on("download", on_download)
    try:
        page.get_by_text("ADIDAS FINANCIAL INVOICE PDF", exact=True).first.click()
        print("3.2 ADIDAS FINANCIAL INVOICE PDF clicked")

        for _ in range(30):                              # wait up to 15 s
            if resps or dls:
                break
            page.wait_for_timeout(500)
        page.wait_for_timeout(1_000)

        data = None

        # a) real file download
        for d in dls:
            try:
                tmp = OUTPUT_DIR / "_download.tmp"
                d.save_as(str(tmp))
                raw = tmp.read_bytes()
                tmp.unlink()
                log.append(f"download: {d.suggested_filename}, {len(raw)} bytes, head={raw[:40]!r}")
                data = _valid_pdf(raw)
                if data:
                    break
            except Exception as e:
                log.append(f"download error: {e}")

        # b) PDF responses seen on the network
        for r in resps if not data else []:
            try:
                raw = r.body()
            except Exception as e:
                raw = b""
                log.append(f"body() error: {e}")
            log.append(f"response: {r.status} {r.url[:120]} type={r.headers.get('content-type')} "
                       f"{len(raw)} bytes head={raw[:40]!r}")
            data = _valid_pdf(raw)
            if not data:                                  # retry as a full download with the login cookies
                try:
                    raw = ctx.request.get(r.url).body()
                    log.append(f"  refetch: {len(raw)} bytes head={raw[:40]!r}")
                    data = _valid_pdf(raw)
                except Exception as e:
                    log.append(f"  refetch error: {e}")
            if data:
                break

        # c) PDF shown in a viewer (embed / iframe / blob)
        cands = []
        if not data:
            cands = _pdf_candidates(page)
            for u in cands:
                data = _fetch_pdf(page, u)
                log.append(f"candidate: {u[:120]} -> {'PDF' if data else 'not a PDF'}")
                if data:
                    break
    finally:
        ctx.remove_listener("response", on_response)
        page.remove_listener("download", on_download)

    if not data:
        page.screenshot(path="debug_step3_failed.png")
        Path("debug_step3_log.txt").write_text(
            "page: " + page.url + "\n\n" + "\n".join(log) + "\n\nframes:\n"
            + "\n".join(f.url for f in page.frames), encoding="utf-8")
        raise RuntimeError("PDF not captured - see debug_step3_log.txt and debug_step3_failed.png")

    target.write_bytes(data)
    print("3.3 PDF saved:", target, f"({len(data):,} bytes)")
    return target


def step_3_2_gui_print(page, invoice_no):
    """3.2 click 'ADIDAS FINANCIAL INVOICE PDF', then markers 1-4 through Chrome's print
    dialog with print_gui.py (default; PRINT_MODE=direct in .env switches to direct capture)."""
    from print_gui import save_pdf_via_print_dialog
    OUTPUT_DIR.mkdir(exist_ok=True)
    target = OUTPUT_DIR / f"{invoice_no}-INV.pdf"
    page.get_by_text("ADIDAS FINANCIAL INVOICE PDF", exact=True).first.click()
    print("3.2 ADIDAS FINANCIAL INVOICE PDF clicked")
    page.wait_for_timeout(4_000)          # let the PDF viewer finish loading
    page.bring_to_front()                 # print_gui clicks on the real screen
    save_pdf_via_print_dialog(target, page.title())
    print("3.3 PDF saved:", target)
    return target


def is_logged_in(page, timeout=15_000, grace=8_000):
    """True once the top menu is visible.

    - The login form can flash for a moment while the site checks the session, so only give
      up if it stays on screen for `grace` ms (or nothing shows up within `timeout` ms).
    - If Nexus shows 'session timed out - Re-Activate', it is re-activated with the password
      (once) and we keep waiting for the menu.
    """
    from infornexus_step1_login import reactivate_if_needed
    menu = page.get_by_text(re.compile(r"^\s*applications\s*$", re.I)).first
    form = page.get_by_placeholder("Password").first
    t0 = time.time()
    form_since = None
    reactivated = False
    while (time.time() - t0) * 1000 < timeout:
        try:
            menu_visible = menu.is_visible()
            form_visible = form.is_visible()
        except Exception:
            menu_visible = form_visible = False
        if menu_visible:
            return True
        if not reactivated and reactivate_if_needed(page):
            reactivated = True
            form_since = None
            t0 = time.time()                       # fresh wait for the menu after re-activating
            continue
        if form_visible:
            form_since = form_since or time.time()
            if (time.time() - form_since) * 1000 > grace:
                return False
        else:
            form_since = None
        page.wait_for_timeout(500)
    return False


def load_home_url():
    try:
        return HOME_FILE.read_text(encoding="utf-8").strip() or URL
    except OSError:
        return URL


def remember_home(page):
    """The bare NEXUS_URL is the login page and can show the login form even when we are
    logged in. So remember the page we actually land on after login and use that as 'home'."""
    STATE["home"] = page.url
    try:
        HOME_FILE.write_text(page.url, encoding="utf-8")
    except OSError:
        pass
    print("Home page:", page.url)


def read_invoices_from_excel(path):
    """Unique invoice numbers (in order) from the 'INV.NO.' column of an .xlsx/.xlsm file."""
    import openpyxl
    p = Path(path)
    if not p.exists():
        sys.exit(f"Excel file not found: {p}")
    if p.suffix.lower() not in (".xlsx", ".xlsm"):
        sys.exit("Please save the list as .xlsx (the old .xls format is not supported)")
    wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
    sheet = os.getenv("SHEET", "").strip()
    ws = wb[sheet] if sheet else wb.active
    rows = list(ws.iter_rows(values_only=True))

    col, start = None, 0
    for i, row in enumerate(rows[:40]):                       # find the header row
        for j, v in enumerate(row):
            if isinstance(v, str) and re.sub(r"[^a-z]", "", v.lower()) in ("invno", "invoiceno", "invoicenumber"):
                col, start = j, i + 1
                break
        if col is not None:
            break
    if col is None:
        col, start = 7, 0                                      # column H
        print("Header 'INV.NO.' not found - using column H")

    seen, out = set(), []
    for row in rows[start:]:
        v = row[col] if col < len(row) else None
        if v is None:
            continue
        v = str(v).strip().upper()
        if not v or not re.search(r"\d", v) or v in seen:     # blank / not an invoice / repeated
            continue
        seen.add(v)
        out.append(v)
    return out


def _parse_date(v):
    """Excel cell or text -> date (day-first for text like 10/11/2026)."""
    if v is None:
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, (int, float)) and 30000 < v < 80000:          # Excel serial number
        return date(1899, 12, 30) + timedelta(days=int(v))
    t = str(v).strip()
    if not t:
        return None
    for f in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%d-%b-%Y", "%d-%b-%y",
              "%d %b %Y", "%d %B %Y", "%b %d, %Y", "%d/%m/%y"):
        try:
            return datetime.strptime(t, f).date()
        except ValueError:
            pass
    return None


DATE_CONFLICTS = {}


def read_invoice_dates_from_excel(path):
    """{invoice_no: date or None} from the same sheet as the 'INV.NO.' column.
    Date column = header named in .env DATE_COLUMN, else the first header like
    'Customer Request Date (CRD)' (or INV DATE / Invoice Date / Issue Date / Document Date / Date).
    If one invoice has different dates on different rows, it is left as None and listed in
    DATE_CONFLICTS, so a date is never picked by guess."""
    import openpyxl
    p = Path(path)
    if not p.exists():
        raise ValueError(f"Excel file not found: {p}")
    wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
    sheet = os.getenv("SHEET", "").strip()
    ws = wb[sheet] if sheet else wb.active
    rows = list(ws.iter_rows(values_only=True))
    norm = lambda v: re.sub(r"[^a-z]", "", str(v).lower())

    inv_col, hdr = None, None
    for i, row in enumerate(rows[:40]):
        for j, v in enumerate(row):
            if isinstance(v, str) and norm(v) in ("invno", "invoiceno", "invoicenumber"):
                inv_col, hdr = j, i
                break
        if inv_col is not None:
            break
    if inv_col is None:
        inv_col, hdr = 7, -1                                       # column H, no header row

    wanted = norm(os.getenv("DATE_COLUMN", ""))
    names = ("customerrequestdatecrd", "customerrequestdate", "crd", "invdate", "invoicedate", "invoiceissuedate", "issuedate", "documentdate", "docdate", "date")
    date_col = None
    headers = rows[hdr] if hdr >= 0 else ()
    if wanted:
        date_col = next((j for j, v in enumerate(headers) if v is not None and norm(v) == wanted), None)
    else:
        for nm in names:
            date_col = next((j for j, v in enumerate(headers) if isinstance(v, str) and norm(v) == nm), None)
            if date_col is not None:
                break
    if date_col is None:
        shown = [str(v) for v in headers if v is not None]
        raise ValueError("No date column found in the Excel (looked for INV DATE / Invoice Date / "
                         f"Issue Date / Date; set DATE_COLUMN in .env). Headers seen: {shown}")

    out, seen = {}, {}
    DATE_CONFLICTS.clear()
    for row in rows[hdr + 1:]:
        inv = row[inv_col] if inv_col < len(row) else None
        if inv is None:
            continue
        inv = str(inv).strip().upper()
        if not inv or not re.search(r"\d", inv):
            continue
        d = _parse_date(row[date_col] if date_col < len(row) else None)
        out.setdefault(inv, None)
        if d:
            seen.setdefault(inv, set()).add(d)
    for inv, ds in seen.items():
        if len(ds) == 1:
            out[inv] = next(iter(ds))
        else:                                       # different dates on different rows
            DATE_CONFLICTS[inv] = sorted(ds)
    return out


def process_invoice(page, invoice_no, excel_date=None):
    """Run steps 2.1 ... 3.3 for one invoice (page = the main tab, on the home page)."""
    step_2_1_open_applications(page)
    step_2_2_click_invoices(page)
    step_2_3_fill_invoice_number(page, invoice_no)
    step_2_4_click_apply(page, invoice_no)
    page = step_2_5_open_invoice(page, invoice_no)
    step_2_6_select_dates(page, DOC_DATE, excel_date=excel_date)
    page = step_2_7_click_preview(page)
    step_3_1_open_print_menu(page)
    if os.getenv("PRINT_MODE", "gui").lower() != "direct":
        step_3_2_gui_print(page, invoice_no)
    else:
        step_3_2_and_3_save_financial_pdf(page, invoice_no)


def back_to_home(context, main_page):
    """Close extra tabs and return the main tab to the home page (the page seen after login)."""
    for pg in list(context.pages):
        if pg is not main_page:
            try:
                pg.close()
            except Exception:
                pass
    main_page.bring_to_front()
    main_page.goto(STATE["home"], wait_until="domcontentloaded")
    if is_logged_in(main_page, timeout=30_000):
        return
    print(f"Home page showed the login form (url: {main_page.url}) - logging in again...")
    from infornexus_step1_login import login
    login(main_page)                      # leaves the page on the home page; do NOT goto(URL) again
    if not is_logged_in(main_page, timeout=30_000):
        main_page.screenshot(path="debug_relogin_failed.png", full_page=True)
        raise RuntimeError("Logged out even after logging in again - see debug_relogin_failed.png")
    context.storage_state(path=str(SESSION_FILE))
    remember_home(main_page)


def main():
    arg = INPUT_ARG
    listed = [x.strip().upper() for x in os.getenv("INVOICE_LIST", "").split(",") if x.strip()]
    if listed:                                      # e.g. chosen in the Streamlit app
        invoices = listed
        print(f"{len(invoices)} invoices: {', '.join(invoices)}")
    elif arg.lower().endswith((".xlsx", ".xlsm", ".xls")):
        invoices = read_invoices_from_excel(arg)
        if not invoices:
            sys.exit("No invoice numbers found in the Excel file")
        print(f"{len(invoices)} unique invoices found in {Path(arg).name}: {', '.join(invoices)}")
    else:
        invoices = [arg]

    dates, excel_mode = {}, False                   # per-invoice dates from the Excel
    raw = os.getenv("INVOICE_DATES", "").strip()
    if raw:                                         # sent by the Streamlit app
        excel_mode = True
        dates = {k.upper(): _parse_date(v) for k, v in json.loads(raw).items()}
    elif arg.lower().endswith((".xlsx", ".xlsm", ".xls")):
        excel_mode = True
        try:
            dates = read_invoice_dates_from_excel(arg)
        except ValueError as e:
            sys.exit(str(e))

    skip_existing = os.getenv("SKIP_EXISTING", "true").lower() != "false"
    results = []                                   # (invoice, status, detail)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=HEADLESS, args=["--start-maximized"])
        kwargs = {"no_viewport": True}      # use the real window size (small screens too)
        if SESSION_FILE.exists():
            kwargs["storage_state"] = str(SESSION_FILE)
        context = browser.new_context(**kwargs)
        context.on("page", lambda pg: pg.on("dialog", lambda d: d.accept()))   # 'leave page?' prompts
        page = context.new_page()
        page.on("dialog", lambda d: d.accept())
        STATE["home"] = load_home_url()
        page.goto(STATE["home"], wait_until="domcontentloaded")

        if not is_logged_in(page):
            print("Saved session not valid - logging in again...")
            from infornexus_step1_login import login   # reuse step 1
            login(page)
            if not is_logged_in(page, timeout=30_000):
                page.screenshot(path="debug_step2_nomenu.png", full_page=True)
                sys.exit("Logged in but top menu not found - see debug_step2_nomenu.png")
            context.storage_state(path=str(SESSION_FILE))
        remember_home(page)
        context.storage_state(path=str(SESSION_FILE))   # keep the saved cookies fresh

        failures_in_a_row = 0
        for n, inv in enumerate(invoices, start=1):
            print(f"\n===== [{n}/{len(invoices)}] {inv} =====")
            if skip_existing and (OUTPUT_DIR / f"{inv}-INV.pdf").exists():
                print("PDF already exists - skipped (SKIP_EXISTING=false to redo)")
                results.append((inv, "SKIPPED", "PDF already exists"))
                continue
            if excel_mode and not dates.get(inv):
                print(f"FAILED {inv}: no usable date in the Excel (missing or different dates on different rows) - nothing was changed")
                results.append((inv, "FAILED", "No usable date in Excel"))
                continue
            try:
                back_to_home(context, page)
                process_invoice(page, inv, dates.get(inv))
                results.append((inv, "OK", f"{inv}-INV.pdf"))
                failures_in_a_row = 0
            except (PWTimeout, RuntimeError) as e:
                msg = str(e).splitlines()[0][:200]
                print(f"FAILED {inv}: {msg}")
                results.append((inv, "FAILED", msg))
                failures_in_a_row += 1
                try:
                    (context.pages[-1]).screenshot(path=f"debug_{inv}_failed.png")
                except Exception:
                    pass
                if failures_in_a_row >= 3:
                    print("3 failures in a row - stopping so nothing is clicked blindly.")
                    break

        # summary + results file
        try:
            from print_gui import restore_console
            restore_console()
        except Exception:
            pass
        with open("batch_results.csv", "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(["invoice", "status", "detail", "time"])
            stamp = time.strftime("%Y-%m-%d %H:%M:%S")
            for r in results:
                w.writerow([*r, stamp])
        ok = sum(1 for r in results if r[1] == "OK")
        print(f"\nDone: {ok} OK, {sum(1 for r in results if r[1]=='SKIPPED')} skipped, "
              f"{sum(1 for r in results if r[1]=='FAILED')} failed  ->  batch_results.csv")
        # On Windows, stdin=DEVNULL still reports isatty()=True, so also check the
        # NONINTERACTIVE flag set by the Streamlit app and catch EOFError.
        if os.getenv("NONINTERACTIVE") != "1" and sys.stdin and sys.stdin.isatty():
            try:
                input("Press Enter to close the browser...")
            except EOFError:
                pass
        browser.close()


if __name__ == "__main__":
    main()

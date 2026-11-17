"""
dispatched_loop - run the Dispatched_Step process for every invoice in a sheet.

For each INV.NO. in the sheet it picks ONE PO number (the first PO listed under that invoice),
searches that PO in Nexus Dispatches, and runs dispatched_step's steps 1-10 with that invoice
number. One dispatch holds all the POs of an invoice, so one PO is enough to find it.

Keep these files in the same folder:
    dispatched_loop.py, dispatched_step.py, infornexus_step1_login.py, infornexus_step2_menu.py, .env
    pip install openpyxl        (only needed for .xlsx files)

Sheet: .xlsx or .csv with a header row containing 'INV.NO.' and 'PO Number'. Blank rows are ignored.

Run (PowerShell, from the project folder):
    python dispatched_loop.py invoices.xlsx --dry          <- only show the plan (no browser)
    python dispatched_loop.py invoices.xlsx                <- run every invoice
    python dispatched_loop.py invoices.xlsx --only SC260478,SC260479
    (file can also be set with INVOICE_FILE=... in .env)

Options:
    --sheet NAME   sheet to read (default: the active one)
    --only A,B     run only these invoice numbers
    --dry          list invoice -> PO and stop
    --redo         ignore the log and run invoices already marked DONE / UNCERTAIN / SKIPPED again

Every invoice result is appended to dispatched_loop_log.csv (DONE / FAILED / UNCERTAIN / SKIPPED).
On the next run, DONE, UNCERTAIN and SKIPPED invoices are skipped so nothing is signed twice;
FAILED ones are tried again. UNCERTAIN = Ok in the Sign dialog was clicked but the result
could not be confirmed: check that dispatch in Nexus by hand. SKIPPED = the dispatch was
already approved in Nexus (Preview page without Validate / Approve): nothing was done.
"""
import argparse
import csv
import datetime
import os
import re
import sys
from pathlib import Path

LOG_FILE = Path("dispatched_loop_log.csv")


# ---------------------------------------------------------------- reading the sheet
def clean_po(value):
    """PO number as digits text, or '' if the cell is not a PO number."""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    text = str(value).strip()
    return text if re.fullmatch(r"\d{6,}", text) else ""


def read_invoices(path, sheet=None):
    """Return {invoice number: [PO numbers in sheet order]} (invoices in sheet order)."""
    path = Path(path)
    if not path.exists():
        sys.exit("File not found: %s" % path)
    if path.suffix.lower() in (".xlsx", ".xlsm"):
        try:
            from openpyxl import load_workbook
        except ImportError:
            sys.exit("openpyxl is missing: pip install openpyxl")
        wb = load_workbook(path, data_only=True, read_only=True)
        ws = wb[sheet] if sheet else wb.active
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
    else:
        with open(path, newline="", encoding="utf-8-sig") as f:
            rows = list(csv.reader(f))

    inv_col = po_col = hdr = None
    for i, r in enumerate(rows[:50]):
        ic = next((j for j, v in enumerate(r) if v and re.search(r"inv\w*\.?\s*no", str(v), re.I)), None)
        pc = next((j for j, v in enumerate(r) if v and re.search(r"po\s*number", str(v), re.I)), None)
        if ic is not None and pc is not None:
            inv_col, po_col, hdr = ic, pc, i
            break
    if hdr is None:
        sys.exit("Could not find the 'INV.NO.' and 'PO Number' headers in the first 50 rows of %s" % path)

    groups = {}
    for r in rows[hdr + 1:]:
        inv = r[inv_col] if inv_col < len(r) else None
        po = clean_po(r[po_col] if po_col < len(r) else None)
        inv = str(inv).strip() if inv is not None else ""
        if not inv or not po:
            continue
        pos = groups.setdefault(inv, [])
        if po not in pos:
            pos.append(po)
    return groups


# ---------------------------------------------------------------- result log
def load_log():
    """Last status of every invoice in the log file."""
    status = {}
    if LOG_FILE.exists():
        with open(LOG_FILE, newline="", encoding="utf-8") as f:
            for row in csv.DictReader(f):
                status[row["invoice"]] = row["status"]
    return status


def write_log(invoice, po, status, note=""):
    new = not LOG_FILE.exists()
    with open(LOG_FILE, "a", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        if new:
            w.writerow(["time", "invoice", "po", "status", "note"])
        w.writerow([datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), invoice, po, status, note])


# ---------------------------------------------------------------- the loop
def run_all(todo):
    import dispatched_step as ds                       # imported here so --dry needs no browser files
    from playwright.sync_api import sync_playwright

    results = {}
    with sync_playwright() as p:
        browser, context, page = ds.open_session(p)
        for n, (inv, po, all_pos) in enumerate(todo, 1):
            print("\n" + "#" * 64)
            print("[%d/%d] Invoice %s | searching PO %s (%d PO in this invoice)" % (n, len(todo), inv, po, len(all_pos)))
            print("#" * 64)
            ds.PROGRESS["sign_clicked"] = False        # fresh invoice: allow restarts again
            ds.PROGRESS["note"] = ''
            note = ""
            try:
                code, _ = ds.run_with_recovery(context, page, po, inv)
            except Exception as e:                     # never let one invoice stop the loop
                code, note = 1, "%s %s" % (type(e).__name__, str(e).splitlines()[0][:150] if str(e) else "")
                print("FAILED:", note)
            status = {0: "DONE", 2: "UNCERTAIN", 3: "SKIPPED"}.get(code, "FAILED")
            if status == "SKIPPED":
                note = ds.PROGRESS.get("note", "") or "already approved in Nexus"
            results[inv] = status
            write_log(inv, po, status, note)
            print("-> %s %s" % (inv, status))
            try:                                       # back to the home page for the next invoice
                page = ds.recover_to_home(context, page)
            except Exception as e:
                print("  could not return to home (next invoice will try again):",
                      str(e).splitlines()[0][:150] if str(e) else type(e).__name__)

        if os.getenv("NONINTERACTIVE") != "1" and sys.stdin and sys.stdin.isatty():
            try:
                input("\nPress Enter to close the browser...")
            except EOFError:
                pass
        browser.close()
    return results


def main():
    ap = argparse.ArgumentParser(description="Run Dispatched_Step for every invoice in a sheet")
    ap.add_argument("file", nargs="?", default=os.getenv("INVOICE_FILE", "invoices.xlsx"))
    ap.add_argument("--sheet")
    ap.add_argument("--only", help="comma separated invoice numbers to run")
    ap.add_argument("--dry", action="store_true", help="show the plan and stop")
    ap.add_argument("--redo", action="store_true", help="ignore the log (run DONE / UNCERTAIN / SKIPPED again)")
    args = ap.parse_args()

    groups = read_invoices(args.file, args.sheet)
    if not groups:
        sys.exit("No invoice / PO rows found in %s" % args.file)
    if args.only:
        wanted = {x.strip() for x in args.only.split(",") if x.strip()}
        missing = wanted - set(groups)
        if missing:
            print("Not in the sheet:", ", ".join(sorted(missing)))
        groups = {k: v for k, v in groups.items() if k in wanted}

    log = {} if args.redo else load_log()
    todo = []
    print("%-14s %-12s %s" % ("INV.NO.", "PO to search", "POs in invoice"))
    for inv, pos in groups.items():
        state = log.get(inv)
        skip = state in ("DONE", "UNCERTAIN", "SKIPPED")
        print("%-14s %-12s %d%s" % (inv, pos[0], len(pos), "   (skipped: %s in log)" % state if skip else ""))
        if not skip:
            todo.append((inv, pos[0], pos))
    print("\n%d invoice(s) in the sheet, %d to run." % (len(groups), len(todo)))
    if args.dry or not todo:
        return

    results = run_all(todo)
    print("\n" + "=" * 64)
    for inv, st in results.items():
        print("%-14s %s" % (inv, st))
    bad = [i for i, s in results.items() if s not in ("DONE", "SKIPPED")]
    print("\nDone: %d | Not done: %d (see %s)" % (len(results) - len(bad), len(bad), LOG_FILE))
    sys.exit(1 if bad else 0)


if __name__ == "__main__":
    main()

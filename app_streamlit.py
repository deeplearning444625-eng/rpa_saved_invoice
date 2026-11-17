"""
Infor Nexus - Invoice PDF Robot (Streamlit UI)

Run on the SAME Windows PC that has the screen (the robot opens a browser and
clicks Chrome's print dialog on that screen):

    pip install streamlit pandas openpyxl playwright python-dotenv pyautogui opencv-python pillow
    streamlit run app_streamlit.py

Files needed in the same folder: infornexus_step1_login.py, infornexus_step2_menu.py,
print_gui.py, .env (NEXUS_USER / NEXUS_PASS).
For the Dispatch process (sidebar) also: dispatched_step.py, dispatched_loop.py.
For the Orders process (reopen + HAWB + approve + sign) also: orders_step.py, infornexus_login2.py
and NEXUS_USER_2 / NEXUS_PASS_2 in .env.
"""
import hashlib
import json
import os
import re
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st
from dotenv import dotenv_values

IS_WINDOWS = os.name == "nt"

BASE = Path(__file__).parent
UPLOAD_DIR = BASE / "uploads"
OUTPUT_DIR = BASE / "output_pdfs"
LOG_FILE = BASE / "robot_log.txt"
RESULTS_CSV = BASE / "batch_results.csv"
SCRIPT = BASE / "infornexus_step2_menu.py"
D_SCRIPT = BASE / "dispatched_loop.py"                 # Dispatch process (runs dispatched_step per invoice)
D_LOG_FILE = BASE / "dispatch_robot_log.txt"
D_RESULTS_CSV = BASE / "dispatched_loop_log.csv"       # written by dispatched_loop.py
O_SCRIPT = BASE / "orders_step.py"                     # Orders process (reopen + HAWB + approve + sign)
O_LOG_FILE = BASE / "orders_robot_log.txt"
O_RESULTS_CSV = BASE / "orders_results.csv"            # written by orders_step.py (--batch)

st.set_page_config(page_title="Nexus Invoice PDF Robot", page_icon="🧾", layout="wide")

# On Streamlit Cloud the credentials come from Secrets instead of a .env file
try:
    for _k in ("NEXUS_USER", "NEXUS_PASS", "NEXUS_URL", "NEXUS_ACCESS_CODE",
               "NEXUS_USER_2", "NEXUS_PASS_2", "NEXUS_ACCESS_CODE_2"):
        if _k in st.secrets:
            os.environ[_k] = str(st.secrets[_k])
except Exception:
    pass                                      # no secrets file (normal when running locally)


# ---------------------------------------------------------------- helpers
def read_invoices(xlsx_path):
    """Reuse the robot's own Excel reader so the UI and the robot always agree."""
    sys.path.insert(0, str(BASE))
    import importlib
    mod = importlib.import_module("infornexus_step2_menu")
    try:
        return mod.read_invoices_from_excel(xlsx_path)
    except SystemExit as e:                       # the reader exits with a message
        raise ValueError(str(e))


def read_dates(xlsx_path):
    """{invoice: date or None} from the Excel (same reader the robot uses)."""
    sys.path.insert(0, str(BASE))
    import importlib
    mod = importlib.import_module("infornexus_step2_menu")
    dates = mod.read_invoice_dates_from_excel(xlsx_path)
    st.session_state["date_conflicts"] = dict(mod.DATE_CONFLICTS)
    return dates


def is_running():
    p = st.session_state.get("proc")
    return p is not None and p.poll() is None


def start_robot(invoices, opts):
    env = os.environ.copy()
    env.update({
        "INVOICE_LIST": ",".join(invoices),
        "PRINT_MODE": opts["print_mode"],
        "SKIP_EXISTING": "true" if opts["skip_existing"] else "false",
        "PREVIEW_WAIT_MS": str(opts["preview_wait"]),
        "HEADLESS": "true" if opts["headless"] else "false",
        "NONINTERACTIVE": "1",                    # never wait for 'Press Enter'
        "INVOICE_DATES": json.dumps({i: st.session_state.get("inv_dates", {})[i].strftime("%Y-%m-%d")
                                    for i in invoices
                                    if st.session_state.get("inv_dates", {}).get(i)}),
        "PYTHONUNBUFFERED": "1",
        "PYTHONIOENCODING": "utf-8",
    })
    if opts["doc_date"]:
        env["DOC_DATE"] = opts["doc_date"].strftime("%Y-%m-%d")
    if RESULTS_CSV.exists():
        RESULTS_CSV.unlink()
    log = open(LOG_FILE, "w", encoding="utf-8")
    st.session_state["proc"] = subprocess.Popen(
        [sys.executable, "-u", str(SCRIPT)],
        cwd=str(BASE), env=env, stdout=log, stderr=subprocess.STDOUT,
        stdin=subprocess.DEVNULL,                 # no 'Press Enter' prompt
        start_new_session=not IS_WINDOWS,         # lets Stop kill the browser too (Linux)
    )
    st.session_state["started"] = time.time()
    st.session_state["total"] = len(invoices)


def kill_robot(proc):
    """Stop the robot and the browser it started."""
    if IS_WINDOWS:
        subprocess.run(["taskkill", "/F", "/T", "/PID", str(proc.pid)], capture_output=True)
    else:
        import signal
        try:
            os.killpg(os.getpgid(proc.pid), signal.SIGKILL)
        except Exception:
            proc.kill()


@st.cache_resource(show_spinner="Installing the browser (first start only)...")
def ensure_browser():
    """On a Linux server (Streamlit Cloud) download Playwright's Chromium once."""
    if not IS_WINDOWS:
        subprocess.run([sys.executable, "-m", "playwright", "install", "chromium"], check=False)
    return True


ensure_browser()


def read_log():
    if not LOG_FILE.exists():
        return ""
    return LOG_FILE.read_text(encoding="utf-8", errors="replace")


# ---------------------------------------------------------------- dispatch process
def d_is_running():
    p = st.session_state.get("dproc")
    return p is not None and p.poll() is None


def read_dispatch_plan(xlsx_path):
    """{invoice: [PO numbers]} from the Excel, using dispatched_loop's own reader."""
    sys.path.insert(0, str(BASE))
    import importlib
    mod = importlib.import_module("dispatched_loop")
    try:
        return mod.read_invoices(str(xlsx_path))
    except SystemExit as e:                       # the reader exits with a message
        raise ValueError(str(e))


def d_log_status():
    """{invoice: last status} from dispatched_loop_log.csv."""
    status = {}
    if D_RESULTS_CSV.exists():
        try:
            df = pd.read_csv(D_RESULTS_CSV, dtype=str, keep_default_na=False)
            for inv, st_ in zip(df["invoice"], df["status"]):
                status[inv] = st_
        except Exception:
            pass
    return status


def d_read_log():
    if not D_LOG_FILE.exists():
        return ""
    return D_LOG_FILE.read_text(encoding="utf-8", errors="replace")


def start_dispatch(xlsx_path, invoices, opts):
    env = os.environ.copy()
    env.update({
        "NONINTERACTIVE": "1",                    # never wait for 'Press Enter'
        "HEADLESS": "true" if opts["headless"] else "false",
        "MAX_ATTEMPTS": str(opts["max_attempts"]),
        "PYTHONUNBUFFERED": "1",
        "PYTHONIOENCODING": "utf-8",
    })
    cmd = [sys.executable, "-u", str(D_SCRIPT), str(xlsx_path), "--only", ",".join(invoices)]
    if opts["redo"]:
        cmd.append("--redo")
    log = open(D_LOG_FILE, "w", encoding="utf-8")
    st.session_state["dproc"] = subprocess.Popen(
        cmd, cwd=str(BASE), env=env, stdout=log, stderr=subprocess.STDOUT,
        stdin=subprocess.DEVNULL,
        start_new_session=not IS_WINDOWS,         # lets Stop kill the browser too (Linux)
    )
    st.session_state["d_started"] = time.time()
    st.session_state["d_started_str"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    st.session_state["d_total"] = len(invoices)


def dispatch_page():
    busy = d_is_running()
    other_busy = is_running() or o_is_running()   # the PDF robot / the Orders robot
    left, right = st.columns([1, 1.4], gap="large")

    # ---------- left: upload + options
    with left:
        st.subheader("1 · Upload Excel")
        up = st.file_uploader("Excel with INV.NO. and PO Number columns", type=["xlsx", "xlsm"],
                              key="d_upload", disabled=busy or other_busy)
        plan, saved_path = {}, None
        if up is not None:
            key = f"{up.name}-{up.size}"
            if st.session_state.get("d_up_key") != key:      # read the file only once, not on every refresh
                UPLOAD_DIR.mkdir(exist_ok=True)
                saved = UPLOAD_DIR / f"{datetime.now():%Y%m%d_%H%M%S}_{up.name}"
                saved.write_bytes(up.getbuffer())
                try:
                    st.session_state["d_plan"] = read_dispatch_plan(saved)
                    st.session_state["d_error"] = None
                except Exception as e:
                    st.session_state["d_plan"] = {}
                    st.session_state["d_error"] = str(e)
                st.session_state["d_path"] = str(saved)
                st.session_state["d_up_key"] = key
            plan = st.session_state.get("d_plan", {})
            saved_path = st.session_state.get("d_path")
            if st.session_state.get("d_error"):
                st.error(f"Could not read the file: {st.session_state['d_error']}")
            elif plan:
                st.success(f"{len(plan)} invoices found - the first PO of each invoice is searched")
        else:
            st.session_state.pop("d_up_key", None)

        status = d_log_status()
        chosen = []
        if plan:
            st.dataframe(pd.DataFrame({
                "invoice": list(plan),
                "PO to search": [p[0] for p in plan.values()],
                "POs in invoice": [len(p) for p in plan.values()],
                "last status": [status.get(i, "") for i in plan]}),
                hide_index=True, use_container_width=True)
            chosen = st.multiselect("Invoices to run", list(plan), default=list(plan), disabled=busy)

        st.subheader("2 · Options")
        skip_done = st.checkbox("Skip invoices already DONE / UNCERTAIN / SKIPPED in the log", value=True, disabled=busy)
        max_attempts = st.number_input("Attempts per invoice (rebuild on error)", 1, 10, 3, 1, disabled=busy)
        with st.expander("More"):
            headless = st.checkbox("Headless browser", value=False, disabled=busy)
            if not IS_WINDOWS:
                headless = True                   # a server has no screen

        to_run = [i for i in chosen if not (skip_done and status.get(i) in ("DONE", "UNCERTAIN", "SKIPPED"))]

        st.subheader("3 · Run")
        st.warning("This robot **approves and signs** dispatches in Infor Nexus. Try one invoice first.")
        if chosen and not to_run:
            st.info("Every selected invoice is already DONE / UNCERTAIN / SKIPPED in the log. "
                    "Untick the skip option to run them again.")
        confirmed = st.checkbox("I understand the dispatches of the selected invoices will be approved and signed",
                                value=False, disabled=busy)
        if other_busy:
            st.caption("Another robot is running - wait for it to finish.")
        c1, c2 = st.columns(2)
        if c1.button("▶ Start dispatch", type="primary", key="d_start", use_container_width=True,
                     disabled=busy or other_busy or not to_run or not confirmed or not saved_path):
            start_dispatch(saved_path, chosen, dict(redo=not skip_done, headless=headless,
                                                    max_attempts=int(max_attempts)))
            st.rerun()
        if c2.button("■ Stop", key="d_stop", disabled=not busy, use_container_width=True):
            kill_robot(st.session_state["dproc"])
            time.sleep(1)
            st.rerun()

    # ---------- right: progress, results
    with right:
        st.subheader("Progress")
        log = d_read_log()
        cur = None
        for m in re.finditer(r"\[(\d+)/(\d+)\] Invoice (\S+) \|", log):
            cur = m
        if cur:
            n, total, inv = int(cur.group(1)), int(cur.group(2)), cur.group(3)
            done_n = n - 1 if busy else n
            st.progress(min(done_n / total, 1.0) if total else 0.0,
                        text=f"{'Working on' if busy else 'Last'}: {inv}  ({n}/{total})")
        if busy:
            st.info("🟢 Dispatch robot is running…")
        elif st.session_state.get("dproc") is not None:
            code = st.session_state["dproc"].returncode
            (st.success if code == 0 else st.error)(
                "Finished" if code == 0 else f"Finished with problems or stopped (exit code {code}) - see Results and the log")

        with st.expander("Live log", expanded=busy):
            st.code("\n".join(log.splitlines()[-40:]) or "(no log yet)", language="text")

        if D_RESULTS_CSV.exists():
            try:
                rdf = pd.read_csv(D_RESULTS_CSV, dtype=str, keep_default_na=False)
            except Exception:
                rdf = pd.DataFrame()
            if not rdf.empty:
                started = st.session_state.get("d_started_str")
                st.subheader("Results (this run)")
                this = rdf[rdf["time"] >= started] if started else rdf.iloc[0:0]
                if this.empty:
                    st.caption("No result yet in this run.")
                else:
                    st.dataframe(this.iloc[::-1], use_container_width=True, hide_index=True)
                with st.expander("Full log of all runs"):
                    st.dataframe(rdf.iloc[::-1], use_container_width=True, hide_index=True)
                st.caption("DONE = signed. SKIPPED = already approved in Nexus, nothing done. "
                           "UNCERTAIN = Ok was clicked but not confirmed: check that dispatch in Nexus. "
                           "FAILED = nothing was signed (it is tried again next run).")

        shots = sorted(BASE.glob("debug_dispatched_failed*.png"), key=lambda p: p.stat().st_mtime, reverse=True)
        if shots and shots[0].stat().st_mtime >= st.session_state.get("d_started", float("inf")):
            with st.expander("Last failure screenshot"):
                st.image(str(shots[0]))

    if d_is_running():                            # auto-refresh while the robot is running
        time.sleep(1.5)
        st.rerun()


# ---------------------------------------------------------------- orders process (reopen + HAWB + approve + sign)
O_MODES = {                                    # key: (label shown in the UI, extra flags for orders_step.py)
    "full": ("Full run: reopen, HAWB, validate, approve and sign", []),
    "no_sign": ("Stop at the Sign pop-up (Ok is NOT clicked)", ["--no-sign"]),
    "no_approve": ("Stop before Approve (reopen, HAWB and validate only)", ["--no-approve"]),
    "dry": ("Stop before Reopen (nothing is changed in Nexus)", ["--dry-run"]),
}


def o_is_running():
    p = st.session_state.get("oproc")
    return p is not None and p.poll() is None


def read_orders_plan(path):
    """[(order, HAWB)] from the list file, using orders_step's own reader (1 order = 1 HAWB number)."""
    sys.path.insert(0, str(BASE))
    import importlib
    mod = importlib.import_module("orders_step")
    try:
        return mod.read_pairs(str(path))
    except SystemExit as e:                       # the reader exits with a message
        raise ValueError(str(e))


def _cell(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():          # Excel numbers come as 4502408413.0
        v = int(v)
    return str(v).strip().strip('"').strip("'")


def _read_columns(path, want):
    """Rows of {name: value} from the first sheet. `want` = {name: header regex}.
    The header row is searched in the first 30 rows (title rows / filters above it are fine)."""
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    rows = [[_cell(c) for c in r] for r in wb.worksheets[0].iter_rows(values_only=True)]
    for hi, cells in enumerate(rows[:30]):
        idx = {}
        for name, rx in want.items():
            for i, c in enumerate(cells):
                if c and re.search(rx, c, re.I):
                    idx[name] = i
                    break
        if len(idx) == len(want):
            return [{n: (r[i] if i < len(r) else "") for n, i in idx.items()} for r in rows[hi + 1:]]
    raise ValueError("Header not found in %s - needed columns: %s" % (Path(path).name, ", ".join(want.values())))


def read_two_excels(po_path, awb_path):
    """Excel 1 (INV.NO. + PO Number): the FIRST PO of each invoice.
    Excel 2 (Inv.no. + AWB no.): the FIRST AWB of each invoice.
    Matched by invoice number. Returns (pairs [(PO, AWB)], {PO: invoice}, [warnings])."""
    warns = []
    po_of = {}
    for r in _read_columns(po_path, {"inv": r"^\s*inv", "po": r"^\s*po\b"}):
        inv = r["inv"].upper()
        if inv and r["po"] and inv not in po_of:
            po_of[inv] = r["po"]
    awb_of, awb_all = {}, {}
    for r in _read_columns(awb_path, {"inv": r"^\s*inv", "awb": r"awb"}):
        inv = r["inv"].upper()
        if inv and r["awb"]:
            awb_of.setdefault(inv, r["awb"])
            awb_all.setdefault(inv, set()).add(r["awb"])
    if not po_of:
        raise ValueError("No invoice + PO rows found in the first Excel")
    for inv, s in awb_all.items():
        if len(s) > 1:
            warns.append(f"{inv} has {len(s)} different AWB numbers in Excel 2 ({', '.join(sorted(s))}) "
                         f"- the first one ({awb_of[inv]}) is used.")
    pairs, inv_of, seen = [], {}, set()
    for inv, po in po_of.items():
        if po in seen:
            warns.append(f"PO {po} appears for more than one invoice - only {inv_of[po]} is kept.")
            continue
        seen.add(po)
        inv_of[po] = inv
        pairs.append((po, awb_of.get(inv, "")))
    no_awb = [inv for inv in po_of if inv not in awb_of]
    if no_awb:
        warns.append(f"No AWB in Excel 2 for invoice: {', '.join(no_awb)} - these orders are not run.")
    only2 = [inv for inv in awb_of if inv not in po_of]
    if only2:
        warns.append(f"In Excel 2 but not in Excel 1 (ignored): {', '.join(only2)}")
    return pairs, inv_of, warns


def o_log_status():
    """{order: last status} from orders_results.csv."""
    status = {}
    if O_RESULTS_CSV.exists():
        try:
            df = pd.read_csv(O_RESULTS_CSV, dtype=str, keep_default_na=False, encoding="utf-8-sig")
            for o, s_ in zip(df["order"], df["status"]):
                status[o] = s_
        except Exception:
            pass
    return status


def o_read_log():
    if not O_LOG_FILE.exists():
        return ""
    return O_LOG_FILE.read_text(encoding="utf-8", errors="replace")


def start_orders(list_path, mode_key, opts):
    env = os.environ.copy()
    env.update({
        "NONINTERACTIVE": "1",                    # never wait for 'Press Enter'
        "HEADLESS": "true" if opts["headless"] else "false",
        "MAX_ATTEMPTS": str(opts["max_attempts"]),
        "RESULTS_FILE": str(O_RESULTS_CSV),
        "PYTHONUNBUFFERED": "1",
        "PYTHONIOENCODING": "utf-8",
    })
    cmd = [sys.executable, "-u", str(O_SCRIPT), "--batch", str(list_path)] + O_MODES[mode_key][1]
    if opts["redo"]:
        cmd.append("--redo")
    log = open(O_LOG_FILE, "w", encoding="utf-8")
    st.session_state["oproc"] = subprocess.Popen(
        cmd, cwd=str(BASE), env=env, stdout=log, stderr=subprocess.STDOUT,
        stdin=subprocess.DEVNULL,
        start_new_session=not IS_WINDOWS,         # lets Stop kill the browser too (Linux)
    )
    st.session_state["o_started"] = time.time()
    st.session_state["o_started_str"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    st.session_state["o_mode"] = mode_key


def orders_page():
    busy = o_is_running()
    other_busy = is_running() or d_is_running()   # the other robots (same Nexus session)
    left, right = st.columns([1, 1.4], gap="large")

    # ---------- left: list + options
    with left:
        st.subheader("1 · Order list")
        TWO = "Two Excel files (matched by invoice no.)"
        ONE = "One list (order, HAWB)"
        src_mode = st.radio("List source", [TWO, ONE], horizontal=True, key="o_srcmode",
                            disabled=busy or other_busy)
        two = src_mode == TWO
        up, pasted, src_key = None, "", None
        if two:
            up1 = st.file_uploader("Excel 1 - INV.NO. + PO Number (the first PO of each invoice is used)",
                                   type=["xlsx", "xlsm"], key="o_up1", disabled=busy or other_busy)
            up2 = st.file_uploader("Excel 2 - Inv.no. + AWB no. (the first AWB of each invoice is used)",
                                   type=["xlsx", "xlsm"], key="o_up2", disabled=busy or other_busy)
            if up1 is not None and up2 is not None:
                src_key = f"two-{up1.name}-{up1.size}-{up2.name}-{up2.size}"
                if st.session_state.get("o_src_key") != src_key:    # read only once, not on every refresh
                    UPLOAD_DIR.mkdir(exist_ok=True)
                    stamp = f"{datetime.now():%Y%m%d_%H%M%S}"
                    p1, p2 = UPLOAD_DIR / f"{stamp}_po_{up1.name}", UPLOAD_DIR / f"{stamp}_awb_{up2.name}"
                    p1.write_bytes(up1.getbuffer())
                    p2.write_bytes(up2.getbuffer())
                    try:
                        pairs, inv_of, warns = read_two_excels(str(p1), str(p2))
                        st.session_state.update(o_plan=pairs, o_inv=inv_of, o_warn=warns, o_error=None)
                    except Exception as e:
                        st.session_state.update(o_plan=[], o_inv={}, o_warn=[], o_error=str(e))
                    st.session_state["o_src_key"] = src_key
        else:
            up = st.file_uploader("Excel / CSV with an Order number column and a HAWB number column (1 order = 1 HAWB)",
                                  type=["xlsx", "xlsm", "csv", "txt"], key="o_upload", disabled=busy or other_busy)
            with st.expander("Or paste the list here (one line per order:  order, HAWB)"):
                pasted = st.text_area("order,hawb", key="o_paste", height=120, label_visibility="collapsed",
                                      placeholder="4502408413,6634365323", disabled=busy or other_busy)
            if up is not None:
                src_key = f"file-{up.name}-{up.size}"
            elif pasted.strip():
                src_key = "paste-" + hashlib.md5(pasted.encode("utf-8")).hexdigest()
        if src_key is None:
            for k in ("o_src_key", "o_plan", "o_error", "o_path"):
                st.session_state.pop(k, None)
        elif not two and st.session_state.get("o_src_key") != src_key:      # read the list only once, not on every refresh
            UPLOAD_DIR.mkdir(exist_ok=True)
            stamp = f"{datetime.now():%Y%m%d_%H%M%S}"
            if up is not None:
                saved = UPLOAD_DIR / f"{stamp}_{up.name}"
                saved.write_bytes(up.getbuffer())
            else:
                saved = UPLOAD_DIR / f"{stamp}_orders_pasted.txt"
                saved.write_text(pasted, encoding="utf-8")
            try:
                st.session_state["o_plan"] = read_orders_plan(saved)
                st.session_state["o_error"] = None
            except Exception as e:
                st.session_state["o_plan"] = []
                st.session_state["o_error"] = str(e)
            st.session_state["o_src_key"] = src_key

        plan = st.session_state.get("o_plan", [])
        status = o_log_status()
        hawb_of = dict(plan)
        chosen = []
        if st.session_state.get("o_error"):
            st.error(f"Could not read the list: {st.session_state['o_error']}")
        elif plan:
            st.success(f"{len(plan)} orders found (repeats and blanks removed)")
            missing = [o for o, h in plan if not h]
            if missing and not two:                      # in two-file mode the warnings below already say it
                st.warning(f"No HAWB number for: {', '.join(missing)} - these are not run (marked NO HAWB).")
            if two:
                for w in st.session_state.get("o_warn", []):
                    st.warning(w)
            inv_of = st.session_state.get("o_inv", {}) if two else {}
            cols = {}
            if two:
                cols["invoice"] = [inv_of.get(o, "") for o, _ in plan]
            cols["order (PO)"] = [o for o, _ in plan]
            cols["HAWB (AWB)"] = [h or "(missing)" for _, h in plan]
            cols["last status"] = [status.get(o, "") for o, _ in plan]
            st.dataframe(pd.DataFrame(cols), hide_index=True, use_container_width=True)
            chosen = st.multiselect("Orders to run", [o for o, _ in plan],
                                    default=[o for o, _ in plan], disabled=busy)
        elif src_key is None:
            st.caption("Excel: columns are found by their headers (Order ... / HAWB ...); without headers "
                       "column A = order and column B = HAWB. CSV / text: `order,hawb` per line.")

        st.subheader("2 · Options")
        mode_key = st.radio("What should the robot do?", list(O_MODES), format_func=lambda k: O_MODES[k][0],
                            disabled=busy)
        skip_done = st.checkbox("Skip orders already DONE in the results log", value=True, disabled=busy)
        max_attempts = st.number_input("Attempts per order (rebuild on error)", 1, 10, 3, 1, disabled=busy)
        with st.expander("More"):
            headless = st.checkbox("Headless browser", value=False, disabled=busy)
            if not IS_WINDOWS:
                headless = True                   # a server has no screen

        run_list = [o for o in chosen if hawb_of.get(o)]
        to_run = [o for o in run_list if not (skip_done and mode_key == "full" and status.get(o) == "DONE")]

        st.subheader("3 · Run")
        needs_confirm = mode_key != "dry"
        if mode_key == "full":
            st.warning("This robot **reopens, approves and signs** packing lists in Infor Nexus. "
                       "Try one order first, or use *Stop at the Sign pop-up*.")
        elif needs_confirm:
            st.warning("This robot **reopens** the packing lists and types the HAWB number in Infor Nexus.")
        if run_list and not to_run:
            st.info("Every selected order is already DONE in the results log. Untick the skip option to run them again.")
        confirmed = (not needs_confirm) or st.checkbox(
            "I understand the selected orders will be changed in Infor Nexus", value=False, disabled=busy)
        if other_busy:
            st.caption("Another robot is running - wait for it to finish (they share the same Nexus session).")
        c1, c2 = st.columns(2)
        if c1.button("▶ Start orders", type="primary", key="o_start", use_container_width=True,
                     disabled=busy or other_busy or not to_run or not confirmed):
            UPLOAD_DIR.mkdir(exist_ok=True)
            run_file = UPLOAD_DIR / f"orders_run_{datetime.now():%Y%m%d_%H%M%S}.csv"
            run_file.write_text("order,hawb\n" + "\n".join(f"{o},{hawb_of[o]}" for o in run_list) + "\n",
                                encoding="utf-8")
            start_orders(run_file, mode_key, dict(redo=not skip_done, headless=headless,
                                                  max_attempts=int(max_attempts)))
            st.rerun()
        if c2.button("■ Stop", key="o_stop", disabled=not busy, use_container_width=True):
            kill_robot(st.session_state["oproc"])
            time.sleep(1)
            st.rerun()

    # ---------- right: progress, results
    with right:
        st.subheader("Progress")
        log = o_read_log()
        cur = None
        for m in re.finditer(r"# Order (\d+)/(\d+) : (\S+)", log):
            cur = m
        if cur:
            n, total, order = int(cur.group(1)), int(cur.group(2)), cur.group(3)
            done_n = n - 1 if busy else n
            st.progress(min(done_n / total, 1.0) if total else 0.0,
                        text=f"{'Working on' if busy else 'Last'}: {order}  ({n}/{total})")
        if busy:
            st.info("🟢 Orders robot is running…")
        elif st.session_state.get("oproc") is not None:
            code = st.session_state["oproc"].returncode
            (st.success if code == 0 else st.error)(
                "Finished" if code == 0 else f"Finished with problems or stopped (exit code {code}) - see Results and the log")

        with st.expander("Live log", expanded=busy):
            st.code("\n".join(log.splitlines()[-40:]) or "(no log yet)", language="text")

        if O_RESULTS_CSV.exists():
            try:
                rdf = pd.read_csv(O_RESULTS_CSV, dtype=str, keep_default_na=False, encoding="utf-8-sig")
            except Exception:
                rdf = pd.DataFrame()
            if not rdf.empty:
                started = st.session_state.get("o_started_str")
                st.subheader("Results (this run)")
                this = rdf[rdf["time"] >= started] if started else rdf.iloc[0:0]
                if this.empty:
                    st.caption("No result yet in this run.")
                else:
                    st.dataframe(this.iloc[::-1], use_container_width=True, hide_index=True)
                with st.expander("Full log of all runs"):
                    st.dataframe(rdf.iloc[::-1], use_container_width=True, hide_index=True)
                st.caption("DONE = reopened, HAWB typed, approved and signed. TEST = a test run (stopped on purpose). "
                           "CHECK = it failed AFTER Approve was clicked: open that order in Nexus and look at it by hand "
                           "(it is never repeated automatically). FAILED = it did not get to Approve. "
                           "NO HAWB = the list has no HAWB number for that order.")

        shots = sorted(BASE.glob("debug_orders_failed*.png"), key=lambda p: p.stat().st_mtime, reverse=True)
        if shots and shots[0].stat().st_mtime >= st.session_state.get("o_started", float("inf")):
            with st.expander("Last failure screenshot"):
                st.image(str(shots[0]))

    if o_is_running():                            # auto-refresh while the robot is running
        time.sleep(1.5)
        st.rerun()


# ---------------------------------------------------------------- page
MODE_PDF, MODE_DISPATCH = "🧾 Invoice PDF", "🚚 Dispatch (validate + approve)"
MODE_ORDERS = "📦 Orders (reopen + HAWB + approve)"
mode = st.radio("Process", [MODE_PDF, MODE_DISPATCH, MODE_ORDERS], horizontal=True, label_visibility="collapsed")
if mode == MODE_DISPATCH:
    st.title("🚚 Nexus Dispatch Robot")
    st.caption("Upload the Excel (INV.NO. + PO Number) → for each invoice the robot searches one PO in "
               "Infor Nexus Dispatches, fills the invoice number, then Preview → Validate → Approve → Ok.")
elif mode == MODE_ORDERS:
    st.title("📦 Nexus Orders Robot")
    st.caption("Upload the list (Order + HAWB, 1 order = 1 HAWB) → for each order the robot opens the Packing List, "
               "Reopen → types the HAWB number → Preview → Validate → Approve → Sign Ok.")
else:
    st.title("🧾 Nexus Invoice PDF Robot")
    st.caption("Upload the Excel list → the robot opens each invoice in Infor Nexus and saves "
               "`<invoice>-INV.pdf` into `output_pdfs`.")

env_vals = dotenv_values(BASE / ".env")
if mode == MODE_ORDERS:
    if not ((env_vals.get("NEXUS_USER_2") or os.getenv("NEXUS_USER_2"))
            and (env_vals.get("NEXUS_PASS_2") or os.getenv("NEXUS_PASS_2"))):
        st.error("NEXUS_USER_2 / NEXUS_PASS_2 are missing (`.env` next to this app, or Secrets on Streamlit Cloud).")
elif not ((env_vals.get("NEXUS_USER") or os.getenv("NEXUS_USER"))
          and (env_vals.get("NEXUS_PASS") or os.getenv("NEXUS_PASS"))):
    st.error("NEXUS_USER / NEXUS_PASS are missing (`.env` next to this app, or Secrets on Streamlit Cloud).")

if mode == MODE_DISPATCH:
    dispatch_page()
    st.stop()
if mode == MODE_ORDERS:
    orders_page()
    st.stop()
if d_is_running() or o_is_running():
    st.warning("🚚 The Dispatch / Orders robot is running - Start is disabled here until it finishes "
               "(the robots share the same Nexus session).")

left, right = st.columns([1, 1.4], gap="large")

# ---------- left: upload + options
with left:
    st.subheader("1 · Upload Excel")
    up = st.file_uploader("Excel with an INV.NO. column", type=["xlsx", "xlsm"],
                          disabled=is_running())
    invoices = st.session_state.get("invoices", [])
    if up is not None:
        key = f"{up.name}-{up.size}"
        if st.session_state.get("up_key") != key:          # read the file only once, not on every refresh
            UPLOAD_DIR.mkdir(exist_ok=True)
            saved = UPLOAD_DIR / f"{datetime.now():%Y%m%d_%H%M%S}_{up.name}"
            saved.write_bytes(up.getbuffer())
            try:
                st.session_state["invoices"] = read_invoices(str(saved))
                st.session_state["inv_dates"] = read_dates(str(saved))
                st.session_state["up_error"] = None
            except Exception as e:
                st.session_state["invoices"] = []
                st.session_state["up_error"] = str(e)
            st.session_state["up_key"] = key
        invoices = st.session_state["invoices"]
        if st.session_state.get("up_error"):
            st.error(f"Could not read the file: {st.session_state['up_error']}")
        elif invoices:
            st.success(f"{len(invoices)} unique invoices found (repeats and blanks removed)")
            dts = st.session_state.get("inv_dates", {})
            missing = [i for i in invoices if not dts.get(i)]
            conf = st.session_state.get("date_conflicts", {})
            for i, ds in conf.items():
                if i in invoices:
                    st.warning(f"{i} has different dates on different rows "
                               f"({', '.join(d.strftime('%d %b %Y') for d in ds)}) - it will be skipped as FAILED.")
            missing = [i for i in missing if i not in conf]
            if missing:
                st.warning(f"No readable date in the Excel for: {', '.join(missing)} "
                           "- these will be marked FAILED and not touched.")
            with st.expander("Dates read from the Excel"):
                st.dataframe(pd.DataFrame({
                    "invoice": invoices,
                    "date": [dts[i].strftime("%d %b %Y") if dts.get(i) else "(missing)" for i in invoices]}),
                    hide_index=True, use_container_width=True)
    else:
        st.session_state.pop("up_key", None)
        invoices = []

    chosen = []
    if invoices:
        chosen = st.multiselect("Invoices to run", invoices, default=invoices,
                                disabled=is_running())

    st.subheader("2 · Options")
    print_mode = st.radio(
        "PDF method", ["gui", "direct"] if IS_WINDOWS else ["direct"], horizontal=True, disabled=is_running(),
        format_func=lambda x: "Print dialog (1-2-3-4, needs the screen)" if x == "gui"
        else "Direct capture (no screen clicks)")
    skip_existing = st.checkbox("Skip invoices whose PDF already exists", value=True,
                                disabled=is_running())
    preview_wait = st.number_input("Wait after Preview (ms)", 1000, 30000, 6000, 500,
                                   disabled=is_running())
    with st.expander("More"):
        headless = st.checkbox("Headless browser (direct method only)", value=False,
                               disabled=is_running() or print_mode == "gui")
        if not IS_WINDOWS:
            headless = True                       # a server has no screen
        doc_date = None                              # dates now come from the Excel

    opts = dict(print_mode=print_mode, skip_existing=skip_existing, preview_wait=int(preview_wait),
                headless=headless, doc_date=doc_date)

    st.subheader("3 · Run")
    if print_mode == "gui":
        st.warning("While it runs the robot browser takes over the screen. "
                   "**Don't touch the mouse or keyboard** until it finishes.")
    c1, c2 = st.columns(2)
    if c1.button("▶ Start", type="primary", disabled=is_running() or d_is_running() or o_is_running() or not chosen, use_container_width=True):
        start_robot(chosen, opts)
        st.rerun()
    if c2.button("■ Stop", disabled=not is_running(), use_container_width=True):
        kill_robot(st.session_state["proc"])      # robot + its browser
        time.sleep(1)
        st.rerun()

# ---------- right: progress, results, files
with right:
    st.subheader("Progress")
    log = read_log()
    running = is_running()

    done_n = 0
    cur = None
    for m in re.finditer(r"===== \[(\d+)/(\d+)\] (\S+) =====", log):
        cur = m
    total = st.session_state.get("total", 0)
    if cur:
        n, total, inv = int(cur.group(1)), int(cur.group(2)), cur.group(3)
        done_n = n - 1 if running else n
        st.progress(min(done_n / total, 1.0) if total else 0.0,
                    text=f"{'Working on' if running else 'Last'}: {inv}  ({n}/{total})")
    if running:
        st.info("🟢 Robot is running…")
    elif st.session_state.get("proc") is not None:
        code = st.session_state["proc"].returncode
        (st.success if code == 0 else st.error)(
            "Finished" if code == 0 else f"Stopped (exit code {code}) - see the log below")

    with st.expander("Live log", expanded=running):
        st.code("\n".join(log.splitlines()[-40:]) or "(no log yet)", language="text")

    if RESULTS_CSV.exists():
        st.subheader("Results")
        df = pd.read_csv(RESULTS_CSV, encoding="utf-8-sig")
        st.dataframe(df, use_container_width=True, hide_index=True)

    st.subheader("Saved PDFs")
    pdfs = sorted(OUTPUT_DIR.glob("*-INV.pdf"), key=lambda p: p.stat().st_mtime, reverse=True) \
        if OUTPUT_DIR.exists() else []
    if pdfs:
        if IS_WINDOWS and st.button("📂 Open output folder"):
            os.startfile(str(OUTPUT_DIR))          # Windows
        for p in pdfs[:30]:
            a, b = st.columns([3, 1])
            a.write(f"📄 {p.name}  ·  {p.stat().st_size / 1024:,.0f} KB  ·  "
                    f"{datetime.fromtimestamp(p.stat().st_mtime):%H:%M}")
            b.download_button("Download", p.read_bytes(), file_name=p.name,
                              mime="application/pdf", key=f"dl_{p.name}")
    else:
        st.caption("No PDFs yet.")

# auto-refresh while the robot is running
if is_running():
    time.sleep(1.5)
    st.rerun()

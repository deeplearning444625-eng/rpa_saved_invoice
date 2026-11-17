"""
Infor Nexus - Invoice PDF Robot (Streamlit UI)

Run on the SAME Windows PC that has the screen (the robot opens a browser and
clicks Chrome's print dialog on that screen):

    pip install streamlit pandas openpyxl playwright python-dotenv pyautogui opencv-python pillow
    streamlit run app_streamlit.py

Files needed in the same folder: infornexus_step1_login.py, infornexus_step2_menu.py,
print_gui.py, .env (NEXUS_USER / NEXUS_PASS).
For the Dispatch process (sidebar) also: dispatched_step.py, dispatched_loop.py.
"""
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

st.set_page_config(page_title="Nexus Invoice PDF Robot", page_icon="🧾", layout="wide")

# On Streamlit Cloud the credentials come from Secrets instead of a .env file
try:
    for _k in ("NEXUS_USER", "NEXUS_PASS", "NEXUS_URL", "NEXUS_ACCESS_CODE"):
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
    other_busy = is_running()                     # the PDF robot
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
            st.caption("The Invoice PDF robot is running - wait for it to finish.")
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


# ---------------------------------------------------------------- page
MODE_PDF, MODE_DISPATCH = "🧾 Invoice PDF", "🚚 Dispatch (validate + approve)"
mode = st.radio("Process", [MODE_PDF, MODE_DISPATCH], horizontal=True, label_visibility="collapsed")
if mode == MODE_DISPATCH:
    st.title("🚚 Nexus Dispatch Robot")
    st.caption("Upload the Excel (INV.NO. + PO Number) → for each invoice the robot searches one PO in "
               "Infor Nexus Dispatches, fills the invoice number, then Preview → Validate → Approve → Ok.")
else:
    st.title("🧾 Nexus Invoice PDF Robot")
    st.caption("Upload the Excel list → the robot opens each invoice in Infor Nexus and saves "
               "`<invoice>-INV.pdf` into `output_pdfs`.")

env_vals = dotenv_values(BASE / ".env")
if not ((env_vals.get("NEXUS_USER") or os.getenv("NEXUS_USER"))
        and (env_vals.get("NEXUS_PASS") or os.getenv("NEXUS_PASS"))):
    st.error("NEXUS_USER / NEXUS_PASS are missing (`.env` next to this app, or Secrets on Streamlit Cloud).")

if mode == MODE_DISPATCH:
    dispatch_page()
    st.stop()
if d_is_running():
    st.warning("🚚 The Dispatch robot is running - Start is disabled here until it finishes "
               "(both robots share the same Nexus session).")

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
    if c1.button("▶ Start", type="primary", disabled=is_running() or d_is_running() or not chosen, use_container_width=True):
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

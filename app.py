"""Smart Form Filler: upload Excel, record one entry on any website, link values to columns, replay for every row."""
import difflib
import json
import os
import re
import time
from datetime import datetime

import pandas as pd
import streamlit as st

import engine

HERE = os.path.dirname(os.path.abspath(__file__))
RECIPES = os.path.join(HERE, "recipes")
RESULTS = os.path.join(HERE, "results")
DONE_STATUSES = {"Registered", "Already exists", "Submitted"}
VALUE_ACTIONS = ("fill", "select", "check")

COPY = f" - copy {engine.SLOT}" if engine.SLOT > 1 else ""  # several copies can run at once, one per website
st.set_page_config(page_title="Smart Form Filler" + COPY, page_icon="📝", layout="wide")
ss = st.session_state
for k, v in {"steps": None, "start_url": "", "recorder": None, "results": None}.items():
    ss.setdefault(k, v)


def clear_widgets():
    for k in [k for k in ss if str(k).startswith(("use_", "src_", "fmt_", "val_", "when_", "wait_"))]:
        del ss[k]


def norm(s):
    return re.sub(r"[^a-z0-9]", "", str(s).lower())


SYNONYMS = {
    "dob": "dateofbirth", "birthdate": "dateofbirth", "birthday": "dateofbirth", "dateofbirthddmmyyyy": "dateofbirth",
    "surname": "lastname", "familyname": "lastname", "lname": "lastname",
    "forename": "firstname", "givenname": "firstname", "fname": "firstname",
    "secondname": "middlename", "mname": "middlename",
    "sex": "gender", "emailaddress": "email", "mail": "email",
    "mobile": "phone", "mobileno": "phone", "phoneno": "phone", "telephone": "phone", "contact": "phone",
    "nationalityname": "nationality", "countryname": "country", "zip": "postcode", "zipcode": "postcode",
    "city": "towncity", "town": "towncity",
}


def canon(s):
    """Normalised name with brackets dropped and synonyms unified: 'DOB (dd.mm.yyyy)' -> 'dateofbirth'."""
    n = norm(re.sub(r"\(.*?\)", "", str(s)))
    return SYNONYMS.get(n, n)


def looks_date(name):
    return "date" in canon(name) or "birth" in norm(name)


def guess_source(step, df, columns):
    """(column, date format) an Excel column probably feeds this step, or (None, None).

    Best clue: the value typed while recording equals that column's value in one of the first rows.
    """
    v = str(step.get("value", "")).strip().lower()
    if v and step["action"] != "check":
        for col in columns:
            for cell in df[col].head(5):
                for fmt in engine.DATE_FORMATS:
                    if engine.to_text(cell, fmt).lower() == v:
                        return col, (None if fmt == "as is" else fmt)
    label = canon(step.get("label", ""))
    if label:
        keyed = {canon(c): c for c in columns}
        hit = keyed.get(label) or next(iter(difflib.get_close_matches(label, list(keyed), n=1, cutoff=0.8)), None)
        hit = keyed.get(hit, hit)
        if hit:
            return hit, ("dd.mm.yyyy" if looks_date(hit) and step["action"] == "fill" else None)
    return None, None


def save_results(df, src_name):
    os.makedirs(RESULTS, exist_ok=True)
    # Re-running a results file keeps one "_result" name instead of "_result_result".
    base = re.sub(r"(_result(_copy\d+)?(_\d{6})?)+$", "", os.path.splitext(src_name)[0])
    tag = f"_copy{engine.SLOT}" if engine.SLOT > 1 else ""  # two copies never write the same file
    path = os.path.join(RESULTS, f"{base}_result{tag}.xlsx")
    try:
        df.to_excel(path, index=False)
    except PermissionError:  # file open in Excel
        path = os.path.join(RESULTS, f"{base}_result{tag}_{datetime.now():%H%M%S}.xlsx")
        df.to_excel(path, index=False)
    return path


# ============================================================================ sidebar
with st.sidebar:
    st.header("Browser" + COPY)
    url = st.text_input("Website to open", "https://example.com/login")
    if st.button("Open browser", type="primary", width="stretch"):
        if engine.is_connected():
            if url.strip():
                engine.open_url(url.strip())
                st.success("Opened in a new tab of the tool's browser.")
            else:
                st.info("Browser is already open.")
        else:
            try:
                engine.launch_browser(url)
                st.success("Opened. Log in and go to the page where you add records.")
            except Exception as e:
                st.error(str(e))
    st.caption("This opens a separate Chrome window just for this tool. Log in there "
               "(the login is remembered next time) and open the page where you add records.")
    if engine.is_connected():
        st.success("Browser connected")
    else:
        st.error("Browser not connected")
    st.button("Refresh status", width="stretch")

    st.divider()
    st.header("Saved websites")
    os.makedirs(RECIPES, exist_ok=True)
    saved = sorted(f[:-5] for f in os.listdir(RECIPES) if f.endswith(".json"))
    if saved:
        choice = st.selectbox("Setup", saved)
        if st.button("Load setup", width="stretch"):
            with open(os.path.join(RECIPES, choice + ".json"), encoding="utf-8") as fh:
                r = json.load(fh)
            if r.get("version") != 2:
                st.error("This setup was made with an older version. Please record it again.")
            else:
                ss.steps, ss.start_url, ss.loaded, ss.recipe_name = r["steps"], r.get("start_url", ""), r, choice
                clear_widgets()
                st.success(f"Loaded '{choice}'")
    else:
        st.caption("Nothing saved yet. Finish step 5 and click **Save setup**.")

st.title("📝 Smart Form Filler" + COPY)
st.caption("Upload Excel → choose columns → record yourself adding one record → it repeats that for every row.")

# ============================================================================ step 1
st.header("1. Upload Excel")
up = st.file_uploader("Excel or CSV file", type=["xlsx", "xlsm", "csv"])
if not up:
    st.stop()

c1, c2 = st.columns(2)
if up.name.lower().endswith(".csv"):
    sheet = None
else:
    xl = pd.ExcelFile(up)
    sheet = c1.selectbox("Sheet", xl.sheet_names) if len(xl.sheet_names) > 1 else xl.sheet_names[0]
header_row = c2.number_input("Header row (the row with column names)", 1, 50, 1)
up.seek(0)
if sheet is None:
    df = pd.read_csv(up, header=header_row - 1, dtype=object)
else:
    df = pd.read_excel(up, sheet_name=sheet, header=header_row - 1, dtype=object)
df = df.dropna(how="all").reset_index(drop=True)
df.columns = [str(c).strip() for c in df.columns]
first_excel_row = header_row + 1
st.write(f"**{len(df)} rows**, {len(df.columns)} columns")
st.dataframe(df.head(10), width="stretch")

# ============================================================================ step 2
st.header("2. Choose the columns to use")
loaded_cols = [s["source"]["column"] for s in (ss.steps or []) if (s.get("source") or {}).get("column")]
data_cols = [c for c in df.columns if c not in ("Status", "Message", "Processed at")]
selected = st.multiselect("Columns", data_cols, default=[c for c in dict.fromkeys(loaded_cols) if c in data_cols],
                          help="Only these columns will be typed into the website.")
if not selected:
    st.info("Pick at least one column.")
    st.stop()

# ============================================================================ step 3
st.header("3. Record yourself adding one record")
rec = ss.recorder
if rec is not None and rec.running:
    st.error("⏺ **Recording.** In the tool's browser, add **one** record exactly as you normally do. "
             "When it is saved on the website, click **Stop recording** here.")

    @st.fragment(run_every=1.5)
    def live_steps():
        steps = rec.snapshot()
        if not steps:
            st.caption("Waiting for your first click…")
        for i, s in enumerate(steps, 1):
            val = f" → `{s['value']}`" if s.get("value") else ""
            st.markdown(f"{i}. {engine.describe(s)}{val}")

    live_steps()
    if st.button("⏹ Stop recording", type="primary"):
        ss.steps, ss.start_url = rec.stop(), rec.start_url
        ss.recorder = None
        ss.pop("loaded", None)
        clear_widgets()
        st.rerun()
    st.stop()

if rec is not None and rec.error:
    st.error(f"Recording stopped: {rec.error}")
    ss.recorder = None
st.markdown(
    "1. In the tool's browser, go to the page where you **start** adding a record "
    "(logged in, cohort selected, etc.).\n"
    "2. Click **Start recording**, then add **one** record there by hand, all the way to saving it. "
    "Every click, typed value and dropdown choice is recorded.\n"
    "3. Click **Stop recording**.\n\n"
    "**Tips:** type the details of an Excel row while recording (e.g. the first one below) - the tool then "
    "links each typed value to its column by itself. Start from a page you can always get back to, and include "
    "any choice that page needs (e.g. picking the cohort) in your recording."
)
st.dataframe(df[selected].head(1), width="stretch", hide_index=True)
label = "⏺ Record again" if ss.steps else "⏺ Start recording"
if st.button(label, type="secondary" if ss.steps else "primary"):
    if not engine.is_connected():
        st.error("Open the browser first (left panel).")
    else:
        try:
            ss.recorder = engine.Recorder().start()
            st.rerun()
        except Exception as e:
            st.error(f"Could not start recording: {e}")
if not ss.steps:
    st.stop()

# ============================================================================ step 4
st.header("4. Check the steps and link your Excel columns")
st.caption("Untick any step you didn't mean to record. Choose **when** each step happens: once before the first "
           "row, for every row, or once after the last row. For each typed or chosen value, pick the Excel column it "
           "should come from - or keep the recorded value to use the same one for everyone.")
with st.expander("Example: add every student in one form, then Close and Save at the end"):
    st.markdown("Record one full pass: *Create New Candidate* → fill the form → *Create* → *Add another candidate* "
                "→ *Close* → *Save*. Then set:\n"
                "- *Create New Candidate* → **Once at start**\n"
                "- the fields, *Create*, *Add another candidate* → **Every row**\n"
                "- *Close*, *Save* → **Once at end**\n\n"
                "If a record goes wrong and the form gets closed, the tool redoes the *Once at start* steps "
                "to open it again before the next row.")
WHEN = {"Every row": "row", "Once at start": "start", "Once at end": "end"}
plan = []
loaded_sources = {s["id"]: s.get("source") for s in ss.steps} if ss.get("loaded") else {}
COLS = [0.5, 4.2, 2, 3.3, 1.5, 1.5]
head = st.columns(COLS)
for col, title in zip(head[1:], ["Step", "When", "Value from", "Format", "Wait after (s)"]):
    col.caption(title)
for n, s in enumerate(ss.steps, 1):
    s = dict(s)
    if s["action"] == "dialog":
        st.caption(f"{n}. {engine.describe(s)}")
        continue
    row = st.columns(COLS)
    s["use"] = row[0].checkbox("use", value=s.get("use", not s.get("secret")), key=f"use_{s['id']}",
                               label_visibility="collapsed")
    text = f"**{n}.** {engine.describe(s)}"
    if s.get("secret"):
        text += "  \n<small>Passwords are not recorded. Log in before recording instead.</small>"
    row[1].markdown(text, unsafe_allow_html=True)
    when_names = list(WHEN)
    when = row[2].selectbox("when", when_names, key=f"when_{s['id']}", label_visibility="collapsed",
                            index=list(WHEN.values()).index(s.get("when", "row")))
    s["when"] = WHEN[when]
    s["source"] = None
    if s["action"] in VALUE_ACTIONS and not s.get("secret") and s["when"] == "row":
        fixed = f"Always: '{s.get('value', '')}'"
        choices = [fixed] + selected
        prev = loaded_sources.get(s["id"])
        col, fmt = (prev["column"], prev.get("date_format")) if prev and prev.get("column") in selected \
            else guess_source(s, df, selected)
        pick = row[3].selectbox("value from", choices, index=choices.index(col) if col in choices else 0,
                                key=f"src_{s['id']}", label_visibility="collapsed")
        if pick != fixed:
            fmt_choice = None
            if s["action"] == "fill" and s.get("kind") != "date":
                fmts = list(engine.DATE_FORMATS)
                f = row[4].selectbox("format", fmts, index=fmts.index(fmt or "as is"), key=f"fmt_{s['id']}",
                                     label_visibility="collapsed", help="Only matters for dates")
                fmt_choice = None if f == "as is" else f
            s["source"] = {"column": pick, "date_format": fmt_choice}
    s["wait"] = row[5].number_input("wait", 0.0, 600.0, value=s.get("wait"), step=0.5, key=f"wait_{s['id']}",
                                    label_visibility="collapsed", placeholder="default",
                                    help="Seconds to wait after this step. Empty = the general "
                                         "'Pause between steps' in step 6.")
    plan.append(s)

used = [s for s in plan if s["use"]]
linked = [s for s in used if s["source"]]
groups = {w: [s for s in used if s["when"] == w] for w in ("start", "row", "end")}
st.info(f"Once at start: **{len(groups['start'])}** step(s) · Every row: **{len(groups['row'])}** step(s) · "
        f"Once at end: **{len(groups['end'])}** step(s)")
if not groups["row"]:
    st.error("At least one step must be done for every row.")
    st.stop()
if not linked:
    st.warning("No step takes its value from Excel yet, so every record would be identical.")

# Dropdowns only accept their own wording, so translate each Excel value to an option.
NO_MATCH = "✗ no match (row will fail)"
value_map = {}
loaded_values = ss.get("loaded", {}).get("value_map", {})
for s in [s for s in linked if s["action"] == "select" and s.get("options")]:
    col, key = s["source"]["column"], str(s["id"])
    options = s["options"]
    uniq = sorted({engine.to_text(v) for v in df[col]} - {""})[:200]
    as_opts = [{"t": o, "v": o} for o in options]
    guesses = {}
    for v in uniq:
        saved = loaded_values.get(key, {}).get(v)
        hit = engine.pick_option(as_opts, v)
        guesses[v] = saved if saved in options else (hit["t"] if hit else NO_MATCH)
    missing = sum(g == NO_MATCH for g in guesses.values())
    title = f"Dropdown values for **{col}** → {s.get('label')}: {len(uniq) - missing}/{len(uniq)} matched"
    with st.expander(("⚠️ " if missing else "✅ ") + title, expanded=bool(missing)):
        st.caption("Each different value in your Excel column, and the dropdown option it will choose.")
        value_map[key] = {}
        grid = st.columns(2)
        for i, v in enumerate(uniq):
            choices = [NO_MATCH] + options
            pick = grid[i % 2].selectbox(f"Excel: {v}", choices, index=choices.index(guesses[v]),
                                         key=f"val_{key}_{v}")
            if pick != NO_MATCH:
                value_map[key][v] = pick

# ============================================================================ step 5
st.header("5. How do I check each record worked?")
st.caption("After each record, the tool checks that a **new** table row appeared on the website with exactly these "
           "values (instead of you pressing Ctrl+F). Use name **and** date of birth so two people with the same name "
           "are never confused. Leave empty if the website shows no list.")
linked_cols = list(dict.fromkeys(s["source"]["column"] for s in linked))
verify_default = [c for c in ss.get("loaded", {}).get("verify_columns", []) if c in linked_cols] or \
                 [c for c in linked_cols if canon(c) in ("firstname", "lastname", "dateofbirth")]
verify_cols = st.multiselect("Check columns", linked_cols, default=verify_default)
fmt_of = {s["source"]["column"]: s["source"]["date_format"] for s in reversed(linked)}
if verify_cols:
    keys = df[verify_cols].apply(lambda c: c.map(lambda v: engine.to_text(v, fmt_of.get(c.name)).lower())).agg("|".join, axis=1)
    dup = keys[keys.duplicated(keep=False) & (keys.str.strip("|") != "")]
    if len(dup):
        rows_txt = ", ".join(str(i + first_excel_row) for i in dup.index[:30])
        st.warning(f"**{len(dup)} Excel rows look like the same person** (same {', '.join(verify_cols)}): rows {rows_txt}. "
                   "After the first one is created, the others will be marked *Already exists* and skipped. "
                   "If they really are different people, add another check column or untick "
                   "*Skip records already in the table*.")

c1, c2 = st.columns([3, 1])
name = c1.text_input("Save this setup as", ss.get("recipe_name", "my-website"))
c2.write("")
if c2.button("Save setup"):
    rec_out = {"version": 2, "start_url": ss.start_url, "steps": plan, "verify_columns": verify_cols,
               "value_map": value_map}
    safe_name = re.sub(r'[\\/:*?"<>|]', "_", name.strip()) or "setup"
    with open(os.path.join(RECIPES, safe_name + ".json"), "w", encoding="utf-8") as fh:
        json.dump(rec_out, fh, indent=2)
    ss.recipe_name = safe_name
    st.success(f"Saved. Next time, load '{safe_name}' from the left panel - no need to record again.")


def row_values(r):
    """{step id: text} for this Excel row, plus the values used to find it in the website's table."""
    vals = {}
    for s in linked:
        text = engine.to_text(df.at[r, s["source"]["column"]], s["source"]["date_format"])
        vals[s["id"]] = value_map.get(str(s["id"]), {}).get(text, text)
    verify = [engine.to_text(df.at[r, c], fmt_of.get(c)) for c in verify_cols]
    return vals, [v for v in verify if v]


# ============================================================================ step 6
st.header("6. Run")
st.caption(f"Row numbers are Excel row numbers (data starts at row {first_excel_row}). "
           "Before running, have the browser on the page where your recording started.")

with st.expander("🧪 Test: do one row without saving it", expanded=True):
    st.caption("Does the steps for one row but stops before the click that saves, so you can check the form. "
               "Nothing is saved.")
    c1, c2 = st.columns([1, 3])
    test_row = c1.number_input("Excel row", first_excel_row, first_excel_row + len(df) - 1, first_excel_row)
    test_clicked = c2.button("Fill this row (don't save)")
    test_box = st.empty()

c1, c2 = st.columns(2)
step_pause = c1.number_input("Pause between steps (seconds)", 0.0, 10.0, 1.0, 0.5,
                             help="Wait this long after each field or click, e.g. type first name, wait, type last name.")
slow_typing = c2.checkbox("Type letter by letter, like a person", False,
                          help="Slower. Use it if the website reacts to each key (e.g. auto-fills other boxes).")
c1, c2, c3, c4 = st.columns(4)
row_from = c1.number_input("From Excel row", first_excel_row, first_excel_row + len(df) - 1, first_excel_row)
row_to = c2.number_input("To Excel row", first_excel_row, first_excel_row + len(df) - 1, first_excel_row + len(df) - 1)
delay = c3.number_input("Wait after each record (seconds)", 0.5, 30.0, 1.5, 0.5)
timeout = c4.number_input("Give up waiting after (seconds)", 5, 120, 20)
c1, c2 = st.columns(2)
skip_existing = c1.checkbox("Skip records already in the table", True, disabled=not verify_cols)
resume = c2.checkbox("Skip rows already marked done in a Status column", "Status" in df.columns,
                     disabled="Status" not in df.columns)
on_fail = st.radio("When a row has a data error (e.g. a value the website rejects)",
                   ["Mark it Failed and continue to the end", "Stop so I can look at it"], horizontal=True)
stop_on_fail = on_fail.startswith("Stop")
auto_home = st.checkbox(
    f"If a record leaves the page somewhere else, go back to the start page automatically ({ss.start_url[:70]})",
    value=True, disabled=not ss.start_url.startswith("http"),
    help="Safe when your recording starts from a page you can reopen by its address and includes everything "
         "needed there (e.g. choosing the cohort). Otherwise untick it: the run then pauses for you to fix the page.")
st.caption("If the internet drops, you get logged out or the page changes, the run **pauses on that row** and "
           "continues by itself once things are back. It never skips a row because of that.")

if test_clicked:
    vals, _ = row_values(test_row - first_excel_row)
    try:
        with engine.Runner(used, ss.start_url, step_pause_s=step_pause, slow_typing=slow_typing) as runner:
            runner.recover()  # e.g. open the form with the 'once at start' steps
            runner.play(vals, dry_run=True)
        test_box.success("Done. Check it in the browser, then close or cancel it yourself.")
    except Exception as e:
        test_box.error(f"Could not complete the steps: {e}")

go = st.button(f"▶ Fill rows {row_from} to {row_to}", type="primary")
progress_box, table_box = st.empty(), st.empty()

if go:
    out = df.copy()
    for c in ("Status", "Message", "Processed at"):
        if c not in out.columns:
            out[c] = ""
    out[["Status", "Message", "Processed at"]] = out[["Status", "Message", "Processed at"]].fillna("").astype(str)
    idxs = list(range(row_from - first_excel_row, row_to - first_excel_row + 1))
    st.button("⏹ Stop", help="Stops after the current record.")
    bar = progress_box.progress(0.0, "Starting…")
    pause_box = st.empty()
    counts = {}
    done, started = 0, datetime.now()
    shown_cols = linked_cols + ["Status", "Message"]

    def wait_for_site(runner, problem, excel_row):
        """Pause on this row until the browser, internet and page are all back."""
        engine.beep()
        since = datetime.now()
        while problem:
            mins = int((datetime.now() - since).total_seconds() // 60)
            pause_box.warning(f"⏸ **Paused at Excel row {excel_row}** for {mins} min: {problem}.  \n"
                              "Fix it in the browser window (reconnect, log in again, go back to the page where "
                              "your recording started). The run continues by itself - nothing needs clicking here.")
            time.sleep(15)
            problem = runner.health()
        pause_box.success(f"Back - continuing from Excel row {excel_row}.")

    def process(runner, r):
        """One row, retrying through connection problems. Returns (status, message)."""
        vals, verify = row_values(r)
        check_first = skip_existing and bool(verify_cols)
        rejected = 0
        while True:
            try:
                status, msg = runner.submit_row(vals, verify, check_first)
                if status == "Not found":
                    # Saved, but the table may be unreadable because the site just dropped. Look again once back.
                    problem = runner.health()
                    if problem:
                        wait_for_site(runner, problem, r + first_excel_row)
                        row = runner.new_row(verify)
                        if row:
                            return "Registered", "Confirmed after the connection came back. Added: " + row
                return status, msg
            except engine.DataProblem as e:
                problem = runner.health()  # an error popup can also come from a dropped connection
                if not problem:
                    return "Failed", str(e)
                wait_for_site(runner, problem, r + first_excel_row)
                check_first = bool(verify_cols)
            except engine.SiteProblem as e:
                problem = runner.health()
                if problem:
                    wait_for_site(runner, problem, r + first_excel_row)
                else:
                    # Site is fine, so the website itself didn't accept this row. Give it one more go.
                    rejected += 1
                    if rejected >= 2:
                        return "Failed", str(e)
                # The last attempt may have been saved before the problem, so check before creating again.
                check_first = bool(verify_cols)

    engine.keep_awake(True)
    try:
        with engine.Runner(used, ss.start_url, timeout_s=timeout, settle_s=delay, auto_home=auto_home,
                           step_pause_s=step_pause, slow_typing=slow_typing) as runner:
            if runner.start_steps:
                bar.progress(0.0, "Doing the 'once at start' steps…")
                runner.run_start()
            for n, r in enumerate(idxs, 1):
                if resume and str(out.at[r, "Status"]) in DONE_STATUSES:
                    status, msg = "Skipped", "Already done"
                else:
                    status, msg = process(runner, r)
                    out.loc[r, ["Status", "Message", "Processed at"]] = [status, msg, f"{datetime.now():%Y-%m-%d %H:%M:%S}"]
                    done += 1
                    if done % 5 == 0:
                        ss.results = (save_results(out, up.name), out)
                counts[status] = counts.get(status, 0) + 1
                summary = " · ".join(f"{k}: {v}" for k, v in counts.items())
                eta = ""
                if done:
                    left = (datetime.now() - started).total_seconds() / done * (len(idxs) - n)
                    eta = f" — about {int(left // 3600)}h {int(left % 3600 // 60)}m left"
                bar.progress(n / len(idxs), f"Row {r + first_excel_row} ({n}/{len(idxs)}) — {summary}{eta}")
                # Only the latest rows on screen; the results file has them all.
                table_box.dataframe(out.iloc[idxs[max(0, n - 100):n]][shown_cols].iloc[::-1],
                                    width="stretch", height=300)
                if stop_on_fail and status == "Failed":
                    st.error(f"Stopped at Excel row {r + first_excel_row}: {msg}")
                    break
            else:
                # Every row is done (the run was not stopped), so finish off, e.g. Close and Save.
                if runner.end_steps:
                    names = ", ".join(engine.describe(s) for s in runner.end_steps)
                    bar.progress(1.0, "Doing the 'once at end' steps…")
                    try:
                        runner.run_end()
                        st.success(f"Final steps done: {names}.")
                    except Exception as e:
                        st.error(f"The 'once at end' steps could not be completed ({e}). "
                                 f"Please do them yourself in the browser: {names}.")
        st.success("Finished. " + " · ".join(f"{k}: {v}" for k, v in counts.items()))
    except Exception as e:
        st.error(f"Run stopped: {e}")
    finally:
        # Also runs when Stop is clicked, so no finished row is lost from the file.
        engine.keep_awake(False)
        if done:
            ss.results = (save_results(out, up.name), out)

if ss.results:
    path, out = ss.results
    st.markdown(f"**Results file:** `{path}`")
    st.caption("To retry failures: fix them in this file, upload it in step 1, and tick "
               "'Skip rows already marked done'.")
    with open(path, "rb") as fh:
        st.download_button("⬇ Download results", fh.read(), file_name=os.path.basename(path))
    st.info("If the website has its own final Save / Submit button for the whole batch, click it yourself now.")

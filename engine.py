"""Browser side of Smart Form Filler.

Talks to a Chrome window started with --remote-debugging-port. The user logs in by hand,
records one entry (Recorder), and the engine replays those steps for every Excel row (Runner).
"""
import asyncio
import math
import os
import re
import subprocess
import sys
import threading
import time
import urllib.error
import urllib.request
from collections import Counter
from contextlib import contextmanager
from datetime import date, datetime

import pandas as pd
from playwright.sync_api import sync_playwright

if sys.platform == "win32":
    # Playwright needs the Proactor loop to start its driver subprocess on Windows.
    asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())

# Several copies can run at once (one per website). Each copy has its own browser window, control port and
# login profile, so copies never see or touch each other's tabs. Copy 1 keeps the original profile.
SLOT = max(1, int(os.environ.get("SFF_SLOT", "1") or 1))
PORT = 9222 + SLOT - 1
CDP_URL = f"http://127.0.0.1:{PORT}"
HERE = os.path.dirname(os.path.abspath(__file__))
PROFILE_DIR = os.path.join(HERE, "chrome_profile" if SLOT == 1 else f"chrome_profile_{SLOT}")

DATE_FORMATS = {
    "as is": None,
    "dd.mm.yyyy": "%d.%m.%Y",
    "dd/mm/yyyy": "%d/%m/%Y",
    "dd-mm-yyyy": "%d-%m-%Y",
    "mm/dd/yyyy": "%m/%d/%Y",
    "yyyy-mm-dd": "%Y-%m-%d",
}

BROWSER_PATHS = [
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
]


# --------------------------------------------------------------------------- browser

def launch_browser(start_url="", headless=False, profile_dir=PROFILE_DIR):
    exe = next((p for p in BROWSER_PATHS if os.path.exists(p)), None)
    if not exe:
        raise RuntimeError("Chrome or Edge was not found on this PC.")
    args = [exe, f"--remote-debugging-port={PORT}", f"--user-data-dir={profile_dir}",
            "--no-first-run", "--no-default-browser-check"]
    if headless:
        args.append("--headless=new")
    if start_url:
        args.append(start_url)
    return subprocess.Popen(args)


@contextmanager
def connect():
    """Attach to the running browser. Leaving the block only disconnects; the window stays open."""
    with sync_playwright() as p:
        yield p.chromium.connect_over_cdp(CDP_URL, timeout=5000)


def keep_awake(on):
    """Stop Windows from sleeping during a long run."""
    if sys.platform == "win32":
        import ctypes
        ES_CONTINUOUS, ES_SYSTEM_REQUIRED = 0x80000000, 0x00000001
        ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS | (ES_SYSTEM_REQUIRED if on else 0))


def beep():
    if sys.platform == "win32":
        import winsound
        safe(lambda: winsound.MessageBeep(winsound.MB_ICONEXCLAMATION))


def open_url(url):
    """Open url in a new tab of the already running browser."""
    with connect() as browser:
        ctx = browser.contexts[0] if browser.contexts else browser.new_context()
        ctx.new_page().goto(url, wait_until="commit")


def is_connected():
    try:
        with connect():
            return True
    except Exception:
        return False


def is_popup(page):
    """True for a window opened by another tab (window.open). Cached on the page object."""
    if not hasattr(page, "_sff_popup"):
        def check():
            info = page.context.new_cdp_session(page).send("Target.getTargetInfo")["targetInfo"]
            return bool(info.get("openerId"))
        page._sff_popup = safe(check, False)
    return page._sff_popup


def all_pages(browser):
    """Open tabs, with popup windows (e.g. a form opened by a button) listed last."""
    pages = [pg for ctx in browser.contexts for pg in ctx.pages if not pg.is_closed()]
    return sorted(pages, key=is_popup)


def list_tabs():
    with connect() as browser:
        return [{"index": i, "title": safe(lambda: pg.title(), ""), "url": pg.url}
                for i, pg in enumerate(all_pages(browser))]


def safe(fn, default=None):
    try:
        return fn()
    except Exception:
        return default


def pump(browser, ms):
    """Wait while letting Playwright process browser events (new popups, dialogs)."""
    pages = all_pages(browser)
    if pages:
        try:
            pages[0].wait_for_timeout(ms)
            return
        except Exception:
            pass
    time.sleep(ms / 1000)


# --------------------------------------------------------------------------- recording

# Shared page helpers: a stable CSS path for an element, and a human label for it.
HELPERS_JS = r"""
  const clean = t => (t || '').replace(/[*:]/g, ' ').replace(/\s+/g, ' ').trim();
  const unique = sel => { try { return document.querySelectorAll(sel).length === 1; } catch (e) { return false; } };
  const path = el => {
    if (el.id && unique('#' + CSS.escape(el.id))) return '#' + CSS.escape(el.id);
    const tag = el.tagName.toLowerCase();
    const nm = el.getAttribute('name');
    if (nm) {
      const s = tag + '[name="' + nm.replace(/"/g, '\\"') + '"]';
      if (unique(s)) return s;
    }
    const parts = [];
    let e = el;
    while (e && e.nodeType === 1 && e !== document.body) {
      if (e.id && unique('#' + CSS.escape(e.id))) { parts.unshift('#' + CSS.escape(e.id)); return parts.join(' > '); }
      let p = e.tagName.toLowerCase();
      const sibs = [...e.parentNode.children].filter(c => c.tagName === e.tagName);
      if (sibs.length > 1) p += ':nth-of-type(' + (sibs.indexOf(e) + 1) + ')';
      parts.unshift(p);
      e = e.parentElement;
    }
    parts.unshift('body');
    return parts.join(' > ');
  };
  const labelOf = el => {
    if (el.labels && el.labels.length) { const t = clean(el.labels[0].innerText); if (t) return t; }
    const aria = el.getAttribute('aria-label'); if (aria) return clean(aria);
    if (el.placeholder) return clean(el.placeholder);
    const cell = el.closest('td,th');
    if (cell) {
      for (let p = cell.previousElementSibling; p; p = p.previousElementSibling) {
        const t = clean(p.innerText); if (t) return t.slice(0, 60);
      }
    }
    for (let p = el.previousElementSibling; p; p = p.previousElementSibling) {
      const t = clean(p.innerText); if (t) return t.slice(0, 60);
    }
    if (el.parentElement) { const t = clean(el.parentElement.innerText); if (t && t.length < 60) return t; }
    return el.getAttribute('name') || el.id || '';
  };
  const textOf = el => clean(el.innerText || el.value || el.getAttribute('aria-label') ||
                             el.getAttribute('alt') || el.getAttribute('title') || '').slice(0, 80);
"""

RECORDER_JS = r"""
(binding) => {
  if (document.__sffRecorder === binding) return;
  document.__sffRecorder = binding;
""" + HELPERS_JS + r"""
  const send = d => {
    try { window[binding](Object.assign(d, { url: location.href, popup: !!window.opener })); } catch (e) {}
  };
  const TYPED = /^(text|email|tel|number|search|url|date|datetime-local|month|week|time|password|)$/;
  const typed = el => el.tagName === 'TEXTAREA' || (el.tagName === 'INPUT' && TYPED.test((el.getAttribute('type') || '').toLowerCase()));
  const fill = el => send({ action: 'fill', selector: path(el), label: labelOf(el),
                            value: el.type === 'password' ? '' : el.value, secret: el.type === 'password',
                            kind: el.type === 'date' ? 'date' : 'text' });

  document.addEventListener('click', e => {
    if (!e.isTrusted || !(e.target instanceof Element)) return;
    const t = e.target;
    // Typing, dropdowns and tick boxes are recorded by their value, not by the click on them.
    if (t.closest('select, option, textarea') || (t.closest('input') && !/^(submit|button|image|reset)$/.test(t.closest('input').type))) return;
    const lab = t.closest('label');
    if (lab && lab.control) return;
    const c = t.closest('button, a, input, [role=button], [role=link], [role=menuitem], [role=tab], [role=option], [onclick], summary') || t;
    send({ action: 'click', selector: path(c), text: textOf(c) });
  }, true);

  document.addEventListener('input', e => { if (e.isTrusted && typed(e.target)) fill(e.target); }, true);
  document.addEventListener('change', e => {
    if (!e.isTrusted) return;
    const el = e.target;
    if (el.tagName === 'SELECT') {
      const o = el.options[el.selectedIndex];
      send({ action: 'select', selector: path(el), label: labelOf(el), value: o ? o.text.trim() : '',
             options: [...el.options].map(o => o.text.trim()).filter(Boolean).slice(0, 400) });
    } else if (el.type === 'checkbox' || el.type === 'radio') {
      send({ action: 'check', selector: path(el), label: labelOf(el) || textOf(el), value: el.checked ? 'yes' : 'no', kind: el.type });
    } else if (typed(el)) {
      fill(el);
    }
  }, true);
  document.addEventListener('keydown', e => {
    if (e.isTrusted && e.key === 'Enter' && typed(e.target) && e.target.tagName === 'INPUT') {
      fill(e.target);
      send({ action: 'press_enter', selector: path(e.target), label: labelOf(e.target) });
    }
  }, true);
}
"""


class Recorder:
    """Watches the tool's browser in a background thread and turns the user's actions into steps."""

    def __init__(self):
        self.steps = []
        self.start_url = ""
        self.error = None
        self.binding = f"__sffRec{int(time.time() * 1000)}"
        self._lock = threading.Lock()
        self._stop = threading.Event()
        self._ready = threading.Event()
        self._last_click = 0.0
        self._thread = None

    def start(self):
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()
        self._ready.wait(20)
        if self.error:
            raise RuntimeError(self.error)
        return self

    def stop(self):
        self._stop.set()
        if self._thread:
            self._thread.join(20)
        return self.snapshot()

    def snapshot(self):
        with self._lock:
            return [dict(s) for s in self.steps]

    @property
    def running(self):
        return bool(self._thread and self._thread.is_alive())

    # ---------------------------------------------------------------- events

    def _add(self, step):
        with self._lock:
            step["id"] = len(self.steps)
            self.steps.append(step)

    def _on_event(self, source, d):
        d = {k: v for k, v in d.items() if v not in (None, "")} | {"value": d.get("value", "")}
        with self._lock:
            if d["action"] in ("select", "check"):
                # These can reload the page too (e.g. choosing a cohort), so that reload is not a separate step.
                self._last_click = time.time()
            if d["action"] in ("fill", "select", "check"):
                # Typing fires many events: keep one step per field, updated, until something is clicked.
                for s in reversed(self.steps):
                    if s["action"] in ("click", "press_enter", "goto"):
                        break
                    if s["action"] == d["action"] and s["selector"] == d["selector"] and s.get("url") == d.get("url"):
                        s["value"] = d["value"]
                        return
            if d["action"] == "click":
                last = self.steps[-1] if self.steps else None
                if last and last["action"] == "click" and last["selector"] == d["selector"] \
                        and time.time() - self._last_click < 0.5:
                    return  # double-click
                self._last_click = time.time()
            if d["action"] == "press_enter":
                self._last_click = time.time()
        self._add(d)

    def _on_navigate(self, page, url):
        # A page change right after a click is caused by that click; replaying the click repeats it.
        if url.startswith("http") and time.time() - self._last_click > 4:
            if not (self.steps and self.steps[-1].get("action") == "goto" and self.steps[-1]["url"] == url):
                self._add({"action": "goto", "url": url, "value": ""})

    def _on_close(self, page):
        # Remember that the last click in a popup closed it, so a replay can tell success from an error.
        with self._lock:
            for s in reversed(self.steps):
                if s["action"] == "click":
                    if s.get("popup"):
                        s["closes_popup"] = True
                    break

    # ---------------------------------------------------------------- thread

    def _run(self):
        try:
            with sync_playwright() as p:
                browser = p.chromium.connect_over_cdp(CDP_URL, timeout=5000)
                ctx = browser.contexts[0]
                ctx.expose_binding(self.binding, self._on_event)

                def install(pg):
                    for fr in pg.frames:
                        safe(lambda: fr.evaluate(RECORDER_JS, self.binding))

                def watch(pg):
                    pg.on("domcontentloaded", lambda: install(pg))
                    pg.on("dialog", lambda dlg: (self._add({"action": "dialog", "text": dlg.message, "value": ""}),
                                                 safe(dlg.accept)))
                    pg.on("framenavigated", lambda fr: fr == pg.main_frame and self._on_navigate(pg, fr.url))
                    pg.on("close", lambda: self._on_close(pg))

                ctx.on("page", watch)
                for pg in ctx.pages:
                    watch(pg)
                    install(pg)
                focused = [pg for pg in ctx.pages if safe(lambda: pg.evaluate("document.hasFocus()"), False)]
                self.start_url = (focused or ctx.pages or [None])[0].url if ctx.pages else ""
                self._ready.set()
                while not self._stop.is_set():
                    # Popups and fresh pages do not always run init scripts, so (re)install regularly.
                    for pg in list(ctx.pages):
                        install(pg)
                    pages = ctx.pages
                    if pages:
                        safe(lambda: pages[0].wait_for_timeout(400))
                    else:
                        time.sleep(0.4)
        except Exception as e:
            self.error = (str(e).splitlines() or ["Could not record"])[0]
            self._ready.set()


def describe(step):
    """Plain-language description of a recorded step."""
    a = step["action"]
    where = step.get("label") or step.get("text") or step.get("selector", "")
    if a == "click":
        return f"Click '{step.get('text') or step.get('selector')}'"
    if a == "fill":
        return f"Type into '{where}'"
    if a == "select":
        return f"Choose in '{where}'"
    if a == "check":
        return f"{'Tick' if step.get('kind') != 'radio' else 'Pick option'} '{where}'"
    if a == "press_enter":
        return f"Press Enter in '{where}'"
    if a == "goto":
        return f"Go to {step['url'][:80]}"
    if a == "dialog":
        return f"Popup message: '{step.get('text', '')[:60]}' (accepted automatically)"
    return a


# --------------------------------------------------------------------------- locating

def _frames(browser, prefer_page=None, popup_first=False):
    pages = all_pages(browser)
    if popup_first:
        pages.sort(key=lambda pg: not is_popup(pg))
    if prefer_page in pages:
        pages.remove(prefer_page)
        pages.insert(0, prefer_page)
    for pg in pages:
        for fr in pg.frames:
            yield pg, fr


def find_visible(browser, selector, text=None, prefer_page=None, popup_first=False):
    """First visible element matching selector (or, failing that, the button text) on any tab/frame."""
    for pg, fr in _frames(browser, prefer_page, popup_first):
        loc = fr.locator(selector).first
        if safe(lambda: loc.is_visible(), False):
            return pg, loc
    if text:
        for pg, fr in _frames(browser, prefer_page, popup_first):
            for loc in (fr.get_by_role("button", name=text, exact=True).first,
                        fr.get_by_role("link", name=text, exact=True).first,
                        fr.locator(f'input[value="{text}"]').first,
                        fr.get_by_text(text, exact=True).first):
                if safe(lambda: loc.is_visible(), False):
                    return pg, loc
    return None, None


def wait_visible(browser, selector, timeout_s, text=None, prefer_page=None, popup_first=False):
    end = time.time() + timeout_s
    while time.time() < end:
        pg, loc = find_visible(browser, selector, text, prefer_page, popup_first)
        if loc:
            return pg, loc
        pump(browser, 250)
    return None, None


ROW_JS = r"""
(vals) => {
  const n = s => (s || '').toLowerCase().replace(/\s+/g, ' ').trim();
  const want = vals.map(n).filter(Boolean);
  if (!want.length) return [];
  const found = [];
  for (const tr of document.querySelectorAll('tr, [role=row]')) {
    if (tr.querySelector('tr, [role=row]')) continue;
    let cells = [...tr.querySelectorAll('td, th, [role=cell], [role=gridcell]')];
    if (!cells.length) cells = [...tr.children];
    const texts = new Set(cells.map(c => {
      const box = c.querySelector('input:not([type=checkbox]):not([type=hidden]), select');
      return n(c.innerText) || (box ? n(box.value) : '');
    }));
    // Whole-cell match, so "Ali" does not match "Alia" or "Al Ali".
    if (want.every(v => texts.has(v))) found.push(tr.innerText.replace(/\s+/g, ' ').trim());
  }
  return found;
}
"""


ERRORS_JS = r"""
() => {
  const sel = '[role=alert], .error, .errors, .alert-danger, .invalid-feedback, .validation-summary-errors, ' +
              '.field-validation-error, [class*="error" i], [id*="error" i], [class*="invalid" i]';
  const seen = new Set(), out = [];
  for (const el of document.querySelectorAll(sel)) {
    const r = el.getBoundingClientRect(), t = (el.innerText || '').replace(/\s+/g, ' ').trim();
    if (r.width > 0 && r.height > 0 && t && t.length < 300 && !seen.has(t)) { seen.add(t); out.push(t); }
  }
  return out.slice(0, 3).join(' | ');
}
"""


def page_errors(browser):
    """Error messages a website shows on the page itself (not in a popup), if any."""
    for _, fr in _frames(browser, popup_first=True):
        text = safe(lambda: fr.evaluate(ERRORS_JS), "")
        if text:
            return text
    return ""


def matching_rows(browser, values, skip_page=None):
    """Text of every table row (on any tab) that has each value as a whole cell."""
    rows = []
    for pg, fr in _frames(browser):
        if pg is not skip_page:
            rows += safe(lambda: fr.evaluate(ROW_JS, values), None) or []
    return rows


def row_exists(browser, values, skip_page=None):
    return bool(matching_rows(browser, values, skip_page))


# --------------------------------------------------------------------------- values

def to_text(value, date_format=None):
    if value is None:
        return ""
    if isinstance(value, float) and math.isnan(value):
        return ""
    if safe(lambda: bool(pd.isna(value)), False):
        return ""
    fmt = DATE_FORMATS.get(date_format or "as is")
    if fmt:
        dt = value if isinstance(value, (datetime, date)) else safe(
            lambda: pd.to_datetime(str(value).strip(), dayfirst=True))
        if dt is not None and not pd.isna(dt):
            return dt.strftime(fmt)
    if isinstance(value, datetime):
        return value.strftime("%d.%m.%Y") if value.time() == datetime.min.time() else str(value)
    if isinstance(value, date):
        return value.strftime("%d.%m.%Y")
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def truthy(text):
    return str(text).strip().lower() in ("1", "yes", "y", "true", "x", "on", "checked", "ticked")


def pick_option(options, wanted):
    """Best matching <option> for an Excel value: exact text, exact value, then partial."""
    w = wanted.lower().strip()
    for key in ("t", "v"):
        for o in options:
            if o[key].lower().strip() == w:
                return o
    # Abbreviations: "UAE" -> "Utd.Arab Emir."
    for o in options:
        if len(w) >= 2 and "".join(p[0] for p in re.split(r"[^a-z]+", o["t"].lower()) if p) == w:
            return o
    for o in options:
        if o["t"] and (o["t"].lower().startswith(w) or w.startswith(o["t"].lower())):
            return o
    for o in options:
        if o["t"] and w in o["t"].lower():
            return o
    return None


def fill_one(loc, ftype, text, slow=False):
    if ftype == "select":
        options = loc.evaluate("s => [...s.options].map(o => ({v: o.value, t: o.text.trim()}))")
        opt = pick_option(options, text)
        if not opt:
            shown = ", ".join(o["t"] for o in options if o["t"])[:200]
            raise ValueError(f"'{text}' is not in the list ({shown})")
        loc.select_option(value=opt["v"])
    elif ftype in ("checkbox", "radio"):
        if truthy(text):
            loc.check()
        elif ftype == "checkbox":
            loc.uncheck()
    else:
        if ftype == "date":
            text = to_text(text, "yyyy-mm-dd")
        if slow:
            loc.fill("")
            loc.press_sequentially(text, delay=80)  # letter by letter, like a person
        else:
            loc.fill(text)
        # Pages that copy typed text elsewhere (e.g. "name printed on certificate") listen for keys.
        loc.dispatch_event("keyup")
        loc.dispatch_event("change")
        loc.dispatch_event("blur")


# --------------------------------------------------------------------------- running

class DataProblem(Exception):
    """Something wrong with this row's data. Mark the row and move on."""


class SiteProblem(Exception):
    """The page, connection or browser was not as expected. Check health before blaming the row."""


def _watch_dialogs(browser, messages):
    def hook(pg):
        pg.on("dialog", lambda d: (messages.append(d.message), safe(d.accept)))
    for ctx in browser.contexts:
        ctx.on("page", hook)
        for pg in ctx.pages:
            hook(pg)


def _open_popup(page):
    return not page.is_closed() and is_popup(page)


def site_reachable(url):
    """True when the website answers at all (any status below 500)."""
    if not url.startswith("http"):
        return True
    origin = "/".join(url.split("/")[:3]) + "/"
    try:
        urllib.request.urlopen(urllib.request.Request(origin, method="HEAD"), timeout=8)
        return True
    except urllib.error.HTTPError as e:
        return e.code < 500
    except Exception:
        return False


ELEMENT_ACTIONS = ("click", "fill", "select", "check", "press_enter")


class Runner:
    """Replays recorded steps, one Excel row at a time, over one browser connection."""

    def __init__(self, steps, start_url="", timeout_s=20, settle_s=1.5, auto_home=False,
                 step_pause_s=1.0, slow_typing=False):
        used = [s for s in steps if s.get("use", True) and s["action"] != "dialog"]
        # "when": done once before the first row, for every row (the default), or once after the last row.
        self.start_steps = [s for s in used if s.get("when") == "start"]
        self.end_steps = [s for s in used if s.get("when") == "end"]
        self.steps = [s for s in used if s.get("when", "row") == "row"]
        self.start_url = start_url
        self.auto_home = auto_home
        self.step_pause_s = step_pause_s
        self.slow_typing = slow_typing
        self.timeout_s = timeout_s
        self.settle_s = settle_s
        self.dialogs = []
        self.site_url = start_url
        self._before = Counter()
        self._pw = None
        self.browser = None

    def __enter__(self):
        self._pw = sync_playwright().start()
        self._connect()
        return self

    def __exit__(self, *exc):
        safe(self._pw.stop)

    def _connect(self):
        self.browser = self._pw.chromium.connect_over_cdp(CDP_URL, timeout=5000)
        _watch_dialogs(self.browser, self.dialogs)
        if not self.site_url:
            self.site_url = next((pg.url for pg in all_pages(self.browser) if pg.url.startswith("http")), "")

    # ----------------------------------------------------------- health

    def ready(self):
        """True when the page is where the recording started (its first step can be done)."""
        first = next((s for s in self.steps if s["action"] in ELEMENT_ACTIONS + ("goto",)), None)
        if not first or first["action"] == "goto":
            return True
        try:
            _, loc = find_visible(self.browser, first["selector"], first.get("text"))
            return loc is not None
        except Exception:
            return False

    def health(self):
        """None when everything looks fine, otherwise what is wrong (in plain words)."""
        if not safe(lambda: self.browser.is_connected(), False):
            try:
                self._connect()
            except Exception:
                return "The tool's browser window was closed"
        if not site_reachable(self.site_url):
            return "No internet connection, or the website is not responding"
        if not self.recover():
            return ("The page is not at the starting point of your recording - you may have been logged out, "
                    "or the page was changed")
        return None

    def recover(self):
        """Make sure each row's first step can be done: redo the start steps, or reopen the start page first."""
        if self.ready():
            return True
        if self.start_steps and self._try(self.start_steps) and self.ready():
            return True
        if self.auto_home:
            self.go_home()
            if self.ready():
                return True
            if self.start_steps and self._try(self.start_steps) and self.ready():
                return True
        return False

    def _try(self, steps):
        try:
            self.play({}, steps=steps)
            return True
        except Exception:
            return False

    def run_start(self):
        """Steps marked 'once at start' (e.g. opening the form the first time)."""
        if self.start_steps:
            self.play({}, steps=self.start_steps)

    def run_end(self):
        """Steps marked 'once at end' (e.g. Close, then Save)."""
        if self.end_steps:
            self.dialogs.clear()
            self.recover()  # if the last row failed, its form was closed; the end steps expect it open
            self.play({}, steps=self.end_steps)

    def go_home(self):
        """Back to the page where the recording started (after a record left the page somewhere else)."""
        self.reset()
        pg = next((p for p in all_pages(self.browser) if not is_popup(p)), None)
        if pg and self.start_url.startswith("http"):
            safe(lambda: pg.goto(self.start_url, wait_until="domcontentloaded", timeout=self.timeout_s * 1000))
            pump(self.browser, 500)

    def reset(self):
        """Close half-filled popup windows so the next attempt starts clean."""
        for pg in safe(lambda: all_pages(self.browser), []):
            if _open_popup(pg):
                safe(pg.close)

    # ----------------------------------------------------------- playing

    def save_step(self):
        """Index of the click that saves: the first click (or Enter) after the last typed or chosen value.

        Not simply the last click - after saving, a site may still need e.g. "Close" on a confirmation page.
        """
        last_value = max((i for i, s in enumerate(self.steps) if s["action"] in ("fill", "select", "check")),
                         default=-1)
        return next((i for i, s in enumerate(self.steps)
                     if i > last_value and s["action"] in ("click", "press_enter")),
                    max((i for i, s in enumerate(self.steps) if s["action"] == "click"), default=-1))

    def play(self, values, dry_run=False, steps=None):
        """Do the row's steps (or the given steps). values: {step id: text} for steps fed from Excel.

        dry_run stops before the save click, so the record is filled in but not saved.
        """
        steps = self.steps if steps is None else steps
        stop = self.save_step() if dry_run else -1
        for i, s in enumerate(steps):
            if i == stop:
                return
            a, n = s["action"], f"Step {i + 1} ({describe(s)})"
            if a == "goto":
                pg = next((p for p in all_pages(self.browser) if not is_popup(p)), None)
                if not pg:
                    raise SiteProblem(f"{n}: no browser tab open")
                try:
                    pg.goto(s["url"], wait_until="domcontentloaded", timeout=self.timeout_s * 1000)
                except Exception as e:
                    raise SiteProblem(f"{n}: {(str(e).splitlines() or [''])[0]}") from None
                self._pause_after(s)
                continue
            popup_first = bool(s.get("popup"))
            pg, loc = wait_visible(self.browser, s["selector"], self.timeout_s * 0.6, popup_first=popup_first)
            if not loc:
                # Missing because the website rejected what was typed? Then report that, rather than
                # clicking a look-alike button (e.g. the form's own "Close") by its text.
                shown = page_errors(self.browser)
                if self.dialogs or shown:
                    raise DataProblem("; ".join(self.dialogs) or "The website shows: " + shown)
                if a == "click" and s.get("text"):
                    pg, loc = wait_visible(self.browser, s["selector"], self.timeout_s * 0.4, s["text"],
                                           popup_first=popup_first)
            if not loc:
                raise SiteProblem(f"{n}: not found on the page")
            text = values.get(s["id"], s.get("value", ""))
            try:
                if a == "click":
                    loc.click(timeout=self.timeout_s * 1000)
                elif a == "fill":
                    fill_one(loc, s.get("kind", "text"), text, self.slow_typing)
                elif a == "select":
                    if text != "":
                        fill_one(loc, "select", text)
                elif a == "check":
                    fill_one(loc, s.get("kind", "checkbox"), text)
                elif a == "press_enter":
                    loc.press("Enter")
            except ValueError as e:
                raise DataProblem(f"{s.get('label') or describe(s)}: {e}") from None
            except Exception as e:
                raise SiteProblem(f"{n}: {(str(e).splitlines() or ['failed'])[0]}") from None
            if a == "click" and s.get("closes_popup") and not dry_run:
                # When recorded, this click closed the popup. If it stays open, the website refused the record.
                end = time.time() + self.timeout_s
                while not pg.is_closed():
                    if time.time() > end:
                        if self.dialogs:
                            raise DataProblem("; ".join(self.dialogs))
                        raise SiteProblem(f"{n}: the window stayed open (the website may show an error on it)")
                    pump(self.browser, 250)
                if s.get("wait") is not None:
                    self._pause_after(s)
            else:
                self._pause_after(s)

    def _pause_after(self, step):
        """Wait after a step: its own 'wait after' if set, otherwise the general pause between steps.

        Pausing lets the website keep up (and it looks like a person at work).
        """
        own = step.get("wait")
        seconds = own if own is not None else self.step_pause_s
        floor = 0.3 if own is None and step["action"] == "click" else 0.05
        pump(self.browser, int(max(seconds, floor) * 1000))

    def submit_row(self, values, verify_values, skip_existing=True):
        """Do all steps for one record and return (status, message).

        Raises DataProblem for a bad value in this row, and SiteProblem when the page did not
        behave; the caller then checks health() to tell a lost connection from a row the website rejected.
        """
        self.dialogs.clear()
        try:
            if not self.recover():
                raise SiteProblem("The page is not at the starting point of your recording")
            before = matching_rows(self.browser, verify_values) if verify_values else []
            if skip_existing and before:
                return "Already exists", "Already in the table: " + before[0]
            # Remember what matched before, so success means a *new* matching row appeared.
            self._before = Counter(before)
            self.play(values)
        except (DataProblem, SiteProblem):
            self.reset()
            raise
        except Exception as e:  # browser/page went away mid-row
            self.reset()
            raise SiteProblem((str(e).splitlines() or ["browser error"])[0]) from None

        pump(self.browser, int(self.settle_s * 1000))
        note = "; ".join(self.dialogs)
        if not verify_values:
            return "Submitted", note
        end = time.time() + self.timeout_s / 2
        while time.time() < end:
            row = self.new_row(verify_values)
            if row:
                return "Registered", (note + "; " if note else "") + "Added: " + row
            pump(self.browser, 500)
        shown = page_errors(self.browser)
        if shown:
            raise DataProblem("The website shows: " + shown)
        return "Not found", (note + "; " if note else "") + "Done, but no new matching row appeared in the table"

    def new_row(self, verify_values):
        """Text of a matching row that was not in the table before this record was submitted, or None."""
        open_popups = [pg for pg in all_pages(self.browser) if is_popup(pg)]
        added = Counter(matching_rows(self.browser, verify_values,
                                      skip_page=open_popups[0] if open_popups else None)) - self._before
        return next(iter(added), None)

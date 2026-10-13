"""KGB document upload automation for EHRM.

Attaches to a Chrome instance the operator already started with
--remote-debugging-port=9222 and logged into, then walks the employee rows of
the Excel workbook and fills one KGB form per employee.

The script is deliberately semi-automatic: it pauses with input() at each
operator checkpoint so the EHRM page can be corrected between steps.
"""

import functools
import glob
import os
import time
from datetime import date, datetime

import pandas as pd
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

# --- Connection / environment ------------------------------------------------

DEBUGGER_ADDRESS = "127.0.0.1:9222"
EHRM_BASE_URL = "https://ehrm.pu.go.id"
KGB_FOLDER = "kgb"
SHEET_NAME = "PNS"

WAIT_SECONDS = 10
SETTLE_SECONDS = 3
MAX_SCROLL_ATTEMPTS = 5
ADD_DATA_ATTEMPTS = 3

# Both date fields are typed straight into the input, so the EHRM datepicker
# never has to be navigated. The fields re-parse whatever is typed, so a
# format they do not understand comes back as "NaN"; DATE_STYLES lists the
# shapes to try, most likely first, and DATE_INPUT_STYLE in .env can pin one.
MONTH_NAMES_EN = [
    "",
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
]

MONTH_NAMES_ID = [
    "",
    "Januari",
    "Februari",
    "Maret",
    "April",
    "Mei",
    "Juni",
    "Juli",
    "Agustus",
    "September",
    "Oktober",
    "November",
    "Desember",
]

MONTH_SHORT_ID = [
    "",
    "Jan",
    "Feb",
    "Mar",
    "Apr",
    "Mei",
    "Jun",
    "Jul",
    "Agu",
    "Sep",
    "Okt",
    "Nov",
    "Des",
]

DATE_STYLES = {
    "d mmmm id": lambda d: f"{d.day} {MONTH_NAMES_ID[d.month]} {d.year}",
    "d mmm id": lambda d: f"{d.day} {MONTH_SHORT_ID[d.month]} {d.year}",
    "d mmmm yyyy": lambda d: f"{d.day} {MONTH_NAMES_EN[d.month]} {d.year}",
    "dd-mm-yyyy": lambda d: d.strftime("%d-%m-%Y"),
    "dd/mm/yyyy": lambda d: d.strftime("%d/%m/%Y"),
    "mmmm d, yyyy": lambda d: f"{MONTH_NAMES_EN[d.month]} {d.day}, {d.year}",
    "mmmm d yyyy id": lambda d: f"{MONTH_NAMES_ID[d.month]} {d.day} {d.year}",
    "yyyy-mm-dd": lambda d: d.strftime("%Y-%m-%d"),
    "d mmm yyyy": lambda d: f"{d.day} {MONTH_NAMES_EN[d.month][:3]} {d.year}",
}

DEFAULT_DATE_STYLES = (
    "d mmmm id",
    "d mmm id",
    "d mmmm yyyy",
    "dd-mm-yyyy",
    "dd/mm/yyyy",
    "mmmm d yyyy id",
    "mmmm d, yyyy",
    "yyyy-mm-dd",
)

# Month names the workbook uses ("26 Juni 2026") and the usual abbreviations.
MONTH_NUMBERS = {
    "januari": 1,
    "jan": 1,
    "january": 1,
    "februari": 2,
    "feb": 2,
    "february": 2,
    "maret": 3,
    "mar": 3,
    "march": 3,
    "april": 4,
    "apr": 4,
    "mei": 5,
    "may": 5,
    "juni": 6,
    "jun": 6,
    "june": 6,
    "juli": 7,
    "jul": 7,
    "july": 7,
    "agustus": 8,
    "agu": 8,
    "aug": 8,
    "august": 8,
    "september": 9,
    "sep": 9,
    "sept": 9,
    "oktober": 10,
    "okt": 10,
    "oct": 10,
    "october": 10,
    "november": 11,
    "nov": 11,
    "desember": 12,
    "des": 12,
    "dec": 12,
    "december": 12,
}

# --- Locators ----------------------------------------------------------------
#
# These absolute XPaths are verified against the live EHRM pages. They are
# brittle by nature; when a step starts failing, check this block first and
# re-derive the path in DevTools rather than editing the step functions.

# Every field of the "Tambah Data" KGB form shares this prefix.
FORM = (
    '//*[@id="app"]/div/main/div/section/section/div[2]/div/div[2]/div[2]/div/div[1]/'
)

X_SK_NUMBER = FORM + "div[2]/div[2]/div/div/input"
X_SK_DATE = FORM + "div[3]/div[2]/div/div/div/input"
X_SK_PREV_MONTH = (
    FORM + "div[3]/div[2]/div/div[2]/div/div[2]/div/div[1]/div/div[1]/button[1]/span"
)
X_YEAR_OLD = FORM + "div[5]/div[2]/div/div/input"
X_MONTH_OLD = FORM + "div[6]/div[2]/div/div/input"
X_YEAR_NEW = FORM + "div[7]/div[2]/div/div/input"
X_MONTH_NEW = FORM + "div[8]/div[2]/div/div/input"
X_EFFECTIVE_DATE = FORM + "div[10]/div[2]/div/div[1]/div/input"
X_EFFECTIVE_PREV_MONTH = (
    FORM + "div[10]/div[2]/div/div[2]/div/div[2]/div/div[1]/div/div[1]/button[1]"
)
X_EFFECTIVE_NEXT_MONTH = (
    FORM + "div[10]/div[2]/div/div[2]/div/div[2]/div/div[1]/div/div[1]/button[2]"
)
X_PEJABAT = FORM + "div[11]/div[2]/div/div/input"

X_CONFIRM = "//button[normalize-space(text())='Tambah Data']"
X_CONFIRM_SPAN = "//button[.//span[normalize-space()='Tambah Data']]"
X_CONFIRM_LOOSE = "//button[contains(normalize-space(.), 'Tambah Data')]"
CONFIRM_SELECTORS = (X_CONFIRM, X_CONFIRM_SPAN, X_CONFIRM_LOOSE)

# Employee lookup page (search instead of direct URL navigation).
X_NIP_FIELD = (
    '//*[@id="app"]/div/main/div/div/div/div/div/div[1]/form/div[1]/div[1]/div/input'
)
X_SEARCH_BUTTON = (
    '//*[@id="app"]/div/main/div/div/div/div/div/div[1]/form/div[3]/div[1]/button'
)
X_SEARCH_BUTTON_BY_TEXT = "//button[contains(., 'Cari Data')]"
X_VIEW_EMPLOYEE = '//*[@id="datatable-basic"]/tbody/tr/td[10]/div/div[1]/a/button'
X_SALARY_TAB = (
    '//*[@id="app"]/div/main/div/section/section/div[3]/div/div/div/button[4]'
)
X_ADD_DATA = "//button[normalize-space(text())='Tambah Data']"
# The label is usually wrapped in a <span>, and the toolbar copy varies
# ("Tambah Data" vs "Tambah Data KGB"), so click_add_data_button tries these
# looser locators in order instead of one brittle exact match.
X_ADD_DATA_SPAN = "//button[.//span[normalize-space()='Tambah Data']]"
X_ADD_DATA_LOOSE = "//button[contains(normalize-space(.), 'Tambah')]"
# Specific selector for the full-width "Tambah Data" button on the salary tab
X_ADD_DATA_FULL_WIDTH = (
    "//button[contains(@style, 'width: 100%') and contains(., 'Tambah Data')]"
)
ADD_DATA_SELECTORS = (
    X_ADD_DATA_FULL_WIDTH,
    X_ADD_DATA,
    X_ADD_DATA_SPAN,
    X_ADD_DATA_LOOSE,
)

# Element the operator uses to confirm the row was actually written.
X_SUBMITTED_ROW = '//*[@id="datatable-basic"]/tbody/tr[1]/td[8]'


def employee_salary_url(employee_id):
    """Salary history page for one employee."""
    return (
        f"{EHRM_BASE_URL}/layanan-kepegawaian/peremajaan-data"
        f"/riwayat-gaji/{employee_id}?is-verif=true"
    )


# --- Small helpers -----------------------------------------------------------


def log_info(msg):
    print(f"[INFO] {msg}")


def log_warn(msg):
    print(f"[WARN] {msg}")


def log_error(msg):
    print(f"[ERROR] {msg}")


def log_step(msg):
    print(f"[STEP] {msg}")


def log_debug(msg):
    print(f"[DEBUG] {msg}")


def pause(prompt):
    """Wait for the operator. Used at every human checkpoint."""
    log_info(prompt)
    input()


def locate(driver, selector, wait=None, by=By.XPATH):
    """Wait for a single element and return it."""
    wait = wait or WebDriverWait(driver, WAIT_SECONDS)
    return wait.until(EC.presence_of_element_located((by, selector)))


def send_keys(driver, xpath, value, label, wait=None):
    """Clear an input and type into it."""
    field = locate(driver, xpath, wait)
    log_step(f"{label}: field located")
    field.clear()
    field.send_keys(str(value))
    log_info(f"{label}: '{value}' filled")
    return True


def parse_date(value):
    """Read a workbook or .env value as a real date.

    Accepts timestamps, ISO strings (2026-07-01), Indonesian long dates
    (26 Juni 2026) and anything DATE_STYLES can produce. Returns None when the
    value is blank or unreadable, so the caller can report it instead of
    sending garbage to the form.
    """
    if is_blank(value):
        return None

    if isinstance(value, pd.Timestamp):
        return value.date()
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value

    text = " ".join(str(value).split())
    if not text:
        return None

    for fmt in (
        "%d-%m-%Y",
        "%Y-%m-%d",
        "%d/%m/%Y",
        "%Y/%m/%d",
        "%d-%m-%y",
        "%d-%B-%Y",
        "%d %B %Y",
        "%B %d %Y",
        "%d-%b-%Y",
        "%d %b %Y",
    ):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue

    parts = text.split()
    if len(parts) == 3:
        day, month_name, year = parts
        month = MONTH_NUMBERS.get(month_name.lower().rstrip("."))
        if month and day.isdigit() and year.isdigit():
            return date(int(year), month, int(day))

    return None


def date_styles():
    """Date formats to try, most likely first.

    DATE_INPUT_STYLE in .env pins the order, e.g. "d mmmm yyyy, dd-mm-yyyy".
    """
    raw = os.getenv("DATE_INPUT_STYLE") or ""
    styles = []
    for item in raw.split(","):
        key = " ".join(item.lower().split())
        if key in DATE_STYLES and key not in styles:
            styles.append(key)
    styles += [s for s in DEFAULT_DATE_STYLES if s not in styles]
    return styles


def to_date_literal(value, style=None):
    """Render a value in one DATE_STYLES format, for logging.

    Anything unreadable comes back untouched so a bad cell stays visible on
    the EHRM field instead of being silently reformatted.
    """
    day = parse_date(value)
    if day is None:
        if is_blank(value):
            return None
        log_warn(f"could not read '{value}' as a date, using it verbatim")
        return " ".join(str(value).split())

    style = style or date_styles()[0]
    return DATE_STYLES[style](day)


# The EHRM date fields are third-party datepickers: they own their value and
# re-parse whatever is typed, so text they do not recognise comes back as "NaN"
# (the month/year header of the calendar even shows NaN). Writing the date into
# the component that owns the value is therefore the only reliable way, and the
# owner is not always obvious: a Vue component is found through its model, and a
# web component through the properties and setters on its host element.
JS_SET_DATE_VIA_MODEL = """
var input = arguments[0], year = arguments[1], month = arguments[2], day = arguments[3];

function typeName(inst) {
  var t = (inst && inst.type) || {};
  return String(t.name || t.__name || '');
}
function rootEl(inst) {
  if (inst.subTree && inst.subTree.el) return inst.subTree.el;
  return (inst.vnode && inst.vnode.el) || null;
}
function ownsValue(inst) {
  if (!inst) return false;
  if (/date|calendar/i.test(typeName(inst))) return true;
  var p = inst.props || {};
  if ('modelValue' in p || 'dateFormat' in p) return true;
  var e = inst.emitsOptions;
  if (Array.isArray(e)) return e.indexOf('update:modelValue') !== -1;
  return !!(e && e['update:modelValue']);
}
function walk(vnode, out, depth) {
  if (!vnode || depth > 60) return;
  if (vnode.component) {
    out.push(vnode.component);
    walk(vnode.component.subTree, out, depth + 1);
  }
  var kids = vnode.children;
  if (Array.isArray(kids)) {
    for (var i = 0; i < kids.length; i++) walk(kids[i], out, depth + 1);
  }
  var dyn = vnode.dynamicChildren;
  if (Array.isArray(dyn)) {
    for (var j = 0; j < dyn.length; j++) walk(dyn[j], out, depth + 1);
  }
  if (vnode.suspense) {
    walk(vnode.suspense.activeBranch, out, depth + 1);
    walk(vnode.suspense.pendingBranch, out, depth + 1);
  }
  if (vnode.ssContent) walk(vnode.ssContent, out, depth + 1);
}

var chain = [], node = input;
while (node) { chain.push(node); node = node.parentElement; }

var found = null, how = '', at = -1;

// Development builds tag the root element of every component instance.
for (var a = 0; a < chain.length; a++) {
  if (ownsValue(chain[a].__vueParentComponent)) {
    found = chain[a].__vueParentComponent; how = '__vueParentComponent'; at = a;
    break;
  }
}

// Otherwise walk the mounted tree and keep the innermost value owner that wraps
// our input: chain is ordered from the input upwards, so the deepest wins.
if (!found) {
  var all = [];
  var root = document.getElementById('app') || document.body;
  walk(root._vnode, all, 0);
  for (var c = 0; c < all.length; c++) {
    if (!ownsValue(all[c])) continue;
    var index = chain.indexOf(rootEl(all[c]));
    if (index !== -1 && index > at) { found = all[c]; how = 'mounted tree'; at = index; }
  }
}

if (!found) return {ok: false, error: 'no component owning a value wraps this field'};

var value = new Date(year, month - 1, day, 0, 0, 0, 0);
var props = found.props || {};
try {
  found.emit('update:modelValue', value);
} catch (e) {
  return {ok: false, error: 'emit failed: ' + e};
}
return {
  ok: true,
  name: typeName(found) || 'component',
  how: how,
  dateFormat: props.dateFormat || '',
  emitted: 'update:modelValue'
};
"""

# Web components keep their state on the host element, so a date can be handed
# to a value property or a setter, and the change event told about it.
JS_PROBE_DATE_HOST = """
var input = arguments[0], host = input;
while (host && host.tagName.indexOf('-') === -1) host = host.parentElement;
if (!host) return null;

var proto = Object.getPrototypeOf(host);
var own = Object.keys(host);
var accessors = Object.getOwnPropertyNames(proto);
var isDateish = function (n) { return /value|date|model|select/i.test(n); };
return {
  el: host,
  tag: host.tagName.toLowerCase(),
  props: own.filter(isDateish),
  // Inherited event handlers such as onselect carry no state; skip them.
  accessors: accessors.filter(function (n) { return !/^on/.test(n) && isDateish(n); })
};
"""

JS_SET_HOST_VALUE = """
var host = arguments[0], name = arguments[1], shape = arguments[2], call = arguments[3];
var year = arguments[4], month = arguments[5], day = arguments[6];
var value = (shape === null) ? new Date(year, month - 1, day, 0, 0, 0, 0) : shape;
try {
  if (call) host[name].call(host, value); else host[name] = value;
} catch (e) {
  return String(e);
}
host.dispatchEvent(new Event('input', {bubbles: true}));
host.dispatchEvent(new Event('change', {bubbles: true}));
return '';
"""

# What the field ended up showing, and whether the calendar header is stuck on
# NaN - the sign that the component never understood the value.
JS_READ_FIELD_STATE = """
var input = arguments[0], out = {value: input.value || '', nan: false};
if (/nan/i.test(out.value)) out.nan = true;
var headers = document.querySelectorAll(
  '.dp__month_year_select, .dp__month_year_row, [class*="month_year"]');
for (var i = 0; i < headers.length; i++) {
  if (/nan/i.test(headers[i].textContent || '')) { out.nan = true; break; }
}
return out;
"""

# The calendar overlay stays open over the next fields, so ESC first and hide it
# outright when the component ignores that.
JS_HIDE_DATE_OVERLAY = """
var menus = document.querySelectorAll(
  '.dp__menu, [role="dialog"].dp__menu, .dp__menu_transitioned');
var hidden = 0;
for (var i = 0; i < menus.length; i++) {
  var style = window.getComputedStyle(menus[i]);
  if (style.display === 'none' || style.visibility === 'hidden') continue;
  menus[i].style.display = 'none';
  hidden++;
}
return hidden;
"""


# Controlled-component inputs ignore a plain value assignment, so write through
# the native setter and fire the events the framework listens to.
JS_SET_INPUT_VALUE = """
var input = arguments[0], value = arguments[1];
input.removeAttribute('readonly');
var setter = Object.getOwnPropertyDescriptor(
    window.HTMLInputElement.prototype, 'value').set;
setter.call(input, value);
input.dispatchEvent(new Event('input', {bubbles: true}));
input.dispatchEvent(new Event('change', {bubbles: true}));
"""

# Some fields only react to change/blur, so fire them once the text is in place.
JS_COMMIT_INPUT_VALUE = """
var input = arguments[0];
input.dispatchEvent(new Event('change', {bubbles: true}));
input.dispatchEvent(new Event('blur', {bubbles: true}));
"""


def render_date_mask(day, mask):
    """Render a date in the datepicker's own mask, e.g. dd/mm/yy.

    The component reports its dateFormat, so a typed value can be shaped
    exactly the way the field expects instead of guessing.
    """
    replacements = {
        "dd": f"{day.day:02d}",
        "d": str(day.day),
        "mmmm": MONTH_NAMES_ID[day.month],
        "mmm": MONTH_SHORT_ID[day.month],
        "MMMM": MONTH_NAMES_EN[day.month],
        "MMM": MONTH_NAMES_EN[day.month][:3],
        "mm": f"{day.month:02d}",
        "m": str(day.month),
        "MM": f"{day.month:02d}",
        "M": str(day.month),
        "yyyy": f"{day.year:04d}",
        "yy": f"{day.year % 100:02d}",
    }
    tokens = sorted(replacements, key=len, reverse=True)

    rendered, index = "", 0
    while index < len(mask):
        for token in tokens:
            if mask.startswith(token, index):
                rendered += replacements[token]
                index += len(token)
                break
        else:
            rendered += mask[index]
            index += 1
    return rendered


def read_date_state(driver, field, day):
    """Read back what the field shows, and whether the calendar is on NaN.

    That combination is the only trustworthy proof the date was accepted: a
    field can display the text while the component underneath still holds NaN.
    """
    state = driver.execute_script(JS_READ_FIELD_STATE, field) or {}
    value = (state.get("value") or "").strip()
    if (
        value
        and not state.get("nan")
        and "nan" not in value.lower()
        and str(day.day) in value
    ):
        return True, value
    return False, value


def close_datepicker(driver, field):
    """Dismiss the calendar overlay, which otherwise covers the next fields."""
    try:
        field.send_keys(Keys.ESCAPE)
        time.sleep(0.3)
    except Exception:
        pass

    if driver.execute_script(JS_HIDE_DATE_OVERLAY):
        log_debug("Calendar overlay was still open and has been hidden")


def set_date_via_component(driver, field, day, label):
    """Hand the date to the Vue component that owns this field's value.

    Returns (accepted, date_format). The component is found by its model rather
    than by name, since EHRM wraps several datepicker libraries.
    """
    result = driver.execute_script(
        JS_SET_DATE_VIA_MODEL, field, day.year, day.month, day.day
    )
    if not isinstance(result, dict) or not result.get("ok"):
        reason = (result or {}).get("error", "nothing found")
        log_debug(f"{label} component: {reason}")
        return False, ""

    mask = result.get("dateFormat") or ""
    log_debug(
        f"{label} component: {result.get('name')} via {result.get('how')}, dateFormat={mask or '?'}"
    )

    accepted, value = read_date_state(driver, field, day)
    if not accepted:
        # Some wrappers listen for a plain input event instead of the model.
        driver.execute_script(
            JS_SET_INPUT_VALUE, field, render_date_mask(day, mask) if mask else ""
        )
        time.sleep(0.5)
        accepted, value = read_date_state(driver, field, day)

    close_datepicker(driver, field)
    if accepted:
        log_info(f"{label} set through component: '{value}'")
    else:
        log_warn(f"{label} component did not take the date (field shows '{value}')")
    return accepted, mask


def set_date_via_host(driver, field, day, label):
    """Hand the date to a web component host, for libraries that expose one.

    Tries every value-looking property and setter on the host element with a
    Date first and the common string shapes after, and keeps whichever one the
    field renders.
    """
    host = driver.execute_script(JS_PROBE_DATE_HOST, field)
    if not host:
        log_debug(f"{label} host: no custom element wraps this field")
        return False

    log_debug(
        f"{label} host: <{host.get('tag')}> props={host.get('props')} accessors={host.get('accessors')}"
    )

    names = (
        ["value"]
        + [n for n in host.get("props", []) if n != "value"]
        + host.get("accessors", [])
    )
    shapes = [
        None,
        f"{day.year:04d}-{day.month:02d}-{day.day:02d}",
        f"{day.day:02d}-{day.month:02d}-{day.year:04d}",
        f"{day.day:02d}/{day.month:02d}/{day.year:04d}",
    ]

    for name in names:
        for call in (False, True):
            for shape in shapes:
                driver.execute_script(
                    JS_SET_HOST_VALUE,
                    host.get("el"),
                    name,
                    shape,
                    call,
                    day.year,
                    day.month,
                    day.day,
                )
                time.sleep(0.4)
                accepted, value = read_date_state(driver, field, day)
                if accepted:
                    close_datepicker(driver, field)
                    log_info(
                        f"{label} set through <{host.get('tag')}>.{name}: '{value}'"
                    )
                    return True

    log_warn(f"{label} host: nothing on <{host.get('tag')}> took the date")
    return False


def fill_date_literal(driver, field_xpath, value, label, wait=None):
    """Set a date field without touching the calendar.

    The value owner is written first, since these datepickers re-parse typed
    text and turn it into NaN. Typing is the fallback: the component's own
    format is tried first, then each DATE_STYLES format, and a field showing
    NaN counts as a rejection. The read-back check is what decides, so a date
    the form could not parse never reaches the submission.
    """
    wait = wait or WebDriverWait(driver, WAIT_SECONDS)
    field = locate(driver, field_xpath, wait)
    log_step(f"{label}: field located")

    day = parse_date(value)
    if day is None:
        if is_blank(value):
            raise ValueError(f"no date available for {label}")
        log_warn(f"could not read '{value}' as a date, using it verbatim")
        candidates = [" ".join(str(value).split())]
    else:
        candidates = [DATE_STYLES[style](day) for style in date_styles()]

    if day is not None:
        accepted, mask = set_date_via_component(driver, field, day, label)
        if not accepted:
            accepted = set_date_via_host(driver, field, day, label)
        if accepted:
            return True
        if mask:
            shaped = render_date_mask(day, mask)
            log_debug(f"{label} will be typed in component's format ({mask}): {shaped}")
            candidates.insert(0, shaped)

    for text in candidates:
        driver.execute_script(JS_SET_INPUT_VALUE, field, "")
        field.click()
        time.sleep(0.5)
        field.send_keys(text)
        time.sleep(0.5)
        close_datepicker(driver, field)

        if field.get_attribute("value") != text:
            driver.execute_script(JS_SET_INPUT_VALUE, field, text)
            time.sleep(0.5)

        if field.get_attribute("value") == text and "nan" not in text.lower():
            log_info(f"{label} accepted: '{text}'")
            driver.execute_script(JS_COMMIT_INPUT_VALUE, field)
            close_datepicker(driver, field)
            return True

        shown = field.get_attribute("value")
        log_warn(
            f"{label} rejected '{text}' (field shows '{shown}'), trying next format..."
        )

    close_datepicker(driver, field)
    raise RuntimeError(
        f"{label} did not accept any known date format; tried {candidates}"
    )


def dump_buttons(driver, limit=40):
    """Print the buttons currently on the page, so a broken locator can be
    re-derived in DevTools from what is actually rendered."""
    try:
        buttons = driver.find_elements(By.XPATH, "//button")
    except Exception as e:
        log_error(f"Could not list buttons: {e}")
        return

    log_debug(f"Buttons on page ({len(buttons)}):")
    for index, button in enumerate(buttons[:limit], start=1):
        text = " ".join((button.text or "").split())
        log_debug(f"  {index}. {text!r}")


def is_blank(value):
    """True for None, NaN and empty strings."""
    if value is None:
        return True
    if isinstance(value, str):
        return not value.strip()
    try:
        return bool(pd.isna(value))
    except (TypeError, ValueError):
        return False


def wait_for_scroll_settled(driver, timeout=2, samples=3, pause=0.2):
    """Block until scrolling has stopped.

    Themes often set scroll-behavior: smooth, so the scrollIntoView in
    safe_click is still animating when the click is dispatched - and a click
    aimed at the pre-scroll coordinates lands on whatever moved into them.
    """
    deadline = time.time() + timeout
    last = driver.execute_script("return window.pageYOffset;")
    stable = 0

    while time.time() < deadline and stable < samples:
        time.sleep(pause)
        current = driver.execute_script("return window.pageYOffset;")
        stable = stable + 1 if current == last else 0
        last = current

    return last


def neutralise_submit(driver, element, label):
    """Stop a bare <button> from submitting the form it sits in.

    <button> without a type attribute is type=submit by default, so a click can
    navigate the page away instead of running the app's own handler - which is
    what a click that "does nothing" looks like from here.
    """
    changed = driver.execute_script(
        "var el = arguments[0];"
        "if (el.tagName !== 'BUTTON' || el.getAttribute('type')) return false;"
        "el.setAttribute('type', 'button'); return true;",
        element,
    )
    if changed:
        log_debug(f"Set type=button on '{label}' so the click cannot submit the form")
    return changed


def safe_click(driver, element, label):
    """Click an element, working around the click being intercepted.

    Dropdown panels and datepicker overlays are often rendered outside the
    visible part of their container, so the browser reports a click on
    whatever sits at those coordinates (usually the next form field). Scrolling
    the element into view and waiting for that scroll to finish fixes that; a
    scripted click is the last resort, since it dispatches the event on the
    element itself.
    """
    try:
        driver.execute_script(
            "arguments[0].scrollIntoView({block: 'center'});", element
        )
        wait_for_scroll_settled(driver)
    except Exception:
        pass

    try:
        element.click()
        return True
    except Exception as e:
        log_warn(
            f"Click on '{label}' intercepted ({e.__class__.__name__}), trying JavaScript click..."
        )

    try:
        driver.execute_script("arguments[0].click();", element)
        return True
    except Exception as e:
        log_error(f"JavaScript click on '{label}' failed: {e}")
        return False


# is_displayed() is not enough: a button can be laid out, in the viewport and
# still be a decoy that swallows clicks - transparent, pointer-events: none, or
# simply covered by something else. Hit-testing the centre point is what proves
# the click would reach this element.
JS_IS_LIVE_CONTROL = """
var el = arguments[0], r = el.getBoundingClientRect();
if (r.width <= 0 || r.height <= 0) return false;
if (r.top < 0 || r.bottom > window.innerHeight) return false;

var node = el;
while (node && node.nodeType === 1) {
  var s = window.getComputedStyle(node);
  if (s.display === 'none' || s.visibility === 'hidden') return false;
  if (parseFloat(s.opacity || '1') === 0) return false;
  if (s.pointerEvents === 'none') return false;
  node = node.parentElement;
}

var top = document.elementFromPoint(r.left + r.width / 2, r.top + r.height / 2);
if (!top) return false;
return el === top || el.contains(top) || top.contains(el);
"""


def find_on_screen_button(driver, selector, wait, settle=0.4, samples=2):
    """Wait for a button matching selector that is really the live control.

    Several elements can carry the same label, and one that is transparent or
    covered accepts a click without doing anything. Only a visible, hit-testable
    button inside the viewport whose size has stopped changing is used, so the
    SPA has finished rendering before the click lands.
    """

    def live(_driver):
        for element in _driver.find_elements(By.XPATH, selector):
            if not element.is_displayed():
                continue
            if _driver.execute_script(JS_IS_LIVE_CONTROL, element):
                return element
        return None

    element = wait.until(live)

    last = element.rect
    for _ in range(samples):
        time.sleep(settle)
        if element.rect == last:
            break
        last = element.rect

    return element


def form_is_open(driver):
    """True when the 'Tambah Data' form is on the page.

    Any one of its fields is enough, and they are the only proof that a click on
    'Tambah Data' actually did something.
    """
    for xpath in ('//*[@id="vs3__combobox"]', X_SK_NUMBER):
        try:
            if driver.find_elements(By.XPATH, xpath):
                return True
        except Exception:
            continue
    return False


def select_option_with_scroll(
    driver, option_xpath, listbox_xpath, label, listbox_scroll=150, wait=None
):
    """Open a dropdown already clicked by the caller, then find and click an
    option, scrolling the listbox and retrying when the option is off-screen."""
    wait = wait or WebDriverWait(driver, WAIT_SECONDS)
    option = None

    for attempt in range(1, MAX_SCROLL_ATTEMPTS + 1):
        try:
            option = wait.until(
                EC.visibility_of_element_located((By.XPATH, option_xpath))
            )
            break
        except Exception:
            if attempt == MAX_SCROLL_ATTEMPTS:
                log_warn(f"Max scroll attempts reached for {label}")
                break
            log_debug(
                f"Option not found, scrolling dropdown container... (attempt {attempt}/{MAX_SCROLL_ATTEMPTS})"
            )
            try:
                listbox = driver.find_element(By.XPATH, listbox_xpath)
                driver.execute_script(
                    f"arguments[0].scrollBy(0, {listbox_scroll});", listbox
                )
            except Exception:
                log_debug("Could not find listbox, trying alternative scroll...")
                driver.execute_script("window.scrollBy(0, 100);")
            time.sleep(0.5)

    if option is None:
        return False

    return safe_click(driver, option, label)


def fill_datepicker(driver, field_xpath, date_xpath, label, month_nav=None, wait=None):
    """Fill a datepicker field by clicking the calendar day.

    month_nav is None, or (month_button_xpath, clicks, month_button_label).
    """
    wait = wait or WebDriverWait(driver, WAIT_SECONDS)

    field = locate(driver, field_xpath, wait)
    log_step(f"{label}: field located")
    # Click the input first so the widget is initialised.
    field.click()
    time.sleep(1)

    if month_nav is not None:
        month_xpath, clicks, month_label = month_nav
        button = locate(driver, month_xpath, wait)
        log_step(f"{month_label}: located")
        for step in range(1, clicks + 1):
            log_debug(f"Clicking {month_label.lower()} (click {step}/{clicks})")
            safe_click(driver, button, month_label.lower())
            time.sleep(1)

    date_button = locate(driver, date_xpath, wait)
    print(f"{label} button found!")
    safe_click(driver, date_button, f"{label} day")
    time.sleep(1)
    return True


def manual_step(func):
    """On an unexpected failure, report it, let the operator fix the page, and
    report the step as failed. Every page interaction uses this so a single
    broken selector cannot abort the whole run."""

    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except Exception as e:
            log_error(f"Error in {func.__name__}: {e}")
            pause("Press Enter to continue after fixing the issue...")
            return False

    return wrapper


# --- Session -----------------------------------------------------------------


def open_website():
    """Attach to the already-running Chrome. The operator must leave it on the
    EHRM page, so this deliberately does not navigate anywhere."""
    log_info("Opening EHRM website...")
    options = Options()
    options.add_experimental_option("debuggerAddress", DEBUGGER_ADDRESS)
    return webdriver.Chrome(options=options)


# --- Login and navigation (unused by main; kept for the manual login path) ----


@manual_step
def click_enter_button(driver):
    """Find and click 'Masuk' button"""
    wait = WebDriverWait(driver, WAIT_SECONDS)
    button = locate(driver, "button.btn.btn-primary", wait, By.CSS_SELECTOR)

    log_step("'Masuk' button located")
    log_debug(f"Text: {button.text}")

    button.click()
    log_info("Button clicked")
    time.sleep(SETTLE_SECONDS)
    return True


@manual_step
def click_sso_login_button(driver):
    """Find and click 'Login SSO' button"""
    wait = WebDriverWait(driver, WAIT_SECONDS)
    button = locate(driver, "button.btn.btn-warning", wait, By.CSS_SELECTOR)

    log_step("'Login SSO' button located")
    log_debug(f"Text: {button.text}")

    button.click()
    log_info("'Login SSO' button clicked")
    time.sleep(SETTLE_SECONDS)
    return True


@manual_step
def fill_username(driver, username):
    """Fill username field"""
    field = locate(driver, "UsernameId", by=By.ID)
    log_step("Username field located")
    field.clear()
    field.send_keys(username)
    log_info(f"Username filled")
    return True


@manual_step
def fill_password(driver, password):
    """Fill password field"""
    field = locate(driver, "PasswordId", by=By.ID)
    log_step("Password field located")
    field.clear()
    field.send_keys(password)
    log_info("Password filled")
    return True


@manual_step
def click_login_button(driver):
    """Find and click Login button"""
    wait = WebDriverWait(driver, WAIT_SECONDS)
    button = locate(
        driver, "button[type='submit'].btn.btn-primary", wait, By.CSS_SELECTOR
    )

    log_step("Login button located")
    button.click()
    log_info("Login button clicked")
    time.sleep(SETTLE_SECONDS)
    return True


@manual_step
def open_employee_service_page(driver):
    """Open Employee Service page (Layanan Kepegawaian Daftar Pegawai)"""
    url = f"{EHRM_BASE_URL}/layanan-kepegawaian/daftar-pegawai"
    log_info(f"Opening URL: {url}")
    driver.get(url)
    time.sleep(SETTLE_SECONDS)
    log_info("Employee Service page opened")
    return True


@manual_step
def fill_nip(driver, nip):
    """Fill NIP field"""
    field = locate(driver, X_NIP_FIELD)
    log_step("NIP field located")
    field.clear()
    field.send_keys(nip)
    log_info(f"NIP filled: {nip}")
    return True


@manual_step
def click_search_button(driver):
    """Find and click 'Search Data' button, falling back to a text match"""
    wait = WebDriverWait(driver, WAIT_SECONDS)
    for selector in (X_SEARCH_BUTTON, X_SEARCH_BUTTON_BY_TEXT):
        try:
            button = wait.until(EC.presence_of_element_located((By.XPATH, selector)))
        except Exception:
            continue
        log_step("'Search Data' button located")
        button.click()
        log_info(f"'Search Data' button clicked via {selector}")
        time.sleep(SETTLE_SECONDS)
        return True

    raise RuntimeError(
        f"Failed to find search button, tried {X_SEARCH_BUTTON} "
        f"and {X_SEARCH_BUTTON_BY_TEXT}"
    )


@manual_step
def click_view_employee_button(driver):
    """Find and click 'View Employee' button (Lihat Pegawai)"""
    wait = WebDriverWait(driver, WAIT_SECONDS)
    button = locate(driver, X_VIEW_EMPLOYEE, wait)
    log_step("'View Employee' button located")

    driver.execute_script("arguments[0].scrollIntoView(true);", button)
    time.sleep(1)

    button.click()
    log_info("'View Employee' button clicked")
    time.sleep(SETTLE_SECONDS)
    return True


@manual_step
def click_salary_button(driver):
    """Find and click 'Salary' button (Gaji)"""
    wait = WebDriverWait(driver, WAIT_SECONDS)
    button = locate(driver, X_SALARY_TAB, wait)
    log_step("'Salary' button located")
    button.click()
    log_info("'Salary' button clicked")
    time.sleep(SETTLE_SECONDS)
    return True


# --- Employee lookup ---------------------------------------------------------


def read_all_employees_from_excel():
    """Read every employee row from the workbook's PNS sheet.

    Picks the first *.xlsx in this directory with 'kgb' in the name, sorted so
    the choice is deterministic across machines.
    """
    try:
        matches = sorted(
            f for f in os.listdir(".") if f.endswith(".xlsx") and "kgb" in f.lower()
        )
        if not matches:
            log_error("Excel file with 'kgb' in name not found!")
            return None

        excel_file = matches[0]
        if len(matches) > 1:
            log_warn(f"{len(matches)} candidate workbooks, using {excel_file}")
        log_info(f"Excel file found: {excel_file}")

        df = pd.read_excel(excel_file, sheet_name=SHEET_NAME)
        if len(df) == 0:
            log_error(f"'{SHEET_NAME}' sheet in {excel_file} is empty!")
            return None

        log_info(f"Total employees found: {len(df)}")
        return df

    except Exception as e:
        log_error(f"Error reading Excel: {e}")
        return None


@manual_step
def open_employee_by_id(driver, employee_id):
    """Open the employee salary page directly by ID.

    More reliable than searching for the employee in the table. Returns True on
    success; a failure pauses for the operator.
    """
    if not employee_id:
        raise ValueError("Employee ID is missing")

    url = employee_salary_url(employee_id)
    log_info(f"Navigating to employee salary page: {url}")
    driver.get(url)
    time.sleep(SETTLE_SECONDS)
    return True


@manual_step
def click_add_data_button(driver):
    """Click 'Tambah Data' and confirm the form really opened.

    A click can be delivered to a decoy - a duplicate label, an off-screen
    copy, or a node whose Vue handler is not attached yet - without raising
    anything. So every click is checked against the form appearing, the button
    is re-located on each attempt because the SPA may have replaced it, and a
    click that achieves nothing is retried with the next selector.
    """
    wait = WebDriverWait(driver, WAIT_SECONDS)

    # First, ensure we're on the salary tab where the real "Tambah Data" lives
    try:
        tab = wait.until(EC.element_to_be_clickable((By.XPATH, X_SALARY_TAB)))
        safe_click(driver, tab, "salary tab")
        log_info("Salary tab activated")
        time.sleep(SETTLE_SECONDS)
    except Exception as e:
        log_warn(f"Could not activate salary tab: {e}")

    for attempt in range(1, ADD_DATA_ATTEMPTS + 1):
        for selector in ADD_DATA_SELECTORS:
            try:
                button = find_on_screen_button(
                    driver, selector, WebDriverWait(driver, 3)
                )
            except Exception:
                continue

            log_step(
                f"'Add Data' button found via {selector} (attempt {attempt}/{ADD_DATA_ATTEMPTS})"
            )
            neutralise_submit(driver, button, "'Add Data' button")
            before_url = driver.current_url
            if not safe_click(driver, button, "'Add Data' button"):
                continue
            log_info("'Add Data' button clicked")
            time.sleep(SETTLE_SECONDS + 2)  # Extra wait for form to render

            if driver.current_url != before_url:
                log_warn(
                    f"click navigated away to {driver.current_url}, the button submitted the form"
                )
            elif form_is_open(driver):
                log_info("'Tambah Data' form is open")
                return True
            else:
                log_warn("form did not open, that click went nowhere")

        log_warn(f"Attempt {attempt}/{ADD_DATA_ATTEMPTS} did not open the form")
        time.sleep(SETTLE_SECONDS)

    dump_buttons(driver)
    raise RuntimeError(
        "'Tambah Data' was found and clicked but the form never "
        f"opened; tried {ADD_DATA_SELECTORS} "
        f"{ADD_DATA_ATTEMPTS} times"
    )


# --- KGB form fields ---------------------------------------------------------


@manual_step
@manual_step
def fill_salary_change_type(driver):
    """Select 'Kenaikan Gaji Berkala' in the 'Jenis Perubahan Gaji' dropdown"""
    wait = WebDriverWait(driver, WAIT_SECONDS)

    dropdown = locate(driver, '//*[@id="vs3__combobox"]', wait)
    log_step("Salary change type dropdown located")
    if not safe_click(driver, dropdown, "salary change type dropdown"):
        log_error("could not open the salary change type dropdown")
        return False
    time.sleep(1)

    if not select_option_with_scroll(
        driver,
        "//*[@id='vs3__option-1']",
        "//*[@role='listbox']",
        "salary type",
        wait=wait,
    ):
        log_error("Could not find salary type option")
        return False

    log_info("'Kenaikan Gaji Berkala' selected")
    return True


@manual_step
def fill_sk_number(driver, sk_number):
    """Fill 'No. SK' field"""
    if is_blank(sk_number):
        raise ValueError("nomor_surat is missing from the workbook row")
    return send_keys(driver, X_SK_NUMBER, sk_number, "SK number")


@manual_step
def fill_sk_date(driver, date_value=None, sk_date_xpath=None, prev_month_clicks=1):
    """Fill 'Tanggal SK' by typing the date; the datepicker is only used when
    the field refuses every known format."""
    if not is_blank(date_value):
        try:
            return fill_date_literal(driver, X_SK_DATE, date_value, "SK date")
        except Exception as e:
            log_warn(f"Literal SK date failed ({e}), falling back to the datepicker...")

    return fill_datepicker(
        driver,
        field_xpath=X_SK_DATE,
        date_xpath=sk_date_xpath,
        label="SK date",
        month_nav=(X_SK_PREV_MONTH, prev_month_clicks, "Prev month button"),
    )


@manual_step
def fill_golongan(driver, golongan):
    """Fill 'Golongan' field"""
    if is_blank(golongan):
        raise ValueError("golongan is missing from the workbook row")

    wait = WebDriverWait(driver, WAIT_SECONDS)
    dropdown = locate(driver, '//*[@id="vs4__combobox"]', wait)
    log_step("Golongan dropdown located")
    if not safe_click(driver, dropdown, "golongan dropdown"):
        log_error("could not open the golongan dropdown")
        return False
    driver.execute_script("window.scrollBy(0, 100);")
    time.sleep(1)
    log_debug(f"Target golongan: {golongan}")

    option_xpath = f"//li[@role='option' and text()='{golongan}']"
    if not select_option_with_scroll(
        driver,
        option_xpath,
        '//*[@id="vs4__listbox"]',
        f"golongan '{golongan}'",
        listbox_scroll=200,
        wait=wait,
    ):
        log_error(f"Could not find golongan '{golongan}' in dropdown")
        return False

    log_info(f"Golongan '{golongan}' selected")
    return True


@manual_step
def fill_work_year_old(driver, tahun_masa_kerja_lama):
    """Fill 'Tahun Masa Kerja Lama' field"""
    return send_keys(driver, X_YEAR_OLD, tahun_masa_kerja_lama, "Old work year")


@manual_step
def fill_work_month_old(driver, bulan_masa_kerja):
    """Fill 'Bulan Masa Kerja Lama' field"""
    return send_keys(driver, X_MONTH_OLD, bulan_masa_kerja, "Old work month")


@manual_step
def fill_work_year_new(driver, tahun_masa_kerja_baru):
    """Fill 'Tahun Masa Kerja Baru' field"""
    return send_keys(driver, X_YEAR_NEW, tahun_masa_kerja_baru, "New work year")


@manual_step
def fill_work_month_new(driver, bulan_masa_kerja_baru):
    """Fill 'Bulan Masa Kerja Baru' field"""
    return send_keys(driver, X_MONTH_NEW, bulan_masa_kerja_baru, "New work month")


@manual_step
def fill_salary_new(driver, gaji_pokok_baru):
    """Fill 'Gaji Pokok Baru' field"""
    if is_blank(gaji_pokok_baru):
        raise ValueError("baru_gaji is missing from the workbook row")

    # The money input has no stable id; try the most specific locator first.
    for selector in (
        '//input[@placeholder="Gaji Pokok Baru"]',
        '//input[contains(@class, "v-money3")]',
        '//input[@id="219"]',
    ):
        try:
            field = locate(driver, selector)
        except Exception:
            continue
        log_step(f"New salary field located using selector: {selector}")
        field.clear()
        field.send_keys(str(gaji_pokok_baru))
        log_info(f"New salary filled: {gaji_pokok_baru}")
        return True

    raise RuntimeError("Could not find new salary field with any selector")


@manual_step
def fill_pejabat_penandatangan(driver, pejabat_penandatangan):
    """Fill 'Pejabat Penandatangan' field"""
    return send_keys(driver, X_PEJABAT, pejabat_penandatangan, "Pejabat penandatangan")


@manual_step
def fill_effective_date(
    driver, date_value=None, date_xpath=None, month_clicks=0, direction="next"
):
    """Fill 'Tanggal mulai' by typing the date; the datepicker is only used
    when the field refuses every known format."""
    if not is_blank(date_value):
        try:
            return fill_date_literal(
                driver, X_EFFECTIVE_DATE, date_value, "Effective date"
            )
        except Exception as e:
            log_warn(
                f"Literal effective date failed ({e}), falling back to the datepicker..."
            )

    if month_clicks and direction == "next":
        month_nav = (X_EFFECTIVE_NEXT_MONTH, month_clicks, "Next month button")
    elif month_clicks and direction == "prev":
        month_nav = (X_EFFECTIVE_PREV_MONTH, month_clicks, "Previous month button")
    else:
        month_nav = None

    driver.execute_script("window.scrollBy(0, 100);")
    return fill_datepicker(
        driver,
        field_xpath=X_EFFECTIVE_DATE,
        date_xpath=date_xpath,
        label="Effective date",
        month_nav=month_nav,
    )


@manual_step
def upload_kgb_file(driver, employee_name):
    """Attach the employee's KGB PDF from the kgb/ tree"""
    matches = sorted(
        glob.glob(
            os.path.join(KGB_FOLDER, "**", f"*{employee_name}*.pdf"), recursive=True
        )
    )
    if not matches:
        log_warn(f"No KGB file found for employee: {employee_name}")
        pause("Press Enter to continue after placing the correct file...")
        return False

    file_path = matches[0]
    log_info(f"Found KGB file: {file_path}")

    file_input = locate(driver, '//input[@type="file"]')
    log_step("File upload input located")

    file_input.send_keys(os.path.abspath(file_path))
    log_info(f"KGB file uploaded: {file_path}")
    return True


# --- Submission --------------------------------------------------------------


def verify_submission_success(driver):
    """Best-effort check that the new row was written.

    UNRELIABLE: the EHRM table re-sorts asynchronously, so this can report a
    failure for a row that was in fact saved. The pass/fail print in main()
    should not be relied on - confirm visually.
    """
    try:
        expected_username = os.getenv("EXPECTED_USERNAME")
        if not expected_username:
            print("Warning: EXPECTED_USERNAME not set in .env")
            pause("Press Enter to continue without username verification...")
            return True

        element = locate(driver, X_SUBMITTED_ROW)
        element_text = element.text
        log_debug(f"Element content found: {element_text}")
        log_debug(f"Checking for username: {expected_username}")

        if expected_username in element_text:
            log_info(
                f"Submission verification successful - {expected_username} found in element"
            )
            return True

        log_warn(f"{expected_username} not found in element content")
        log_warn(f"Expected username '{expected_username}' but found different content")
        pause("Press Enter to continue after checking the submission...")
        return False

    except Exception as e:
        log_error(f"Error verifying submission: {e}")
        log_warn("Expected element not found - submission may have failed")
        pause("Press Enter to continue after checking the submission status...")
        return False


def process_employee(
    driver,
    row,
    sk_date_xpath,
    prev_month_clicks,
    effective_date_xpath,
    effective_date_clicks,
    effective_date_direction,
):
    """Fill and submit the KGB form for one workbook row."""
    employee_data = row.to_dict()

    nip = employee_data.get("nip")
    if not nip:
        log_warn("No NIP found for this row; skipping")
        return

    log_info(f"NIP: {nip}")
    employee_id = str(
        employee_data.get("id") or employee_data.get("id_pegawai") or ""
    ).strip()
    if not employee_id:
        log_warn("No employee ID found; skipping this row.")
        return

    if not open_employee_by_id(driver, employee_id):
        return
    if not click_add_data_button(driver):
        return

    driver.execute_script("window.scrollBy(0, 500);")
    time.sleep(2)

    # Dates: .env wins so a whole run can be pinned to one SK date, otherwise
    # they come from the row itself. The field picks the format it accepts.
    sk_date_value = (os.getenv("SK_DATE_LITERAL") or "").strip() or employee_data.get(
        "tanggal_surat"
    )
    effective_date_value = (
        os.getenv("EFFECTIVE_DATE_LITERAL") or ""
    ).strip() or employee_data.get("baru_mulai_tanggal")
    log_info(
        f"SK date {sk_date_value!r} -> "
        f"{to_date_literal(sk_date_value)}, "
        f"effective date {effective_date_value!r} -> "
        f"{to_date_literal(effective_date_value)}"
    )

    if not fill_salary_change_type(driver):
        return
    if not fill_sk_number(driver, employee_data.get("nomor_surat")):
        return
    if not fill_sk_date(driver, sk_date_value, sk_date_xpath, prev_month_clicks):
        return

    driver.execute_script("window.scrollBy(0, 300);")
    time.sleep(1)
    if not fill_golongan(driver, employee_data.get("golongan")):
        return
    if not fill_work_year_old(
        driver, employee_data.get("acuan_terakhir_masa_kerja_tahun")
    ):
        return
    if not fill_work_month_old(
        driver, employee_data.get("acuan_terakhir_masa_kerja_bulan")
    ):
        return
    if not fill_work_year_new(driver, employee_data.get("baru_masa_kerja_tahun")):
        return
    if not fill_work_month_new(driver, employee_data.get("baru_masa_kerja_bulan")):
        return
    if not fill_salary_new(driver, employee_data.get("baru_gaji")):
        return

    driver.execute_script("window.scrollBy(0, 500);")
    time.sleep(1)

    if not fill_effective_date(
        driver,
        effective_date_value,
        effective_date_xpath,
        effective_date_clicks,
        effective_date_direction,
    ):
        return
    if not fill_pejabat_penandatangan(
        driver, employee_data.get("baru_jabatan_pejabat_penandatangan")
    ):
        return
    if not upload_kgb_file(driver, employee_data.get("nama")):
        return

    time.sleep(5)
    # Operator reviews the filled form before the submission goes through.
    pause("Press Enter to continue to next employee anyway...")

    wait = WebDriverWait(driver, WAIT_SECONDS)
    for selector in CONFIRM_SELECTORS:
        try:
            button = wait.until(EC.presence_of_element_located((By.XPATH, selector)))
            neutralise_submit(driver, button, "'Confirm' button")
            if safe_click(driver, button, "'Confirm' button"):
                log_info(f"'Confirm' button clicked via {selector}")
                break
        except Exception:
            continue
    else:
        raise RuntimeError("Confirm button not found with any selector")
    time.sleep(5)

    verify_submission_success(driver)
    log_info(f"Successfully processed employee with NIP: {nip}")


def main():
    """Main function"""
    load_dotenv()

    # Credentials are only needed by the manual login helpers above, which are
    # not part of this run: the attached Chrome session is already logged in.
    os.getenv("USERNAME_EHRM")
    os.getenv("PASSWORD_EHRM")

    sk_date_xpath = os.getenv("SK_DATE_XPATH")
    prev_month_clicks = int(os.getenv("SK_DATE_PREV_MONTH_CLICKS", "1"))
    log_warn("PASTIKAN SK_DATE DI .env SUDAH BENAR!")
    log_info(f"SKDATE_SOURCE_XPATH (env): {sk_date_xpath}")

    effective_date_xpath = os.getenv("EFFECTIVE_DATE_XPATH")
    effective_date_next_clicks = os.getenv("EFFECTIVE_DATE_NEXT_MONTH_CLICKS")
    effective_date_prev_clicks = os.getenv("EFFECTIVE_DATE_PREV_MONTH_CLICKS")
    effective_date_clicks = 0
    effective_date_direction = "next"

    if effective_date_next_clicks and effective_date_next_clicks.strip() != "":
        effective_date_clicks = int(effective_date_next_clicks)
        effective_date_direction = "next"
    elif effective_date_prev_clicks and effective_date_prev_clicks.strip() != "":
        effective_date_clicks = int(effective_date_prev_clicks)
        effective_date_direction = "prev"

    log_warn("PASTIKAN EFFECTIVE_DATE DI .env SUDAH BENAR!")
    log_info(f"EFFECTIVE_DATE_XPATH (env): {effective_date_xpath}")
    log_info(f"EFFECTIVE_DATE_NEXT_MONTH_CLICKS (env): {effective_date_next_clicks}")
    log_info(f"EFFECTIVE_DATE_PREV_MONTH_CLICKS (env): {effective_date_prev_clicks}")
    log_info(
        f"EFFECTIVE_DATE direction: {effective_date_direction}, "
        f"clicks: {effective_date_clicks}"
    )

    sk_date_literal = (os.getenv("SK_DATE_LITERAL") or "").strip()
    effective_date_literal = (os.getenv("EFFECTIVE_DATE_LITERAL") or "").strip()
    log_info(f"SK_DATE_LITERAL: {sk_date_literal or 'from workbook tanggal_surat'}")
    log_info(
        f"EFFECTIVE_DATE_LITERAL: "
        f"{effective_date_literal or 'from workbook baru_mulai_tanggal'}"
    )
    log_info(f"DATE_INPUT_STYLE (formats tried in order): {date_styles()}")

    driver = None

    try:
        employees_df = read_all_employees_from_excel()
        driver = open_website()

        if employees_df is not None:
            for index, row in employees_df.iterrows():
                log_info(f"\n{'=' * 60}")
                log_info(f"Processing employee {index + 1} of {len(employees_df)}")
                log_info(f"{'=' * 60}")

                try:
                    process_employee(
                        driver,
                        row,
                        sk_date_xpath,
                        prev_month_clicks,
                        effective_date_xpath,
                        effective_date_clicks,
                        effective_date_direction,
                    )
                except Exception as e:
                    log_error(f"Error processing employee {index + 1}: {e}")
                    continue

        # Keep the browser open
        pause("Press Enter to close browser...")

    except Exception as e:
        log_error(f"Error in main: {e}")

    finally:
        if driver:
            driver.quit()


if __name__ == "__main__":
    main()

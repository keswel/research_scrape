from datetime import date
import requests
from bs4 import BeautifulSoup, SoupStrainer
from dataclasses import dataclass
from pynput import keyboard
import time
import threading
import ctypes
import tkinter as tk
import csv
import subprocess
import sys
import os


def resource_path(relative):
    """Resolve a bundled data file. When frozen by PyInstaller the files are
    extracted to sys._MEIPASS; otherwise they sit next to this script."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, relative)


@dataclass
class Department:
    dept_id: str
    dept_name: str
    college: str

def get_session_id():
    result = subprocess.run(
        ["powershell", "-ExecutionPolicy", "Bypass", "-File", resource_path("get-session.ps1")],
        capture_output=True, text=True
    )
    for line in result.stdout.splitlines():
        if line.strip() == "QUIT":
            return "__QUIT__"
        if line.startswith("PHPSESSID:"):
            return line.split(":", 1)[1].strip()
    print("Failed to get session ID")
    print(result.stderr)
    return None

def fetch_proposal(session_id, pid):
    url = f"https://dawson2.utsarr.net/comal/osp/pages/proposal.php?pid={pid}"
    headers = {
        "Referer": "https://dawson2.utsarr.net/comal/osp/pages/search_proposal.php?id=0",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0 Safari/537.36",
    }
    try:
        r = requests.get(url, cookies={"PHPSESSID": session_id}, headers=headers, timeout=30)
    except requests.RequestException as e:
        print(f"Could not reach the server for PID {pid}. Check your connection and try again.")
        print(f"  (details: {e})")
        return
    if r.status_code == 200:
        parse_html(r.text)
    else:
        print(f"Failed to fetch PID {pid}: {r.status_code}")

departments = {}
departments_by_id = {}

def load_departments(path='COLLEGE_DATA.csv'):
    global departments, departments_by_id
    departments = {}
    departments_by_id = {}
    with open(path, 'r') as f:
        reader = csv.reader(f)
        next(reader)  # skip header
        for row in reader:
            if len(row) >= 3:
                d = Department(
                    dept_id=row[0].strip(),
                    dept_name=row[1].strip(),
                    college=row[2].strip(),
                )
                departments[d.dept_name.upper()] = d
                departments_by_id[d.dept_id.upper()] = d


# guard to prevent re-entrant typing runs
_listener_lock = threading.Lock()
_type_busy = False

# UI overlay that follows the cursor while Ctrl is held
_ui_root = None
_ui_popup = None

# tracks if data has been scraped and is ready for pasting
_buffer_full = False


# ---------- Low-level keyboard hook: blocks physical input on demand ----------
# WH_KEYBOARD_LL doesn't require admin. The hook callback inspects the
# LLKHF_INJECTED flag so our own pynput-synthesized keys still get through.
from ctypes import wintypes

_WH_KEYBOARD_LL = 13
_HC_ACTION = 0
_LLKHF_INJECTED = 0x10

class _KBDLLHOOKSTRUCT(ctypes.Structure):
    _fields_ = [
        ("vkCode", wintypes.DWORD),
        ("scanCode", wintypes.DWORD),
        ("flags", wintypes.DWORD),
        ("time", wintypes.DWORD),
        ("dwExtraInfo", ctypes.c_void_p),
    ]

_LowLevelKeyboardProc = ctypes.WINFUNCTYPE(
    ctypes.c_long, ctypes.c_int, wintypes.WPARAM, wintypes.LPARAM
)

_user32 = ctypes.WinDLL("user32", use_last_error=True)
_kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
_user32.SetWindowsHookExW.argtypes = [
    ctypes.c_int, _LowLevelKeyboardProc, wintypes.HINSTANCE, wintypes.DWORD,
]
_user32.SetWindowsHookExW.restype = wintypes.HHOOK
_user32.CallNextHookEx.argtypes = [
    wintypes.HHOOK, ctypes.c_int, wintypes.WPARAM, wintypes.LPARAM,
]
_user32.CallNextHookEx.restype = ctypes.c_long
_user32.GetMessageW.argtypes = [
    ctypes.c_void_p, wintypes.HWND, wintypes.UINT, wintypes.UINT,
]
_user32.GetMessageW.restype = ctypes.c_int
_kernel32.GetModuleHandleW.argtypes = [wintypes.LPCWSTR]
_kernel32.GetModuleHandleW.restype = wintypes.HMODULE

_input_block_active = threading.Event()
_kbd_hook_proc_ref = None  # keep alive — Windows holds a raw pointer


def _ll_keyboard_hook_thread():
    global _kbd_hook_proc_ref

    def _proc(nCode, wParam, lParam):
        if nCode == _HC_ACTION and _input_block_active.is_set():
            kb = ctypes.cast(lParam, ctypes.POINTER(_KBDLLHOOKSTRUCT)).contents
            if not (kb.flags & _LLKHF_INJECTED):
                return 1  # swallow real key
        return _user32.CallNextHookEx(None, nCode, wParam, lParam)

    _kbd_hook_proc_ref = _LowLevelKeyboardProc(_proc)
    hook_handle = _user32.SetWindowsHookExW(
        _WH_KEYBOARD_LL, _kbd_hook_proc_ref,
        _kernel32.GetModuleHandleW(None), 0,
    )
    if not hook_handle:
        print(f"Low-level keyboard hook failed: {ctypes.get_last_error()}")
        return
    # LL hooks need a message pump on the installing thread.
    msg = wintypes.MSG()
    while _user32.GetMessageW(ctypes.byref(msg), None, 0, 0) > 0:
        pass


def _start_input_blocker():
    threading.Thread(target=_ll_keyboard_hook_thread, daemon=True).start()


def _get_cursor_pos():
    class POINT(ctypes.Structure):
        _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]

    pt = POINT()
    ctypes.windll.user32.GetCursorPos(ctypes.byref(pt))
    return pt.x, pt.y


class _PopupOverlay:
    def __init__(self, root):
        self.root = root
        self.window = tk.Toplevel(root)
        self.window.withdraw()
        self.window.overrideredirect(True)
        self.window.attributes("-topmost", True)
        self.window.attributes("-alpha", 0.88)
        self.label = tk.Label(
            self.window,
            text="",
            bg="black",
            fg="white",
            padx=10,
            pady=4,
            font=("Segoe UI", 10, "bold"),
            justify=tk.LEFT,  
        )
        self.label.pack()
        self._visible = False

    def _move_to_cursor(self):
        x, y = _get_cursor_pos()
        self.window.geometry(f"+{x+10}+{y-30}")

    def show(self, text="Scraping..."):
        def _show():
            self.label.config(text=text)
            self._move_to_cursor()
            self.window.deiconify()
            self._visible = True
        self.root.after(0, _show)

    def update_text(self, text):
        def _update():
            self.label.config(text=text)
        self.root.after(0, _update)

    def hide(self):
        def _hide():
            if self._visible:
                self.window.withdraw()
                self._visible = False
        self.root.after(0, _hide)

_ui_ready = threading.Event()


def _ui_thread_main():
    global _ui_root, _ui_popup
    _ui_root = tk.Tk()
    _ui_root.withdraw()
    _ui_popup = _PopupOverlay(_ui_root)
    _ui_ready.set()
    _ui_root.mainloop()


def _start_ui():
    # start the Tk UI in a dedicated thread. Other threads can safely request
    # UI updates via `after` once _ui_ready is set.
    threading.Thread(target=_ui_thread_main, daemon=True).start()
    _ui_ready.wait(timeout=3)


def type_row_strict_tabs():
    """
    Simulate genuine Tab key presses (no inserted whitespace).
    Start with focus on column A of the target row. Hotkey: both Ctrl keys held.
    """
    global _type_busy, _buffer_full, project_data
    # prevent re-entry
    with _listener_lock:
        if _type_busy:
            print("Typing already in progress; skipping duplicate trigger.")
            return
        _type_busy = True
    try:
        p = project_data
        if p is None or not _buffer_full:
            # nothing scraped yet — bail before blocking input or typing a partial row
            if _ui_popup:
                _ui_popup.update_text("No data yet — scrape a PID first")
            return
        kbd = keyboard.Controller()
        if _ui_popup:
            _ui_popup.show("Scraping...")
        # reset buffer after paste starts
        _buffer_full = False
        project_data = None
        if _ui_popup and _ui_popup._visible:
            _ui_popup.update_text("Awaiting data…")
        # ensure Alt/Ctrl/Shift aren't held (prevents Alt+Tab, Ctrl+Tab, etc.)
        try:
            for ak in (
                keyboard.Key.alt, keyboard.Key.alt_l, keyboard.Key.alt_r,
                keyboard.Key.ctrl, keyboard.Key.ctrl_l, keyboard.Key.ctrl_r,
                keyboard.Key.shift, keyboard.Key.shift_l, keyboard.Key.shift_r,
            ):
                try:
                    kbd.release(ak)
                except Exception:
                    pass
            time.sleep(0.05)
        except Exception:
            pass

        # suppress physical keyboard input for the duration of typing.
        # Our synthesized keys pass through (LL hook checks LLKHF_INJECTED).
        _input_block_active.set()

        # helper: type text then Tab
        def type_and_tab(text):
            if text is None:
                text = ""
            kbd.type(str(text))
            time.sleep(0.05)
            kbd.press(keyboard.Key.tab); kbd.release(keyboard.Key.tab)
            time.sleep(0.05)
        
        # A: NOI Receipt
        type_and_tab(p.noi_receipt)
        # B: NOI #
        type_and_tab(p.noi_number)
        # C: PI Name
        type_and_tab(p.pi_name)
        # D: College/VP Unit
        type_and_tab(p.college)
        # E: Department Code
        type_and_tab(p.department_code)
        # F: Sponsor
        type_and_tab(p.sponsor)
        # G: Proposal Due Date
        type_and_tab(p.proposal_due_date)

        # Now at H (col 8). Move to Q (col 17) by sending 9 Tabs (H->...->Q)
        for _ in range(9):
            kbd.press(keyboard.Key.tab); kbd.release(keyboard.Key.tab)
            time.sleep(0.05)

        # Q: Sponsor Due Date (skip if same as Proposal Due Date)
        sponsor_due = "" if p.sponsor_due_date == p.proposal_due_date else "Sponsor Deadline is " + p.sponsor_due_date + ";"
        kbd.type(sponsor_due); time.sleep(0.05)

        # move down to next row and reset to column A so the next entry can start there
        kbd.press(keyboard.Key.enter); kbd.release(keyboard.Key.enter)
        time.sleep(0.05)
        for _ in range(10):
            kbd.press(keyboard.Key.shift); kbd.press(keyboard.Key.tab)
            kbd.release(keyboard.Key.tab); kbd.release(keyboard.Key.shift)
            time.sleep(0.05)

        if _ui_popup:
            _ui_popup.update_text("Done!")

    except Exception as e:
        print(f"Error typing/restoring row: {e}")
    finally:
        _input_block_active.clear()
        with _listener_lock:
            _type_busy = False


# start a listener that triggers `type_row_strict_tabs` when both Ctrl keys are pressed.
# the listener will only trigger once per press (holding the keys won't retrigger until released).
ctrl_keys_pressed = set()
ctrl_triggered = False


def _on_press(key):
    global ctrl_triggered
    try:
        if key == keyboard.Key.ctrl_l:
            # show the overlay when Left Ctrl is held (UI only, no action)
            if _ui_popup:
                if _buffer_full and project_data:
                    p = project_data
                    dept = f" {p.department_code}" if p.department_code else ""
                    popup_text = (
                        f"Hold Right Ctrl to paste…\n"
                        f"NOI: {p.noi_number}  |  {p.pi_name}\n"
                        f"{p.college}{dept}  |  {p.sponsor}\n"
                        f"Due: {p.proposal_due_date}"
                    )
                    _ui_popup.show(popup_text)
                else:
                    _ui_popup.show("Awaiting data…")
            ctrl_keys_pressed.add(key)
        elif key == keyboard.Key.ctrl_r:
            ctrl_keys_pressed.add(key)
    except Exception:
        return

    # only perform the paste action when both Ctrl keys are held.
    if keyboard.Key.ctrl_l in ctrl_keys_pressed and keyboard.Key.ctrl_r in ctrl_keys_pressed:
        if not ctrl_triggered:
            ctrl_triggered = True
            if _buffer_full and project_data:
                if _ui_popup:
                    _ui_popup.update_text("Pasting…")
                threading.Thread(target=type_row_strict_tabs, daemon=True).start()
            elif _ui_popup:
                _ui_popup.update_text("No data yet — scrape a PID first")


def _on_release(key):
    global ctrl_triggered
    try:
        if key == keyboard.Key.ctrl_l:
            # hide when Left Ctrl is released
            if _ui_popup:
                _ui_popup.hide()
            ctrl_keys_pressed.discard(key)
            ctrl_triggered = False
        elif key == keyboard.Key.ctrl_r:
            ctrl_keys_pressed.discard(key)
            ctrl_triggered = False
    except Exception:
        return


@dataclass
class Project:
    noi_receipt: str        # Col A: today's date
    noi_number: str         # Col B: NOI #
    pi_name: str            # Col C: PI Name
    college: str            # Col D: College/VP Unit
    department_code: str    # Col E: Department Code
    sponsor: str            # Col F: Sponsor
    proposal_due_date: str  # Col G: Proposal Due Date
    sponsor_due_date: str   # Col Q: Sponsor Due Date


def print_data(p, raw_college=None, center=None):
    rows = [
        ("A", "NOI Receipt",        p.noi_receipt),
        ("B", "NOI #",              p.noi_number),
        ("C", "PI Name",            p.pi_name),
        ("-", "Orig. College",      raw_college if raw_college is not None else ""),
        ("-", "Orig. Center",       center if center is not None else ""),
        ("D", "College",            p.college),
        ("E", "Department Code",    p.department_code),
        ("F", "Sponsor",            p.sponsor),
        ("G", "Proposal Due",       p.proposal_due_date),
        ("Q", "Sponsor Due",        p.sponsor_due_date),
    ]
    label_w = max(len(label) for _, label, _ in rows)
    value_w = max(len(str(val)) for _, _, val in rows)
    lines = [f" {col}  {label:<{label_w}}  {str(val):<{value_w}} " for col, label, val in rows]
    width = len(lines[0])
    print("┌" + "─" * width + "┐")
    for line in lines:
        print("│" + line + "│")
    print("└" + "─" * width + "┘")

# Resolves the college name to a standardized abbreviation based on known mappings. Defaults to "VPR" if no match is found.
def resolve_college_name(college_str):
    match college_str.upper():
        case "SCIENCES": return "COS"
        case "ARTS": return "COLFA"
        case "BUSINESS": return "ACOB"
        case "EDUCATION": return "COEHD"
        case "CAICC": return "CAICC"
        case "HCAP": return "HCAP"
        case "KCEID": return "KCEID"
        case _: return "VPR"

# Resolves the college name and department code based on center and raw college data, prioritizing center data when available. Returns a tuple of (college_name, dept_id).
def resolve_college_data(center, raw_college):
    if center != "NULL":
        # prioritize center data — format is "Name - CODE" (e.g. "Urban Education Institute - CTR068")
        code = center.rsplit(" - ", 1)[-1].strip().upper()
        hit = departments_by_id.get(code)
        if hit:
            return resolve_college_name(hit.college), hit.dept_id
    key = raw_college.strip().upper()
    hit = departments.get(key)
    if hit:
        return resolve_college_name(hit.college), hit.dept_id
    return "NULL", "NULL"

# Handles resolving the college and department code based on both center and raw college data, with specific overrides for certain cases.
def resolve_college(center, raw_college):
    college_name, dept_id = resolve_college_data(center, raw_college)

    # -- CTR011 and CTR071 are handled by KCEID as opposed to COS in data.
    if dept_id in ("CTR011", "CTR071"):
        return "KCEID", ""

    # -- COS keeps its dept code.
    if college_name == "COS":
        return college_name, dept_id

    # -- Edge case: "KCEID MECH, AERO, IND EGNR" with no center override keeps
    # its own dept_id (from the raw_college row).
    if center == "NULL" and raw_college.strip().upper() == "KCEID MECH, AERO, IND EGNR":
        return college_name, dept_id

    if center == "NULL" and raw_college == "NULL":
        print("Center & Raw College are NULL. Double check data.")

    return college_name, ""

def parse_html(html_data):
    try:
        strainer = SoupStrainer(id="intake-tab")
        soup = BeautifulSoup(html_data, features="lxml", parse_only=strainer)

        noi_number = soup.find("span", {"class": "text-primary"}).text.strip()

        pi_first_name = soup.find("input", {"id": "pi_first_name"})["value"].strip()
        pi_last_name = soup.find("input", {"id": "pi_last_name"})["value"].strip()
        pi_name = pi_first_name + " " + pi_last_name

        raw_college = soup.find("input", {"id": "pi_department"})["value"].strip()

        sponsor_select = soup.find("select", {"id": "sponsor_id_part0"})
        sponsor_selected = sponsor_select.find("option", {"selected": True}) if sponsor_select else None
        sponsor_text = sponsor_selected.text.strip() if sponsor_selected else ""
        if sponsor_text == "Other":
            sponsor_text = soup.find("input", {"id": "sponsor_other_part0"})["value"].strip()

        proposal_due_date = soup.find("input", {"id": "target_date"})["value"].strip()
        sponsor_due_date = soup.find("input", {"id": "submission_deadline"})["value"].strip()

        select = soup.find("select", {"id": "pi_center_id"})
        selected = select.find("option", {"selected": True})
        center = selected.text.strip() if selected else "NULL"

        college, department_code = resolve_college(center, raw_college)

        global project_data
        project_data = Project(
            noi_receipt=date.today().strftime("%m/%d/%Y"),
            noi_number=noi_number,
            pi_name=pi_name,
            college=college,
            department_code=department_code,
            sponsor=sponsor_text,
            proposal_due_date=proposal_due_date,
            sponsor_due_date=sponsor_due_date,
        )
        print_data(project_data, raw_college=raw_college, center=center)

        global _buffer_full
        _buffer_full = True
        if _ui_popup and _ui_popup._visible:
            _ui_popup.update_text("Hold Right Ctrl to paste…")

    except Exception as e:
        print(f"Error parsing HTML: {e}")


_BANNER = [
    "             _          ____        __",
    "  ___  ___  (_)______  / / /__ ____/ /____  ____",
    " / _ \\/ _ \\/ / __/ _ \\/ / / -_) __/ __/ _ \\/ __/",
    "/_//_/\\___/_/\\__/\\___/_/_/\\__/\\__/\\__/\\___/_/",
]


def print_banner():
    width = max(len(line) for line in _BANNER)
    print()
    print("\n".join(_BANNER))
    print()
    print("Enter your UTSA Login".center(width))
    print()
    sys.stdout.flush()


if __name__ == "__main__":
    # Give the app its own taskbar identity (distinct icon/grouping) rather than
    # being lumped in with generic python/console-host entries.
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("UTSA.NOICollector")
    except Exception:
        pass

    print_banner()
    load_departments(resource_path('COLLEGE_DATA.csv'))

    session_id = get_session_id()
    if session_id == "__QUIT__":
        print("Goodbye!")
        raise SystemExit(0)
    if session_id is None:
        raise SystemExit(1)

    _start_ui()
    _start_input_blocker()

    listener = keyboard.Listener(on_press=_on_press, on_release=_on_release)
    listener.start()

    while True:
        pid = input("Enter PID (q to quit): ").strip()
        if pid.lower() == "q":
            break
        if not pid:
            continue
        if not pid.isdigit():
            print("PID must be a number. Try again.")
            continue
        fetch_proposal(session_id, pid)


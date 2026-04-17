from datetime import date
import requests
from bs4 import BeautifulSoup, SoupStrainer
from dataclasses import dataclass
from pynput import keyboard
import pyperclip
import time
import threading
import ctypes
import tkinter as tk
import csv
import subprocess

@dataclass
class Department:
    dept_id: str
    dept_name: str
    college: str

def get_session_id():
    result = subprocess.run(
        ["powershell", "-File", "./get-session.ps1"],
        capture_output=True, text=True
    )
    for line in result.stdout.splitlines():
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
    r = requests.get(url, cookies={"PHPSESSID": session_id}, headers=headers)
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
        kbd = keyboard.Controller()
        if _ui_popup:
            _ui_popup.show("Scraping...")
        # reset buffer after paste starts
        _buffer_full = False
        project_data = None
        if _ui_popup and _ui_popup._visible:
            _ui_popup.update_text("Awaiting data…")
        # ensure Alt isn't held (prevents Alt+Tab when we send Tabs)
        try:
            for ak in (keyboard.Key.alt, keyboard.Key.alt_l, keyboard.Key.alt_r):
                try:
                    kbd.release(ak)
                except Exception:
                    pass
            time.sleep(0.05)
        except Exception:
            pass

        # save current clipboard
        try:
            old_clip = pyperclip.paste()
        except Exception:
            old_clip = None

        # go to K (10 Tabs from A) and copy it
        for _ in range(10):
            kbd.press(keyboard.Key.tab); kbd.release(keyboard.Key.tab)
            time.sleep(0.05)
        # Ctrl+C to copy K
        kbd.press(keyboard.Key.ctrl); kbd.press('c'); kbd.release('c'); kbd.release(keyboard.Key.ctrl)
        time.sleep(0.12)
        k_content = pyperclip.paste()

        # return to A (10 Shift+Tabs)
        for _ in range(10):
            kbd.press(keyboard.Key.shift); kbd.press(keyboard.Key.tab)
            kbd.release(keyboard.Key.tab); kbd.release(keyboard.Key.shift)
            time.sleep(0.05)

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

        # Restore K: move back from Q to K with 7 Shift+Tabs (stay on same row)
        for _ in range(7):
            kbd.press(keyboard.Key.shift); kbd.press(keyboard.Key.tab)
            kbd.release(keyboard.Key.tab); kbd.release(keyboard.Key.shift)
            time.sleep(0.05)

        # paste saved K content (Ctrl+V)
        if k_content is not None:
            pyperclip.copy(k_content)
            kbd.press(keyboard.Key.ctrl); kbd.press('v'); kbd.release('v'); kbd.release(keyboard.Key.ctrl)
            time.sleep(0.05)

        # move down to next row and reset to column A so the next entry can start there
        kbd.press(keyboard.Key.enter); kbd.release(keyboard.Key.enter)
        time.sleep(0.05)
        for _ in range(10):
            kbd.press(keyboard.Key.shift); kbd.press(keyboard.Key.tab)
            kbd.release(keyboard.Key.tab); kbd.release(keyboard.Key.shift)
            time.sleep(0.05)

        # restore user's clipboard
        if old_clip is not None:
            pyperclip.copy(old_clip)

        if _ui_popup:
            _ui_popup.update_text("Done!")

    except Exception as e:
        print(f"Error typing/restoring row: {e}")
    finally:
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
            if _ui_popup:
                _ui_popup.update_text("Scraping…")
            import threading
            threading.Thread(target=type_row_strict_tabs, daemon=True).start()


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


def print_data(p):
    print(
        p.noi_receipt
        + "\t" + p.noi_number
        + "\t" + p.pi_name
        + "\t" + p.college
        + "\t" + p.department_code
        + "\t" + p.sponsor
        + "\t" + p.proposal_due_date
        + "\t" + p.sponsor_due_date
    )

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
    
    # -- CTR071 is handled by KCEID as opposed to COS in data. 
    if dept_id == "CTR071": return "KCEID", "" 

    # -- KCEID MECH, AERO, IND, EGNR needs to have its dept_id
    if raw_college == "KCEID MECH, AERO, IND EGNR" and college_name == "KCEID":
        return college_name, dept_id
    
    # -- All COS departments with no center override should be resolved to COS with the correct dept code.
    if college_name == "COS": 
        return college_name, dept_id
    
    # -- Default: if no match on center or raw college, return resolved college name (e.g. COS) with NULL code, or VPR if no match at all.
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
        print_data(project_data)

        global _buffer_full
        _buffer_full = True
        if _ui_popup and _ui_popup._visible:
            _ui_popup.update_text("Hold Right Ctrl to paste…")

    except Exception as e:
        print(f"Error parsing HTML: {e}")


if __name__ == "__main__":
    load_departments('COLLEGE_DATA.csv')

    session_id = get_session_id()
    if session_id is None:
        raise SystemExit(1)

    _start_ui()

    listener = keyboard.Listener(on_press=_on_press, on_release=_on_release)
    listener.start()

    while True:
        pid = input("Enter PID (q to quit): ").strip()
        if pid.lower() == "q":
            break
        fetch_proposal(session_id, pid)


from http.server import HTTPServer, BaseHTTPRequestHandler
from datetime import date
from bs4 import BeautifulSoup, SoupStrainer
from dataclasses import dataclass, astuple
from pynput import keyboard
import pyperclip
import time
import threading
import ctypes
import tkinter as tk
import csv

# TODO: Add COS-specific mappings for centers like Chemistry, etc.

# DEPARTMENT / CENTER DATA LOGIC:
# if center selected, use center data (ACRONYM ONLY)
# e.g. if department=kceid but center=COS Physics, use COS as department and its corresponding code [we can call this center override].

# edge cases to be aware of. 
#   -If the department ends up being COS (via center override or department if not center selected), then ALWAYS put COS and then in the next column put the code. 
#   -If the department is KCEID Mechanical Engineering, use KCEID as the department and also place the corresponding code.
#    [the code should only be used for KCEID (only 1) and COS (every)]
#   -If the department is not in the known list and no center override, default to UNKNOWN.
#   -If the department is being overriden by a center which is unknown, default to UNKNOWN and no code.


# Load department mappings
ctr_to_dept = {}
with open('department_mappings.csv', 'r') as f:
    reader = csv.reader(f)
    next(reader)  # skip header
    for row in reader:
        if len(row) >= 2:
            ctr_to_dept[row[0].strip()] = row[1].strip()
known_departments = set(ctr_to_dept.values())


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
        )
        self.label.pack()
        self._visible = False

    def _move_to_cursor(self):
        x, y = _get_cursor_pos()
        # place slightly above and to the right of the cursor
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

def serve_data_to_clipboard():
    """
    Copy a single tab-separated row to the clipboard ready for Excel paste.

    Column layout (1-based):
    1: proposal date (today)
    2: pid
    3: pi_name
    4: pi_department (acronym)
    5: department code (empty)
    6: sponsor
    7: target_date
    8-10: empty
    11: submission_deadline (sponsor due date)
    """
    if project_data is None:
        print("No project loaded yet")
        return
    try:
        p = project_data
        cols = []
        # 1: proposal date
        cols.append(date.today().strftime("%m/%d/%Y"))
        # 2-4: pid, pi_name, pi_department
        cols.append(p.pid)
        cols.append(p.pi_name)
        cols.append(p.pi_department)
        # 5: department code
        cols.append(p.department_code)
        # 6-7: sponsor, target_date
        cols.append(p.sponsor)
        cols.append(p.target_date)
        # ensure we have at least 10 columns (previous layout) before adding extra gaps
        while len(cols) < 10:
            cols.append("")
        cols.extend([""] * 6)
        # append submission_deadline after the added empty columns
        # If submission_deadline is the same as target_date, leave it blank
        if p.submission_deadline != p.target_date:
            cols.append(p.submission_deadline)
        else:
            cols.append("")

        row = "\t".join(cols)
        pyperclip.copy(row)
        print("Copied tab-separated row to clipboard (11 columns).")
    except Exception as e:
        print(f"Error in serve_data_to_clipboard: {e}")

def type_row_strict_tabs():
    """
    Simulate genuine Tab key presses (no inserted whitespace).
    Start with focus on column A of the target row. Hotkey: Ctrl+Alt+V
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

        # helper: type text then Tab (uses real Tab key)
        def type_and_tab(text):
            if text is None:
                text = ""
            kbd.type(str(text))
            time.sleep(0.05)
            kbd.press(keyboard.Key.tab); kbd.release(keyboard.Key.tab)
            time.sleep(0.05)

        # A: proposal date
        type_and_tab(date.today().strftime("%m/%d/%Y"))
        # B: pid
        type_and_tab(p.pid)
        # C: pi_name
        type_and_tab(p.pi_name)
        # D: pi_department
        type_and_tab(p.pi_department)
        # E: department_code
        type_and_tab(p.department_code)
        # F: sponsor
        type_and_tab(p.sponsor)
        # G: target_date
        type_and_tab(p.target_date)

        # Now at H (col 8). Move to Q (col 17) by sending 9 Tabs (H->...->Q)
        for _ in range(9):
            kbd.press(keyboard.Key.tab); kbd.release(keyboard.Key.tab)
            time.sleep(0.05)

        # Q: submission_deadline (skip if same as target_date)
        q_value = "" if p.submission_deadline == p.target_date else p.submission_deadline
        kbd.type(q_value); time.sleep(0.05)

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

        print("Typed row using real Tabs and preserved column K.")
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
                if _buffer_full:
                    _ui_popup.show("Hold Right Ctrl to paste…")
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

# start UI (popup overlay for the hotkey)
_start_ui()

# start listener in background
listener = keyboard.Listener(on_press=_on_press, on_release=_on_release)
listener.start()

@dataclass
class Project:
    pid: str 
    pi_name: str 
    pi_department: str 
    department_code: str
    sponsor: str 
    target_date: str 
    submission_deadline: str 

    def __iter__(self):
        return iter(astuple(self))

class Handler(BaseHTTPRequestHandler):
    def print_data(self, p):
        print("Cleaned Data:")
        print(date.today().strftime("%m/%d/%Y")
              +"\t"+p.pid
              +"\t"+p.pi_name
              +"\t"+p.pi_department
              +"\t"+p.department_code
              +"\t"+p.sponsor
              +"\t"+p.target_date
              +"\t"+p.submission_deadline
              )
        return

    def department_decide(self, center, pi_department):
        # Handle special case for KCEID Mechanical Engineering
        if pi_department == "KCEID Mechanical Engineering":
            department = "KCEID"
        else:
            # Extract department from the first word
            department = pi_department.split()[0] if pi_department else ""
        
        # If department not in known list, default to VPR
        if department not in known_departments:
            department = "VPR"
        
        # Get department from center if available
        center_dept = ctr_to_dept.get(center, None)
        
        # If center corresponds to a different department, use center's department
        if center_dept and center_dept != department:
            department = center_dept
        
        # Determine department code
        if department in known_departments and ctr_to_dept.get(center) == department:
            code = center
        else:
            code = ""
        
        return department, code

    def parse_html(self, html_data):
        try:
            #strainer = SoupStrainer(["title", "span", "input", "a", "select", "div", "h3"])
            strainer = SoupStrainer(id="intake-tab")
            soup = BeautifulSoup(html_data, features="lxml", parse_only=strainer)

            proposal_id = soup.find("span", {"class": "text-primary"}).text.strip()
            pi_first_name = soup.find("input", {"id": "pi_first_name"})["value"].strip()
            pi_last_name = soup.find("input", {"id": "pi_last_name"})["value"].strip()
            pi_name = pi_first_name + " " + pi_last_name
            pi_department = soup.find("input", {"id": "pi_department"})["value"].strip()
            
            sponsor = soup.find("a", {"class": "chosen-single"})
            sponsor_text = sponsor.find("span").text.strip()
            if sponsor_text == "Other":
                sponsor_text = soup.find("input", {"id": "sponsor_other_part0"})["value"].strip()
            
            target_date = soup.find("input", {"id": "target_date"})["value"].strip()
            submission_deadline = soup.find("input", {"id": "submission_deadline"})["value"].strip()

            select = soup.find("select", {"id": "pi_center_id"})
            selected = select.find("option", {"selected": True})
            center = selected.text.strip() if selected else "none selected"
            
            pi_department = self.department_decide(center, pi_department)
            
            global project_data
            project_data = Project(proposal_id, pi_name, pi_department[0], pi_department[1], sponsor_text, target_date, submission_deadline)
            # pyperclip.copy(list(project_data)[0])
            self.print_data(project_data)
            serve_data_to_clipboard()
            # update UI if visible
            global _buffer_full
            _buffer_full = True
            if _ui_popup and _ui_popup._visible:
                _ui_popup.update_text("Hold Right Ctrl to paste…")

        except Exception as e:
            print(f"Error parsing HTML: {e}")

    def do_POST(self):
        length = int(self.headers['Content-Length'])
        html_data = self.rfile.read(length).decode()
        print(f"\n===== NEW PAGE =====")

        self.parse_html(html_data)

        # print(html_data)
        self.send_response(200)
        self.end_headers()

    def log_message(self, *args):
        pass

HTTPServer(('localhost', 3000), Handler).serve_forever()

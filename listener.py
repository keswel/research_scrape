from http.server import HTTPServer, BaseHTTPRequestHandler
from datetime import date
from re import match
import re
from bs4 import BeautifulSoup, SoupStrainer
from dataclasses import dataclass
from pynput import keyboard
import pyperclip
import time
import threading
import ctypes
import tkinter as tk
import csv

# TODO: 1. Add COS-specific mappings for centers like Chemistry, etc.
#          Specifically, when Department is COS with no override.  
#          A new CSV is needed with Department Names and ther corresponding center codes.
#       2. Refactor code, functions should only really have one responsibility (See, resolve_college and type_row_strict_tabs).
#       3. Fix department code handling
#          -- KCEID MECH ENG resolves to COS but with the correct code. should be KCEID <correct_code> 

# Load college mappings
ctr_to_college = {}
with open('department_mappings.csv', 'r') as f:
    reader = csv.reader(f)
    next(reader)  # skip header
    for row in reader:
        if len(row) >= 2:
            ctr_to_college[row[0].strip()] = row[1].strip()
known_colleges = set(ctr_to_college.values())


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


# start UI (popup overlay for the hotkey)
_start_ui()

# start listener in background
listener = keyboard.Listener(on_press=_on_press, on_release=_on_release)
listener.start()


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


class Handler(BaseHTTPRequestHandler):
    def print_data(self, p):
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

    def resolve_college(self, center, raw_college):
        # print(f"[resolve_college] inputs: center={center!r}, raw_college={raw_college!r}")
        # Normalize raw college to an acronym
        college = raw_college.split()[0] if raw_college else ""
        # print(f"[resolve_college] normalized college acronym: {college!r}")
        if college not in known_colleges:
            # print(f"[resolve_college] {college!r} not in known_colleges, defaulting to VPR")
            college = "VPR"
        # Extract center code and apply override if it maps to a different college
        department_code = ""
        code_match = re.search(r"CTR[0-9]{3}", center)
        if code_match:
            center_code = code_match.group()
            # print(f"[resolve_college] center code: {center_code!r}")
            center_college = ctr_to_college.get(center_code)
            # print(f"[resolve_college] center_college lookup: {center_code!r} -> {center_college!r}")
            if center_college:
                if center_college != college:
                    # print(f"[resolve_college] center override: {college!r} -> {center_college!r}")
                    college = center_college
                if college == "COS":  # only COS gets a center code
                    department_code = center_code
        # KCEID always gets AEN004, regardless of any center override
        if raw_college == "KCEID MECH, AERO, IND EGNR":
            department_code = "AEN004"
        # print(f"[resolve_college] result: college={college!r}, department_code={department_code!r}")
        return college, department_code

    def parse_html(self, html_data):
        try:
            strainer = SoupStrainer(id="intake-tab")
            soup = BeautifulSoup(html_data, features="lxml", parse_only=strainer)

            noi_number = soup.find("span", {"class": "text-primary"}).text.strip()

            pi_first_name = soup.find("input", {"id": "pi_first_name"})["value"].strip()
            pi_last_name = soup.find("input", {"id": "pi_last_name"})["value"].strip()
            pi_name = pi_first_name + " " + pi_last_name

            raw_college = soup.find("input", {"id": "pi_department"})["value"].strip()

            sponsor = soup.find("a", {"class": "chosen-single"})
            sponsor_text = sponsor.find("span").text.strip()
            if sponsor_text == "Other":
                sponsor_text = soup.find("input", {"id": "sponsor_other_part0"})["value"].strip()

            proposal_due_date = soup.find("input", {"id": "target_date"})["value"].strip()
            sponsor_due_date = soup.find("input", {"id": "submission_deadline"})["value"].strip()

            select = soup.find("select", {"id": "pi_center_id"})
            selected = select.find("option", {"selected": True})
            center = selected.text.strip() if selected else "none selected"

            college, department_code = self.resolve_college(center, raw_college)

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
            self.print_data(project_data)

            global _buffer_full
            _buffer_full = True
            if _ui_popup and _ui_popup._visible:
                _ui_popup.update_text("Hold Right Ctrl to paste…")

        except Exception as e:
            print(f"Error parsing HTML: {e}")

    def do_POST(self):
        length = int(self.headers['Content-Length'])
        html_data = self.rfile.read(length).decode()
        print(f"\n===== PROJECT FOUND =====")

        self.parse_html(html_data)

        self.send_response(200)
        self.end_headers()

    def log_message(self, *args):
        pass


HTTPServer(('localhost', 3000), Handler).serve_forever()

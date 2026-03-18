from http.server import HTTPServer, BaseHTTPRequestHandler
from datetime import date
from bs4 import BeautifulSoup, SoupStrainer
from dataclasses import dataclass, astuple
from pynput import keyboard
import pyperclip
import time
import threading


# TODO: Add logic for department according to institution. 

project_data = None 
# guard to prevent re-entrant typing runs
_listener_lock = threading.Lock()
_type_busy = False

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
        # 5: department code (skip -> empty)
        cols.append("")
        # 6-7: sponsor, target_date
        cols.append(p.sponsor)
        cols.append(p.target_date)
        # ensure we have at least 10 columns (previous layout) before adding extra gaps
        while len(cols) < 10:
            cols.append("")
        # user requested 6 more empty tabs between target_date and submission_deadline
        cols.extend([""] * 6)
        # append submission_deadline after the added empty columns
        cols.append(p.submission_deadline)

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
    global _type_busy
    # prevent re-entry
    with _listener_lock:
        if _type_busy:
            print("Typing already in progress; skipping duplicate trigger.")
            return
        _type_busy = True
    try:
        p = project_data
        kbd = keyboard.Controller()

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
        # E: skip (just Tab)
        kbd.press(keyboard.Key.tab); kbd.release(keyboard.Key.tab); time.sleep(0.03)
        # F: sponsor
        type_and_tab(p.sponsor)
        # G: target_date
        type_and_tab(p.target_date)

        # Now at H (col 8). Move to Q (col 17) by sending 9 Tabs (H->...->Q)
        for _ in range(9):
            kbd.press(keyboard.Key.tab); kbd.release(keyboard.Key.tab)
            time.sleep(0.05)

        # Q: submission_deadline
        kbd.type(p.submission_deadline); time.sleep(0.05)
        kbd.press(keyboard.Key.enter); kbd.release(keyboard.Key.enter)
        time.sleep(0.05)

        # Restore K: move back from Q to K with 7 Shift+Tabs
        for _ in range(7):
            kbd.press(keyboard.Key.shift); kbd.press(keyboard.Key.tab)
            kbd.release(keyboard.Key.tab); kbd.release(keyboard.Key.shift)
            time.sleep(0.05)

        # paste saved K content (Ctrl+V)
        if k_content is not None:
            pyperclip.copy(k_content)
            kbd.press(keyboard.Key.ctrl); kbd.press('v'); kbd.release('v'); kbd.release(keyboard.Key.ctrl)
            time.sleep(0.05)

        # restore user's clipboard
        if old_clip is not None:
            pyperclip.copy(old_clip)

        print("Typed row using real Tabs and preserved column K.")
    except Exception as e:
        print(f"Error typing/restoring row: {e}")
    finally:
        with _listener_lock:
            _type_busy = False

# Start a listener that triggers `type_row_strict_tabs` when left+right Ctrl
# are pressed simultaneously. The listener will only trigger once per press
# (holding the keys won't retrigger until released).
ctrl_keys_pressed = set()
ctrl_triggered = False

def _on_press(key):
    global ctrl_triggered
    try:
        if key == keyboard.Key.ctrl_l or key == keyboard.Key.ctrl_r:
            ctrl_keys_pressed.add(key)
    except Exception:
        return
    if keyboard.Key.ctrl_l in ctrl_keys_pressed and keyboard.Key.ctrl_r in ctrl_keys_pressed:
        if not ctrl_triggered:
            ctrl_triggered = True
            import threading
            threading.Thread(target=type_row_strict_tabs, daemon=True).start()

def _on_release(key):
    global ctrl_triggered
    try:
        if key == keyboard.Key.ctrl_l or key == keyboard.Key.ctrl_r:
            ctrl_keys_pressed.discard(key)
            ctrl_triggered = False
    except Exception:
        return

# start listener in background
listener = keyboard.Listener(on_press=_on_press, on_release=_on_release)
listener.start()

@dataclass
class Project:
    pid: str 
    pi_name: str 
    pi_department: str 
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
              +"\t"+p.sponsor
              +"\t"+p.target_date
              +"\t"+p.submission_deadline
              )
        return

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

            global project_data
            project_data = Project(proposal_id, pi_name, pi_department, sponsor_text, target_date, submission_deadline)
            # pyperclip.copy(list(project_data)[0])
            self.print_data(project_data)
            serve_data_to_clipboard()

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

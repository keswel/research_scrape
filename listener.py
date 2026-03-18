from http.server import HTTPServer, BaseHTTPRequestHandler
from datetime import date
from bs4 import BeautifulSoup
from dataclasses import dataclass, astuple
from pynput import keyboard


# TODO: Add logic for department according to institution. 

project_data = None 

data_index = 0
def on_ctrl_v():
    global data_index
    if project_data is None:
        print("No project loaded yet")
        return
    try:
        data_list = list(project_data)
        print(data_list[data_index])
        data_index = (data_index + 1) % 6
    except Exception as e:
        print(f"Error in on_ctrl_v: {e}")

hotkeys = keyboard.GlobalHotKeys({'<ctrl>+v': on_ctrl_v})
hotkeys.start()

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
            soup = BeautifulSoup(html_data, features="lxml")
            
            if not soup.title or soup.title.text.strip() != "UTSA Office of Sponsored Project Administration":
                return 
            
            heading = soup.find("div", {"class": "heading-block"})
            if not (heading and "Notice of Intent" in heading.find("h3").text):
                return

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

            global project_data, data_index
            project_data = Project(proposal_id, pi_name, pi_department, sponsor_text, target_date, submission_deadline)
            data_index = 0  # reset so ctrl+v starts from pid on each new page
            self.print_data(project_data)

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

from http.server import HTTPServer, BaseHTTPRequestHandler
from datetime import date
from bs4 import BeautifulSoup


# TODO: Add logic for department according to institution. 

class Handler(BaseHTTPRequestHandler):
    def print_data(self, pid, pi_name, pi_department, sponsor, target_date, submission_deadline):
        print("Cleaned Data:")
        print(date.today().strftime("%m/%d/%Y")
              +"\t"+pid
              +"\t"+pi_name
              +"\t"+pi_department
              +"\t"+sponsor
              +"\t"+target_date
              +"\t"+submission_deadline
              )
        return


    def parse_html(self, html_data):
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
        
        select = soup.find("select", {"id": "pi_center_id"})
        selected = select.find("option", {"selected": True})
        center = selected.text.strip() if selected else "none selected"

        self.print_data(proposal_id, pi_name, pi_department, sponsor_text, target_date, submission_deadline)

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

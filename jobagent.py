from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
from urllib.parse import quote_plus
import time
import re
from datetime import datetime
from dotenv import load_dotenv
from openpyxl import Workbook
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import html
import os
import traceback
import socket

load_dotenv()

# --- EMAIL CONFIG ---
SENDER_EMAIL = os.getenv("SENDER_EMAIL")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")
RECEIVER_EMAIL = os.getenv("RECEIVER_EMAIL", "laxmimaru66@gmail.com")
SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.gmail.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))

# --- CONFIG ---
JOB_KEYWORD = "MERN OR React OR node js"
JOB_LOCATIONS = "Hyderabad, Telangana, remote"


# Build regex patterns for each search term so matching is case-insensitive
# and can handle variants like Node.js / node js / nodejs.
SEARCH_PATTERNS = []
for term in [t.strip() for t in re.split(r'(?i)\s+OR\s+', JOB_KEYWORD) if t.strip()]:
    if term.lower() == "mern":
        SEARCH_PATTERNS.append(re.compile(r"\bmern\b", re.I))
    elif term.lower() == "react":
        SEARCH_PATTERNS.append(re.compile(r"\breact\b", re.I))
    elif term.lower() in ("node js", "nodejs", "node.js", "node"):
        SEARCH_PATTERNS.append(re.compile(r"\bnode(?:\.js| js|js)?\b", re.I))
    else:
        SEARCH_PATTERNS.append(re.compile(r"\b" + re.escape(term) + r"\b", re.I))

LOCATION_PATTERNS = [
    re.compile(r"\bhyderabad\b", re.I),
    re.compile(r"\btelangana\b", re.I),
    re.compile(r"\b(remote|work from home|wfh)\b", re.I),
]

COUNTRIES = ["India"]
DATE_POSTED = "24h"
EXPERIENCE_LEVELS = []
WORKPLACE_TYPES = ["1","2", "3"]

# DATE_POSTED codes:            # EXPERIENCE_LEVELS codes:          # WORKPLACE_TYPES codes:    
# "any" = Any time              # [] = All levels                   # [] = All types       
# "24h" = Past 24 hours         # ["1"] = Internship                # ["1"] = On-site  
# "week" = Past week            # ["2"] = Entry level               # ["2"] = Remote   
# "month" = Past month          # ["3"] = Associate                 # ["3"] = Hybrid   
                                # ["4"] = Mid-Senior level          
                                # ["5"] = Director                  
                                # ["6"] = Executive                 
           


# --- SCROLL ---
MAX_SCROLL_ATTEMPTS = 200
SCROLL_PAUSE = 5
DETAIL_PAUSE = 2

# --- SAFE FILENAMES ---
safe_keyword = JOB_KEYWORD.replace(" ", "_")
safe_exp = ",".join(EXPERIENCE_LEVELS).replace(",", "_") if EXPERIENCE_LEVELS else "all"
safe_workplace = ",".join(WORKPLACE_TYPES).replace(",", "_") if WORKPLACE_TYPES else "all"
safe_date = DATE_POSTED
safe_location = JOB_LOCATIONS.replace(" ", "_").replace(",", "_")

# --- SETUP CHROME ---
# --- SETUP CHROME ---
options = Options()
options.add_argument("--start-maximized")
options.add_argument("--incognito")

# If running on GitHub cloud servers, force headless mode to prevent a visual window crash
if os.getenv("GITHUB_ACTIONS") == "true":
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")

driver = webdriver.Chrome(options=options)


# --- HELPER FUNCTIONS ---
def build_linkedin_url(keyword, location, exp_levels, workplace_types, date_posted):
    exp_param = ",".join(exp_levels) if exp_levels else ""
    workplace_param = ",".join(workplace_types) if workplace_types else ""
    date_param = ""
    if date_posted == "24h": date_param = "r86400"
    elif date_posted == "week": date_param = "r604800"
    elif date_posted == "month": date_param = "r2592000"

    # Split locations and join them with '%2C' for the URL
    quoted_locations = "%2C".join(quote_plus(loc.strip()) for loc in location.split(','))
    url = f"https://www.linkedin.com/jobs/search/?keywords={quote_plus(keyword)}&location={quoted_locations}"
    if exp_param: url += f"&f_E={exp_param}"
    if workplace_param: url += f"&f_WT={workplace_param}"
    if date_param: url += f"&f_TPR={date_param}"
    url += "&position=1&pageNum=0"
    return url

def scroll_page(driver):
    attempt = 0
    last_height = driver.execute_script("return document.body.scrollHeight")
    while attempt < MAX_SCROLL_ATTEMPTS:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(SCROLL_PAUSE)
        try:
            show_more_btn = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.CLASS_NAME, "infinite-scroller__show-more-button"))
            )
            show_more_btn.click()
            time.sleep(SCROLL_PAUSE)
        except:
            pass
        new_height = driver.execute_script("return document.body.scrollHeight")
        if new_height == last_height:
            break
        last_height = new_height
        attempt += 1

def fetch_job_details(job_url):
    job_desc = ""
    company_desc = ""
    if not job_url: 
        return job_desc, company_desc
    try:
        driver.get(job_url)
        time.sleep(DETAIL_PAUSE)
        WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.CLASS_NAME, "description__text"))
        )
        job_soup = BeautifulSoup(driver.page_source, "html.parser")
        job_div = job_soup.find("div", class_="description__text")
        job_desc = job_div.get_text(separator="\n", strip=True) if job_div else ""
        company_div = job_soup.find("div", class_="show-more-less-html__markup")
        company_desc = company_div.get_text(separator="\n", strip=True) if company_div else ""
    except Exception as e:
        print(f"⚠️ Failed to fetch job detail: {e}")
    return job_desc, company_desc


def build_excel_workbook(jobs):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Jobs"

    headers = list(jobs[0].keys())
    sheet.append(headers)
    for job in jobs:
        sheet.append([job.get(field, "") for field in headers])
    return workbook


def send_job_email(jobs, sender, password, receiver):
    if not jobs:
        print("⚠️ No jobs to email.")
        return False
    missing = []
    if not sender: missing.append('SENDER_EMAIL')
    if not password: missing.append('EMAIL_PASSWORD')
    if missing:
        print(f"⚠️ Email credentials missing: {', '.join(missing)}. Skipping email send.")
        return False

    # Debug summary
    masked_sender = sender
    try:
        local, domain = sender.split('@', 1)
        masked_sender = f"{local[:2]}***@{domain}"
    except Exception:
        pass
    print(f"ℹ️ Email debug: sender={masked_sender}, receiver={receiver}")
    print(f"ℹ️ SMTP server: {SMTP_SERVER}:{SMTP_PORT}")
    print(f"ℹ️ EMAIL_PASSWORD set: {'yes' if password else 'no'}")
    debug_level = os.getenv('EMAIL_DEBUG', '0')
    print(f"ℹ️ EMAIL_DEBUG={debug_level}")

    # Optional quick connectivity test
    try:
        print(f"ℹ️ Testing TCP connection to {SMTP_SERVER}:{SMTP_PORT}...")
        sock = socket.create_connection((SMTP_SERVER, SMTP_PORT), timeout=10)
        sock.close()
        print("✅ TCP connection successful")
    except Exception as e:
        print(f"⚠️ TCP connection failed: {e}")

    headers = list(jobs[0].keys())
    email_headers = ["SNO"] + headers
    # Build HTML table
    table_rows = []
    for index, job in enumerate(jobs, start=1):
        cols = [html.escape(str(index))] + [html.escape(str(job.get(h, ""))) for h in headers]
        table_rows.append("<tr>" + "".join(f"<td>{c}</td>" for c in cols) + "</tr>")
    table_html = (
        "<table border=\"1\" cellpadding=\"4\" cellspacing=\"0\">"
        + "<thead><tr>"
        + "".join(f"<th>{html.escape(h)}</th>" for h in email_headers)
        + "</tr></thead><tbody>"
        + "".join(table_rows)
        + "</tbody></table>"
    )

    html_content = f"<html><body><p>Found {len(jobs)} jobs.</p>{table_html}</body></html>"

    msg = MIMEMultipart("alternative")
    msg["From"] = sender
    msg["To"] = receiver
    msg["Subject"] = f"LinkedIn Job Scraper Results ({len(jobs)} jobs)"
    msg.attach(MIMEText(html_content, "html"))

    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=30) as server:
            # enable debuglevel to print SMTP protocol exchange when requested
            if debug_level == '1':
                server.set_debuglevel(1)
            server.ehlo()
            server.starttls()
            server.ehlo()
            server.login(sender, password)
            server.send_message(msg)
        print(f"✅ Email sent to {receiver}")
        return True
    except Exception as e:
        print(f"⚠️ Failed to send email: {e}")
        traceback.print_exc()
        return False


# --- MAIN SCRAPING LOOP ---
all_jobs = []

# No need to loop through countries, as the location is now combined
print(f"=== Scraping LinkedIn Jobs for {JOB_LOCATIONS} ===")
url = build_linkedin_url(JOB_KEYWORD, JOB_LOCATIONS, EXPERIENCE_LEVELS, WORKPLACE_TYPES, DATE_POSTED)
print(f"🔗 URL: {url}")
driver.get(url)
scroll_page(driver)

page_html = driver.page_source
soup = BeautifulSoup(page_html, "html.parser")
job_cards = soup.find_all("div", class_="base-card") 
    
for idx, card in enumerate(job_cards):
        a_tag = card.find("a", class_="base-card__full-link")
        job_url = a_tag["href"].strip() if a_tag else ""
        job_title = a_tag.find("span", class_="sr-only").text.strip() if a_tag and a_tag.find("span", class_="sr-only") else ""
        company_tag = card.find("h4", class_="base-search-card__subtitle")
        company_a = company_tag.find("a") if company_tag else None
        company_name = company_a.text.strip() if company_a else ""
        company_url = company_a["href"].strip() if company_a else ""
        location = card.find("span", class_="job-search-card__location")
        location = location.text.strip() if location else ""
        benefit = card.find("span", class_="job-posting-benefits__text")
        benefit = benefit.text.strip() if benefit else ""
        posted = card.find("time", class_="job-search-card__listdate")
        posted = posted.text.strip() if posted else ""

        if not any(pattern.search(location) for pattern in LOCATION_PATTERNS):
                    print(f"⚠️ Skipping job because location is not Hyderabad, Telangana or remote: {location}")
                    continue

        print(f"🔍 ({idx + 1}/{len(job_cards)}) Fetching job: {job_title}")
        job_description, company_description = fetch_job_details(job_url)

        combined_text = " ".join([
            job_title,
            company_name,
            location,
            benefit,
            posted,
            job_description,
            company_description
        ])
        if not any(pattern.search(combined_text) for pattern in SEARCH_PATTERNS):
            print(f"⚠️ Skipping job because it does not match case-insensitive terms: {JOB_KEYWORD}")
            continue

        all_jobs.append({
            "country": "India", # Now fixed to India as we are searching specific locations,
            "job_title": job_title,
            "company_name": company_name,
            "company_url": company_url,
            "location": location,
            "benefit": benefit,
            "posted": posted,
            "company_description": company_description,
            "job_url": job_url,
            "job_description": job_description
        })

# --- SAVE TO EXCEL ---
if all_jobs:
    base_name = f"linkedin_jobs_{safe_keyword}_{safe_location}_{safe_exp}_{safe_workplace}_{safe_date}"
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    seq = 1
    while True:
        excel_file = f"{base_name}_{timestamp}_{seq}.xlsx"
        if not os.path.exists(excel_file):
            break
        seq += 1

    workbook = build_excel_workbook(all_jobs)
    # Send jobs as HTML table in email (no attachment)
    send_job_email(all_jobs, SENDER_EMAIL, EMAIL_PASSWORD, RECEIVER_EMAIL)

    workbook.save(excel_file)
    print(f"📁 Saved {len(all_jobs)} jobs to {excel_file}")
else:
    print("⚠️ No jobs extracted.")

driver.quit()
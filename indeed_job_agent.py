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

# Exact regex translation from your original script
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
    re.compile(r"\b(remote|work from home|wfh)\b", re.I),
]

COUNTRIES = ["India"]
DATE_POSTED = "24h"
EXPERIENCE_LEVELS = []
WORKPLACE_TYPES = ["1", "2", "3"]

# --- SCROLL & PAUSE ---
MAX_SCROLL_ATTEMPTS = 5
SCROLL_PAUSE = 4
DETAIL_PAUSE = 3
MAX_JOBS = 10

# --- SAFE FILENAMES ---
safe_keyword = JOB_KEYWORD.replace(" ", "_")
safe_exp = ",".join(EXPERIENCE_LEVELS).replace(",", "_") if EXPERIENCE_LEVELS else "all"
safe_workplace = ",".join(WORKPLACE_TYPES).replace(",", "_") if WORKPLACE_TYPES else "all"
safe_date = DATE_POSTED

# --- SETUP CHROME ---
options = Options()
options.add_argument("--start-maximized")
options.add_argument("--incognito")
# Indeed applies strong bot-checking checks; a robust user agent completely stabilizes target tracking
options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
driver = webdriver.Chrome(options=options)

# --- HELPER FUNCTIONS ---
def build_indeed_url(keyword, location, date_posted):
    # Build query so Indeed treats terms as alternatives (OR),
    # and quote multi-word terms like "node js" so they remain a single unit.
    terms = [t.strip() for t in re.split(r'(?i)\s+OR\s+', keyword) if t.strip()]
    quoted_terms = [f'"{t}"' if ' ' in t else t for t in terms]
    query_str = ' OR '.join(quoted_terms)
    # Use the jobs search path and proper query/location parameters
    url = f"https://www.indeed.com/jobs?q={quote_plus(query_str)}&l={quote_plus(location)}"
    if date_posted == "24h":
        url += "&fromage=1"  # past 24 hours
    elif date_posted == "week":
        url += "&fromage=7"
    elif date_posted == "month":
        url += "&fromage=30"
    return url

def fetch_indeed_job_details(job_url):
    job_desc = ""
    company_desc = ""
    if not job_url: 
        return job_desc, company_desc
    try:
        driver.get(job_url)
        time.sleep(DETAIL_PAUSE)
        
        # Explicit wait tracking for Indeed's modern layout content element markers
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "jobDescriptionText"))
        )
        job_soup = BeautifulSoup(driver.page_source, "html.parser")
        
        desc_div = job_soup.find("div", id="jobDescriptionText")
        job_desc = desc_div.get_text(separator="\n", strip=True) if desc_div else ""
        
        # Indeed does not reliably load isolated company detail fields inside the description page viewport,
        # so this maps a placeholder safely to prevent array row matching faults.
        company_desc = "Company description details can be viewed directly on your target profile layout links."
    except Exception as e:
        print(f"⚠️ Failed to fetch job detail content text patterns: {e}")
    return job_desc, company_desc

def build_excel_workbook(jobs):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Jobs"
    if not jobs:
        return workbook
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

    headers = list(jobs[0].keys())
    email_headers = ["SNO"] + headers
    table_rows = []
    for index, job in enumerate(jobs, start=1):
        cols = [html.escape(str(index))] + [html.escape(str(job.get(h, ""))) for h in headers]
        table_rows.append("<tr>" + "".join(f"<td>{c}</td>" for c in cols) + "</tr>")
        
    table_html = (
        "<table border=\"1\" cellpadding=\"4\" cellspacing=\"0\" style=\"border-collapse: collapse; font-family: Arial, sans-serif;\">"
        + "<thead style=\"background-color: #2557a7; color: white;\"><tr>"
        + "".join(f"<th>{html.escape(h)}</th>" for h in email_headers)
        + "</tr></thead><tbody>"
        + "".join(table_rows)
        + "</tbody></table>"
    )

    html_content = f"<html><body><h3>Indeed Job Scraper Results ({len(jobs)} matches found)</h3><p>Found {len(jobs)} jobs.</p>{table_html}</body></html>"

    msg = MIMEMultipart("alternative")
    msg["From"] = sender
    msg["To"] = receiver
    msg["Subject"] = f"Ïndeed Job Scraper-{datetime.now().strftime('%Y-%m-%d')}"
    msg.attach(MIMEText(html_content, "html"))

    debug_level = os.getenv('EMAIL_DEBUG', '0')
    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=30) as server:
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
        print(f"⚠️ Failed to send email update packet: {e}")
        traceback.print_exc()
        return False

# --- MAIN SCRAPING LOOP ---
all_jobs = []
seen_urls = set()

# Locations to search: Hyderabad city and Remote — Indeed accepts full city strings (case-insensitive)
LOCATIONS = ["Hyderabad, Telangana", "Remote"]

print("=== Scraping Indeed Jobs ===")
for search_loc in LOCATIONS:
    print(f"--- Location: {search_loc} ---")
    url = build_indeed_url(JOB_KEYWORD, search_loc, DATE_POSTED)
    print(f"🔗 URL: {url}")
    driver.get(url)
    time.sleep(SCROLL_PAUSE)

    # Smooth card scrolling operation targeting dynamic lists containers
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight/2);")
    time.sleep(2)

    page_html = driver.page_source
    soup = BeautifulSoup(page_html, "html.parser")

    # Indeed tracks listings inside explicit result card content classes wrappers
    job_cards = soup.find_all("div", class_="job_seen_beacon")

    for idx, card in enumerate(job_cards):
        title_h2 = card.find("h2", class_="jobCardHeading")
        title_a = title_h2.find("a") if title_h2 else None
        job_title = title_a.get_text(strip=True) if title_a else ""

        jk_id = title_a["data-jk"].strip() if title_a and title_a.has_attr("data-jk") else ""
        job_url = f"https://www.indeed.com/viewjob?jk={jk_id}" if jk_id else ""

        company_span = card.find("span", attrs={"data-testid": "company-name"}) or card.find("span", data_testid="company-name")
        company_name = company_span.get_text(strip=True) if company_span else ""
        company_url = "N/A"

        # Robust location extraction
        location = ""
        loc_el = card.find(attrs={"data-testid": "text-location"})
        if not loc_el:
            loc_el = card.find("div", class_="companyLocation") or card.find("span", class_="companyLocation")
        if not loc_el:
            loc_el = card.find("div", class_="location") or card.find("span", class_="location")
        if loc_el:
            location = loc_el.get_text(strip=True)
        else:
            if title_a and title_a.has_attr("aria-label"):
                location = title_a["aria-label"].strip()

        posted_span = card.find("span", class_="date")
        posted = posted_span.get_text(strip=True) if posted_span else ""

        benefit_div = card.find("div", class_="metadata")
        benefit = benefit_div.get_text(strip=True) if benefit_div else "N/A"

        # Print raw location if skipping to aid debugging
        if not any(pattern.search(location) for pattern in LOCATION_PATTERNS):
            print(f"⚠️ Skipping job because location is not Hyderabad or Remote: {location!r}")
            continue

        if not job_url:
            print("⚠️ Skipping fetch: no job URL available for this card")
            continue

        if job_url in seen_urls:
            # already collected from another location search
            continue
        seen_urls.add(job_url)

        if len(all_jobs) >= MAX_JOBS:
            print(f"ℹ️ Reached MAX_JOBS={MAX_JOBS}; stopping further scraping")
            break

        print(f"🔍 ({idx + 1}/{len(job_cards)}) Fetching Indeed job: {job_title}")
        job_description, company_description = fetch_indeed_job_details(job_url)

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
            print(f"⚠️ Skipping job because inner text does not verify keywords match targets: {JOB_KEYWORD}")
            continue

        all_jobs.append({
            "country": "India",
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
    base_name = f"indeed_jobs_{safe_keyword}_India_{safe_exp}_{safe_workplace}_{safe_date}"
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    seq = 1
    while True:
        excel_file = f"{base_name}_{timestamp}_{seq}.xlsx"
        if not os.path.exists(excel_file):
            break
        seq += 1

    workbook = build_excel_workbook(all_jobs)

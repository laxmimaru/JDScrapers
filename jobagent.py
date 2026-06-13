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
# JOB_KEYWORD = "MERN OR React OR node js"
# Use Boolean logic: (MERN OR React OR node js) AND (Remote OR Hyderabad) to get both remote and Hyderabad jobs
JOB_KEYWORD = "(MERN OR React OR node js) AND (Remote OR Hyderabad)"
JOB_LOCATIONS = ""  # Empty to enable global search (remote jobs from anywhere, non-remote filtered by location)


# Build regex patterns for each search term so matching is case-insensitive
# and can handle variants like Node.js / node js / nodejs.
# Extract only the skill keywords from the Boolean logic for local filtering
SKILL_KEYWORDS = "MERN OR React OR node js"
SEARCH_PATTERNS = []
for term in [t.strip() for t in re.split(r'(?i)\s+OR\s+', SKILL_KEYWORDS) if t.strip()]:
    if term.lower() == "mern":
        SEARCH_PATTERNS.append(re.compile(r"\bmern\b", re.I))
    elif term.lower() == "react":
        SEARCH_PATTERNS.append(re.compile(r"\breact\b", re.I))
    elif term.lower() in ("node js", "nodejs", "node.js", "node"):
        SEARCH_PATTERNS.append(re.compile(r"\bnode(?:\.js| js|js)?\b", re.I))
    else:
        SEARCH_PATTERNS.append(re.compile(r"\b" + re.escape(term) + r"\b", re.I))

LOCATION_PATTERNS = [
    re.compile(r"hyderabad", re.I),
    re.compile(r"telangana", re.I),
    re.compile(r"serilingampalli", re.I),
    re.compile(r"india", re.I),
    re.compile(r"(remote|work from home|wfh|global)", re.I),
]

# Pattern to detect remote jobs specifically
REMOTE_PATTERN = re.compile(r"(remote|work from home|wfh|global)", re.I)

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

    # Build URL with Boolean keyword logic - no location parameter for global search
    url = f"https://www.linkedin.com/jobs/search/?keywords={quote_plus(keyword)}"
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

def fetch_job_details(job_url, job_location):
    job_desc = ""
    company_desc = ""
    if not job_url: 
        return job_desc, company_desc
    try:
        driver.get(job_url)
        # Increased wait time for robustness
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, "description__text"))
        )
        time.sleep(DETAIL_PAUSE) # Keep existing sleep
        job_soup = BeautifulSoup(driver.page_source, "html.parser")
        job_div = job_soup.find("div", class_="description__text")
        job_desc = job_div.get_text(separator="\n", strip=True) if job_div else ""
        company_div = job_soup.find("div", class_="show-more-less-html__markup")
        company_desc = company_div.get_text(separator="\n", strip=True) if company_div else ""
    except Exception as e:
        print(f"⚠️ Failed to fetch job detail for URL: {job_url} (Location: {job_location}). Error: {type(e).__name__} - {e} at {driver.current_url}")
        # Optionally, print full traceback for more detailed debugging
        # traceback.print_exc()
    return job_desc, company_desc


def is_remote_job(workplace_type):
    """Check if a job is remote based on workplace type from job card"""
    return workplace_type == "Remote"

def build_excel_workbook(jobs):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Jobs"

    headers = list(jobs[0].keys())
    sheet.append(headers)
    for job in jobs:
        sheet.append([job.get(field, "") for field in headers])
    return workbook


def send_job_email(jobs, job_type, sender, password, receiver):
    if not jobs:
        print(f"⚠️ No {job_type} jobs to email.")
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

    # Build HTML content with single table
    headers = list(jobs[0].keys())
    email_headers = ["SNO"] + headers
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
    
    html_content = f"<html><body><h2>{job_type.capitalize()} Jobs ({len(jobs)})</h2>{table_html}</body></html>"

    msg = MIMEMultipart("alternative")
    msg["From"] = sender
    msg["To"] = receiver
    msg["Subject"] = f"LinkedIn Job Listings - {job_type.capitalize()} ({len(jobs)} jobs)"
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
remote_jobs = []
non_remote_jobs = []
skipped_location = 0
skipped_keywords = 0

# No need to loop through countries, as the location is now combined
print(f"=== Scraping LinkedIn Jobs for {JOB_LOCATIONS} ===")
url = build_linkedin_url(JOB_KEYWORD, JOB_LOCATIONS, EXPERIENCE_LEVELS, WORKPLACE_TYPES, DATE_POSTED)
print(f"🔗 URL: {url}")
driver.get(url)
scroll_page(driver)

page_html = driver.page_source
soup = BeautifulSoup(page_html, "html.parser")
job_cards = soup.find_all("div", class_="base-card") 
print(f"📋 Total job cards found: {len(job_cards)}")
    
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
        
        # Extract workplace type from job card (Remote/On-site/Hybrid)
        workplace_type = ""
        workplace_elem = card.find("span", class_="job-search-card__workplace-type")
        if workplace_elem:
            workplace_type = workplace_elem.text.strip()
            print(f"🏢 Workplace type from element: {workplace_type}")
        else:
            # Try alternative method - check if workplace type is in the card text
            card_text = card.get_text()
            if "Remote" in card_text:
                workplace_type = "Remote"
            elif "On-site" in card_text:
                workplace_type = "On-site"
            elif "Hybrid" in card_text:
                workplace_type = "Hybrid"
            print(f"🏢 Workplace type from text fallback: {workplace_type}")

        # For non-remote jobs, apply location filter to ensure Hyderabad/India only
        # Remote jobs can be from anywhere (global)
        if not is_remote_job(workplace_type):
            location_match = any(pattern.search(location) for pattern in LOCATION_PATTERNS)
            print(f"🔍 Location filter check: {location} matches patterns: {location_match}")
            if not location_match:
                print(f"⚠️ Skipping non-remote job because location is not Hyderabad/Telangana/India: {location} (workplace_type: {workplace_type})")
                skipped_location += 1
                continue

        print(f"🔍 ({idx + 1}/{len(job_cards)}) Fetching job: {job_title} - {location}")
        job_description, company_description = fetch_job_details(job_url, location)

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
            skipped_keywords += 1
            continue

        job_data = {
            "job_title": job_title,
            "company_name": company_name,
            "company_url": company_url,
            "location": location,
            "benefit": benefit,
            "posted": posted,
            "company_description": company_description,
            "job_url": job_url,
            "job_description": job_description,
            "workplace_type": workplace_type
        }
        
        # Determine if job is remote based on workplace type from job card
        if is_remote_job(workplace_type):
            remote_jobs.append(job_data)
            print(f"✅ Added to remote jobs: {job_title} - {location} ({workplace_type})")
        else:
            non_remote_jobs.append(job_data)
            print(f"✅ Added to non-remote jobs: {job_title} - {location} ({workplace_type})")

# --- SAVE TO EXCEL AND SEND EMAILS ---
print(f"📊 Scraping complete. Remote jobs found: {len(remote_jobs)}, Non-remote jobs found: {len(non_remote_jobs)}")
print(f"📊 Skipped due to location filter: {skipped_location}")
print(f"📊 Skipped due to keyword filter: {skipped_keywords}")
if remote_jobs or non_remote_jobs:
    base_name = f"linkedin_jobs_{safe_keyword}_{safe_location}_{safe_exp}_{safe_workplace}_{safe_date}"
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    seq = 1
    
    # Save remote jobs to Excel
    if remote_jobs:
        while True:
            remote_excel_file = f"{base_name}_remote_{timestamp}_{seq}.xlsx"
            if not os.path.exists(remote_excel_file):
                break
            seq += 1
        remote_workbook = build_excel_workbook(remote_jobs)
        remote_workbook.save(remote_excel_file)
        print(f"📁 Saved {len(remote_jobs)} remote jobs to {remote_excel_file}")
        # Send email for remote jobs
        remote_email_sent = send_job_email(remote_jobs, "remote", SENDER_EMAIL, EMAIL_PASSWORD, RECEIVER_EMAIL)
        print(f"ℹ️ Remote email send status: {'Success' if remote_email_sent else 'Failed'}")
    
    # Save non-remote jobs to Excel
    if non_remote_jobs:
        seq = 1
        while True:
            non_remote_excel_file = f"{base_name}_non_remote_{timestamp}_{seq}.xlsx"
            if not os.path.exists(non_remote_excel_file):
                break
            seq += 1
        non_remote_workbook = build_excel_workbook(non_remote_jobs)
        non_remote_workbook.save(non_remote_excel_file)
        print(f"📁 Saved {len(non_remote_jobs)} non-remote jobs to {non_remote_excel_file}")
        # Add delay to avoid SMTP rate limiting
        print("ℹ️ Waiting 10 seconds before sending non-remote email to avoid rate limiting...")
        time.sleep(10)
        # Send email for non-remote jobs
        non_remote_email_sent = send_job_email(non_remote_jobs, "non-remote", SENDER_EMAIL, EMAIL_PASSWORD, RECEIVER_EMAIL)
        print(f"ℹ️ Non-remote email send status: {'Success' if non_remote_email_sent else 'Failed'}")
    
    total_jobs = len(remote_jobs) + len(non_remote_jobs)
    print(f"✅ Total: {total_jobs} jobs ({len(remote_jobs)} remote, {len(non_remote_jobs)} non-remote)")
else:
    print("⚠️ No jobs extracted.")

driver.quit()
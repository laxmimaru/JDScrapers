from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
from urllib.parse import quote_plus
import time
import csv
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime
from dotenv import load_dotenv
import os

load_dotenv()

# --- CONFIG ---
JOB_KEYWORD = "junior developer"
COUNTRIES = ["Belgium", "Netherlands"]
DATE_POSTED = "24h"   # "any", "24h", "week", "month"
EXPERIENCE_LEVELS = []  # [] all, ["2"] Entry, ["2","3"] Entry+Associate

# --- EMAIL ---
SENDER_EMAIL = os.getenv("SENDER_EMAIL")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")
RECEIVER_EMAIL = os.getenv("RECEIVER_EMAIL")

# --- SCROLL ---
MAX_SCROLL_ATTEMPTS = 200
SCROLL_PAUSE = 5
DETAIL_PAUSE = 2

# --- SAFE FILENAMES ---
safe_keyword = JOB_KEYWORD.replace(" ", "_")
safe_exp = ",".join(EXPERIENCE_LEVELS).replace(",", "_") if EXPERIENCE_LEVELS else "all"
safe_date = DATE_POSTED

# --- SETUP CHROME ---
options = Options()
options.add_argument("--start-maximized")
options.add_argument("--incognito")
driver = webdriver.Chrome(options=options)

# --- HELPER FUNCTIONS ---
def build_linkedin_url(keyword, location, exp_levels, date_posted):
    exp_param = ",".join(exp_levels) if exp_levels else ""
    date_param = ""
    if date_posted == "24h": date_param = "r86400"
    elif date_posted == "week": date_param = "r604800"
    elif date_posted == "month": date_param = "r2592000"

    url = f"https://www.linkedin.com/jobs/search/?keywords={quote_plus(keyword)}&location={quote_plus(location)}"
    if exp_param: url += f"&f_E={quote_plus(exp_param)}"
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
    if not job_url: return job_desc, company_desc
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

def send_job_email(jobs, sender, receiver, password):
    if not jobs: return False
    jobs_html = ""
    for job in jobs:
        jobs_html += f"""
        <div style="margin:10px 0; padding:10px; border-left:4px solid #2557a7; background:#fff; border-radius:5px;">
            <strong>{job['job_title']}</strong><br>
            <em>{job['company_name']}</em> - {job['location']}<br>
            <a href="{job['job_url']}" target="_blank">🔗 View Job</a><br>
            <small>Country: {job['country']}</small>
        </div>
        """
    html_content = f"""
    <html><body>
    <h2>LinkedIn Jobs - {len(jobs)} found</h2>
    <p>Date: {datetime.now().strftime('%Y-%m-%d')}</p>
    {jobs_html}
    <p><em>Automated Scraper • {datetime.now().strftime('%Y-%m-%d %H:%M')}</em></p>
    </body></html>
    """
    try:
        msg = MIMEMultipart()
        msg['From'] = sender
        msg['To'] = receiver
        msg['Subject'] = f"{len(jobs)} New LinkedIn Jobs - {datetime.now().strftime('%Y-%m-%d')}"
        msg.attach(MIMEText(html_content, 'html'))
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(sender, password)
            server.send_message(msg)
        print(f"✅ Email sent to {receiver}")
        return True
    except Exception as e:
        print("⚠️ Failed to send email:", e)
        return False

# --- MAIN SCRAPING LOOP ---
all_jobs = []

for country in COUNTRIES:
    print(f"=== Scraping LinkedIn Jobs for {country} ===")
    url = build_linkedin_url(JOB_KEYWORD, country, EXPERIENCE_LEVELS, DATE_POSTED)
    driver.get(url)
    scroll_page(driver)

    html = driver.page_source
    soup = BeautifulSoup(html, "html.parser")
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

        print(f"🔍 ({idx + 1}/{len(job_cards)}) Fetching job: {job_title}")
        job_description, company_description = fetch_job_details(job_url)

        all_jobs.append({
            "country": country,
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

# --- SAVE TO CSV ---
if all_jobs:
    csv_file = f"linkedin_jobs_{safe_keyword}_{'_'.join([c.replace(' ','') for c in COUNTRIES])}_{safe_exp}_{safe_date}.csv"
    with open(csv_file, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=all_jobs[0].keys())
        writer.writeheader()
        writer.writerows(all_jobs)
    print(f"📁 Saved {len(all_jobs)} jobs to {csv_file}")
else:
    print("⚠️ No jobs extracted.")

# --- SEND EMAIL ---
send_job_email(all_jobs, SENDER_EMAIL, RECEIVER_EMAIL, EMAIL_PASSWORD)

driver.quit()

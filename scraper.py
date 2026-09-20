import time
import random
import pandas as pd
from selenium import webdriver
from selenium.common import TimeoutException, StaleElementReferenceException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options

options = Options()

options.add_argument("--disable-blink-features=AutomationControlled")
options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/152.0.0.0 Safari/537.36")


driver = webdriver.Chrome(options=options)
wait = WebDriverWait(driver, 10)
driver.maximize_window()

def safe_text(parent, by, value):
    found = parent.find_elements(by, value)
    return found[0].text.strip() if found else ""

def safe_attr(parent, by, value, attr):
    found = parent.find_elements(by, value)
    return found[0].get_attribute(attr) if found else ""

def mask_phone(phone):
    if not phone:
        return ""
    phone = str(phone)
    if len(phone) < 4:
        return '*' * len(phone)
    return phone[:2] + '*' * (len(phone) - 3) + phone[-1]


BASE_URL = "https://www.google.com/maps/search/plumbers+in+toronto+canada/@43.719717,-79.542464"
OUTPUT_FILE = "plumbers_in_toronto.xlsx"
SAMPLE_FILE = "sample_output.xlsx"

records = []


driver.get(BASE_URL)
time.sleep(random.uniform(2,4))

panel = driver.find_element(By.CSS_SELECTOR, "div[role='feed']")
last_count = 0
jobs = []


for _ in range(20):

    try:
        jobs = wait.until(
            EC.presence_of_all_elements_located((
                By.CSS_SELECTOR, "div[role='article']"
            ))
        )


    except TimeoutException:
        print("Exiting...")
        break

    driver.execute_script('arguments[0].scrollTop = arguments[0].scrollHeight', panel)
    time.sleep(random.uniform(5, 10))

    if len(jobs) == last_count:
        break
    last_count = len(jobs)
    print(f"Loaded listings: {len(jobs)}")

for job in jobs:
    try:
        business_name = safe_text(job, By.CSS_SELECTOR, ".NrDZNb")
        is_sponser = safe_text(job, By.CSS_SELECTOR, ".jHLihd.lV5Ihd")
        job_title = safe_text(job, By.CSS_SELECTOR, ".W4Efsd:nth-of-type(1)>span:nth-of-type(1)>span:nth-of-type(1)")
        address = safe_text(job, By.CSS_SELECTOR, ".W4Efsd:nth-of-type(1)>span:nth-last-child(1)>span:nth-last-child(1):not(span:first-child)")
        phones = safe_text(job, By.CSS_SELECTOR, ".W4Efsd:nth-of-type(2) span:nth-of-type(2) span:nth-of-type(2)")
        rating = safe_text(job, By.CSS_SELECTOR, ".MW4etd")
        reviews = safe_text(job, By.CSS_SELECTOR, ".UY7F9").strip('(').strip(')')
        is_open = safe_text(job, By.CSS_SELECTOR, ".W4Efsd:nth-of-type(2)>span:nth-of-type(1)>span:nth-last-child(1)")
        website = safe_attr(job, By.CSS_SELECTOR, "a[data-value='Website']", attr="href")
        google_maps_link = safe_attr(job, By.CSS_SELECTOR, "a.hfpxzc", attr="href")

        is_sponser = "Yes" if is_sponser else "No"
        has_website = 'Yes' if website else 'No'


        records.append({
            'Business Name': business_name,
            'Ad': is_sponser,
            'Job': job_title,
            'Address': address,
            'Phone Number': phones,
            'Rating': rating,
            'Reviews': reviews,
            'Is Open': is_open,
            'Has Website': has_website,
            'Website': website,
            'Google Maps Link': google_maps_link
        })

    except (StaleElementReferenceException, TimeoutException):
        continue


df = pd.DataFrame(records)
df.drop_duplicates(subset=['Business Name', 'Phone Number', 'Google Maps Link'], inplace=True)

df["Rating"] = pd.to_numeric(df["Rating"], errors='coerce')
df["Reviews"] = pd.to_numeric(df["Reviews"], errors='coerce')

with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name='Plumber Jobs', index=False)
    ws = writer.book['Plumber Jobs']

    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = ws.dimensions

    for column in ws.columns:
        letter = column[0].column_letter
        longest = max((len(str(c.value)) for c in column if c.value), default = 0)
        ws.column_dimensions[letter].width = min(max(longest + 2, 10), 60)

driver.quit()

print(f'Saved {len(df)} businesses to {OUTPUT_FILE}')

sample = df.head(20).copy()
sample['Phone Number'] = sample['Phone Number'].apply(mask_phone)

with pd.ExcelWriter(SAMPLE_FILE, engine='openpyxl') as writer:
    sample.to_excel(writer, sheet_name='Plumber Jobs', index=False)
    ws = writer.book['Plumber Jobs']

    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = ws.dimensions

    for column in ws.columns:
        letter = column[0].column_letter
        longest = max((len(str(c.value)) for c in column if c.value), default = 0)
        ws.column_dimensions[letter].width = min(max(longest + 2, 10), 60)

print(f'Saved {len(sample)} businesses to {SAMPLE_FILE}')

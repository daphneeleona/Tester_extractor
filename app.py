import streamlit as st
import os
import time
import requests
import pandas as pd
from io import BytesIO
from datetime import datetime
import urllib3
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# 📄 Display selenium.log in Streamlit
def show_selenium_log():
    log_path = 'selenium.log'
    if os.path.exists(log_path):
        with open(log_path, 'r') as f:
            st.subheader("🔍 Selenium Log Output")
            st.code(f.read(), language='bash')
    else:
        st.warning("🚫 selenium.log not found.")

# 🚗 Use system-installed ChromeDriver
def get_driver():
    options = Options()
    options.add_argument('--headless=new')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--disable-gpu')
    options.add_argument('--window-size=1920,1080')
    options.add_argument('--disable-blink-features=AutomationControlled')
    options.add_argument('--ignore-certificate-errors')
    options.add_argument('--no-proxy-server')

    try:
        service = Service('/usr/bin/chromedriver')  # ✅ Pre-installed driver path
        driver = webdriver.Chrome(service=service, options=options)
        driver.implicitly_wait(10)
        return driver
    except Exception as e:
        st.error("⚠️ Could not launch Chrome. Check that Chrome and chromedriver are installed.")
        st.code(str(e), language='bash')
        return None

# 🌐 Load site with retries
def get_website_content(url, max_retries=2):
    for attempt in range(max_retries):
        driver = get_driver()
        if driver is None:
            st.warning("Retrying browser setup...")
            continue
        try:
            driver.get(url)
            time.sleep(3)
            return driver
        except Exception as e:
            st.warning(f"Attempt {attempt + 1} failed: {e}")
            driver.quit()
    return None

# 🎛️ Apply year/month filters
def select_filters(driver, wait, year, month):
    try:
        wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".period_drp .my-select__control"))).click()
        wait.until(EC.element_to_be_clickable((By.XPATH, f"//div[contains(text(), '{year}')]"))).click()
        wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".period_drp.me-1 .my-select__control"))).click()
        wait.until(EC.element_to_be_clickable((By.XPATH, f"//div[contains(text(), '{month}')]"))).click()
        time.sleep(10)
        try:
            Select(wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "select[aria-label='Choose a page size']")))).select_by_visible_text("100")
        except:
            pass
        time.sleep(10)
    except Exception as e:
        st.error(f"Filter selection failed: {e}")

# 🔗 Extract Excel file links
def extract_links_from_table(driver, wait):
    links = []

    def extract():
        try:
            table = wait.until(EC.presence_of_element_located((By.XPATH, '//table')))
            for row in table.find_elements(By.TAG_NAME, "tr"):
                for a in row.find_elements(By.TAG_NAME, "a"):
                    href = a.get_attribute("href")
                    if href and isinstance(href, str) and "PSP" in href and href.endswith((".xls", ".xlsx", ".XLS")):
                        try:
                            date_str = href.split("/")[-1].split("_")[0]
                            dt = datetime.strptime(date_str, "%d.%m.%y")
                            links.append((dt, href))
                        except:
                            continue
        except Exception as e:
            st.warning(f"Link extraction error: {e}")

    extract()
    while True:
        try:
            next_btn = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "button[aria-label='Next Page']")))
            if next_btn.is_enabled():
                driver.execute_script("arguments[0].scrollIntoView(true);", next_btn)
                next_btn.click()
                time.sleep(5)
                extract()
            else:
                break
        except:
            break

    return sorted(links, key=lambda x: x[0])

# 📦 Process and parse Excel files
def process_excel_links(excel_links):
    expected = ["Region", "NR", "WR", "SR", "ER", "NER", "Total", "Remarks"]
    data = []

    for dt, url in excel_links:
        try:
            r = requests.get(url, verify=False)
            if r.status_code == 200:
                ext = url.split(".")[-1].lower()
                engine = "openpyxl" if ext == "xlsx" else "xlrd"
                df_full = pd.read_excel(BytesIO(r.content), sheet_name="MOP_E", engine=engine, header=None)
                df = df_full.iloc[5:13, :8].copy()
                df.columns = expected
                df.insert(0, "Date", dt.strftime("%d-%m-%Y"))
                data.append(df)
        except Exception as e:
            st.warning(f"Could not process {url}: {e}")

    return pd.concat(data, ignore_index=True) if data else None

# 🖥️ Streamlit UI
def main():
    st.title("📊 Grid India PSP Extractor")

    years = [f"{y}-{str(y+1)[-2:]}" for y in range(2023, 2026)]
    selected_year = st.selectbox("Select Financial Year", years[::-1])

    months = ["ALL", "April", "May", "June", "July", "August", "September", "October", "November", "December", "January", "February", "March"]
    selected_month = st.selectbox("Select Month", months)

    if st.button("Extract Data"):
        with st.spinner("🔍 Loading and scraping data..."):
            driver = get_website_content("https://grid-india.in/en/reports/daily-psp-report")
            if not driver:
                show_selenium_log()
                return

            try:
                wait = WebDriverWait(driver, 30)
                select_filters(driver, wait, selected_year, selected_month)
                excel_links = extract_links_from_table(driver, wait)
            finally:
                driver.quit()

            if not excel_links:
                st.error("No report links found.")
                show_selenium_log()
                return

            df = process_excel_links(excel_links)
            if df is not None:
                bio = BytesIO()
                df.to_excel(bio, index=False)
                bio.seek(0)
                st.success(f"✅ Extracted {len(df)} rows.")
                st.download_button("📥 Download Excel", bio, "Grid_India_PSP_Report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            else:
                st.error("No valid Excel content.")

            show_selenium_log()

if __name__ == "__main__":
    main()

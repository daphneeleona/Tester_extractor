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
from selenium.common.exceptions import WebDriverException
from webdriver_manager.chrome import ChromeDriverManager

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# 🧰 Display selenium.log in Streamlit
def show_selenium_log():
    log_path = 'selenium.log'
    if os.path.exists(log_path):
        with open(log_path, 'r') as f:
            log_content = f.read()
        st.subheader("🔍 Selenium Log Output")
        st.code(log_content, language='bash')
    else:
        st.warning("🚫 Log file not found.")

# 🚗 Get WebDriver with error handling
def get_driver():
    options = Options()
    options.add_argument('--headless=new')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--disable-gpu')
    options.add_argument('--window-size=1920,1080')
    options.add_argument('--disable-features=NetworkService')
    options.add_argument('--disable-features=VizDisplayCompositor')
    options.add_argument('--disable-blink-features=AutomationControlled')

    try:
        service = Service(ChromeDriverManager().install(), log_path='selenium.log')
        driver = webdriver.Chrome(service=service, options=options)
        driver.implicitly_wait(10)
        return driver
    except WebDriverException as e:
        st.error("⚠️ WebDriver failed to initialize.")
        st.code(str(e), language='bash')
        return None

# 🌐 Load page with retries
def get_website_content(target_url, max_retries=3):
    for attempt in range(max_retries):
        st.info(f"🔁 Attempt {attempt + 1} to load page...")
        driver = get_driver()
        if driver is None:
            continue

        try:
            driver.get(target_url)
            time.sleep(3)
            st.success("✅ Page loaded successfully.")
            return driver
        except Exception as e:
            st.warning(f"⚠️ Failed to load URL: {e}")
            driver.quit()

    st.error("❌ All attempts failed.")
    return None

# 🕸️ Select filters
def select_filters(driver, wait, year, month):
    try:
        dropdown1 = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".period_drp .my-select__control")))
        dropdown1.click()
        wait.until(EC.element_to_be_clickable((By.XPATH, f"//div[contains(text(), '{year}')]"))).click()

        dropdown2 = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".period_drp.me-1 .my-select__control")))
        dropdown2.click()
        wait.until(EC.element_to_be_clickable((By.XPATH, f"//div[contains(text(), '{month}')]"))).click()

        time.sleep(10)
        try:
            Select(wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "select[aria-label='Choose a page size']")))).select_by_visible_text("100")
        except:
            pass
        time.sleep(10)
    except Exception as e:
        st.error(f"Failed to apply filters: {e}")

# 🧵 Extract Excel links
def extract_links_from_table(driver, wait):
    excel_links = []

    def extract():
        try:
            table = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="root"]/div/div[1]/main/div/div[3]/div/div/div[2]/table')))
            rows = table.find_elements(By.TAG_NAME, "tr")
            for row in rows:
                links = row.find_elements(By.TAG_NAME, "a")
                for link in links:
                    href = link.get_attribute("href")
                    if href and "PSP" in href and href.endswith((".xls", ".xlsx", ".XLS")):
                        try:
                            date_str = href.split("/")[-1].split("_")[0]
                            report_date = datetime.strptime(date_str, "%d.%m.%y")
                            excel_links.append((report_date, href))
                        except:
                            continue
        except Exception as e:
            st.error(f"Error locating or reading table: {e}")

    extract()
    while True:
        try:
            next_button = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "button[aria-label='Next Page']")))
            if next_button.is_enabled():
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", next_button)
                time.sleep(1)
                next_button.click()
                time.sleep(5)
                extract()
            else:
                break
        except:
            break

    return sorted(excel_links, key=lambda x: x[0])

# 📦 Process Excel downloads
def process_excel_links(excel_links):
    expected_columns = ["Region", "NR", "WR", "SR", "ER", "NER", "Total", "Remarks"]
    combined_data = []

    for report_date, url in excel_links:
        try:
            response = requests.get(url, verify=False)
            if response.status_code == 200:
                ext = url.split(".")[-1].lower()
                engine = "openpyxl" if ext == "xlsx" else "xlrd"
                df_full = pd.read_excel(BytesIO(response.content), sheet_name="MOP_E", engine=engine, header=None)
                df = df_full.iloc[5:13, :8].copy()
                df.columns = expected_columns
                df.insert(0, "Date", report_date.strftime("%d-%m-%Y"))
                combined_data.append(df)
        except Exception as e:
            st.warning(f"Failed to process {url}: {e}")

    return pd.concat(combined_data, ignore_index=True) if combined_data else None

# 🧪 Streamlit app entry point
def main():
    st.title("📊 Grid India PSP Report Extractor")

    years = [f"{y}-{str(y+1)[-2:]}" for y in range(2023, 2026)]
    selected_year = st.selectbox("Select Financial Year", years[::-1])

    months = ["ALL", "April", "May", "June", "July", "August", "September", "October", "November", "December", "January", "February", "March"]
    selected_month = st.selectbox("Select Month", months)

    if st.button("Extract Data"):
        with st.spinner("🧪 Scraping data... Please wait"):
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
                st.error("No data extracted.")
                show_selenium_log()
                return

            final_df = process_excel_links(excel_links)

            if final_df is not None:
                output = BytesIO()
                final_df.to_excel(output, index=False)
                output.seek(0)

                st.success(f"✅ Data extraction complete! Extracted {len(final_df)} rows.")
                st.download_button(
                    label="📥 Download Excel",
                    data=output,
                    file_name="Grid_India_PSP_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error("No valid Excel data found.")

            show_selenium_log()

if __name__ == "__main__":
    main()

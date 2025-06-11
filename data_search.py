from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import time
import openpyxl

# === CONFIG ===
URL = "https://dir.indiamart.com/search.mp?ss=insurance+agents+in+gujarat&prdsrc=1&v=4&mcatid=172089&catid=328&cq=ahmedabad&tags=res:RC2|ktp:N0|stype:attr=1|mtp:G|wc:4|cq:ahmedabad|qr_nm:gl-gd|cs:17549|com-cf:nl|ptrs:na|mc:169407|cat:328|qry_typ:S|lang:en|rtn:6-0-0-1-2-0-1|tyr:1|qrd:250611|mrd:250611|prdt:250611|msf:ls|pfen:0"
OUTPUT_FILE = "Insurance_Agents_Gujarat.xlsx"
MAX_ENTRIES = 30
# =============

# Step 1: Excel setup
wb = openpyxl.Workbook()
sheet = wb.active
sheet.title = "Insurance Agents"
sheet.append(["Name", "Category", "Address", "Mobile No.", "City"])

# Step 2: Set up Selenium
options = Options()
options.add_argument("--headless")
options.add_argument("--log-level=3")
options.add_experimental_option("excludeSwitches", ["enable-logging"])
driver = webdriver.Chrome(options=options)

try:
    # Step 3: Load the custom URL
    driver.get(URL)
    time.sleep(6)

    # Scroll to bottom to load more listings
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(5)

    # Step 4: Parse business cards
    cards = driver.find_elements(By.CLASS_NAME, "prd-name")
    count = 0

    for card in cards:
        if count >= MAX_ENTRIES:
            break

        try:
            name = card.text.strip()
            container = card.find_element(By.XPATH, "..").find_element(By.XPATH, "..")

            try:
                mobile = container.find_element(By.CLASS_NAME, "mb-view-phone").text.strip()
            except:
                mobile = "Not Available"

            try:
                address = container.find_element(By.CLASS_NAME, "cmpny-location").text.strip()
            except:
                address = "Not Available"

            try:
                category = container.find_element(By.CLASS_NAME, "prd-cat").text.strip()
            except:
                category = "Insurance Agent"

            city = address.split(",")[-1].strip() if "," in address else "Not Available"

            sheet.append([name, category, address, mobile, city])
            count += 1

        except Exception as e:
            print("⚠️ Error parsing a card:", e)

finally:
    wb.save(OUTPUT_FILE)
    driver.quit()
    print(f"{count} entries saved to {OUTPUT_FILE}")

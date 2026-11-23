from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import pandas as pd

# Setup
driver = webdriver.Chrome()
driver.maximize_window()
wait = WebDriverWait(driver, 20)

accessories_data = []
trims_data = []

# --------------------------------------------------
# STEP 1: Open URL
# --------------------------------------------------
driver.get("https://sandbox-admin.partsgorilla.com/login")

# --------------------------------------------------
# STEP 2: Login
# --------------------------------------------------
wait.until(EC.presence_of_element_located((
    By.XPATH,
    "/html/body/main/div/div/div/ui-view/div[1]/div/div/div/div[2]/div/form/div[1]/input"
))).send_keys("izmosupport")

driver.find_element(
    By.XPATH,
    "/html/body/main/div/div/div/ui-view/div[1]/div/div/div/div[2]/div/form/div[2]/input"
).send_keys("adm!n@supp@rt")

wait.until(EC.element_to_be_clickable((
    By.XPATH,
    "/html/body/main/div/div/div/ui-view/div[1]/div/div/div/div[2]/div/form/div[3]/input"
))).click()

wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))

# --------------------------------------------------
# STEP 3: Navigate → Filter → Select Model
# --------------------------------------------------
wait.until(EC.element_to_be_clickable((
    By.XPATH, '//*[@id="sideNavMenu"]/li[3]/a/span[2]'
))).click()

wait.until(EC.element_to_be_clickable((
    By.XPATH, '//*[@id="modeldt_column_filter"]'
))).click()

# Model filter
wait.until(EC.element_to_be_clickable((
    By.XPATH, '//*[@id="filter-items"]/div[3]/span/span/span'
))).click()

model_input = wait.until(EC.element_to_be_clickable((
    By.XPATH, '//*[@id="modeldt_Model"]/div/div/input'
)))
model_input.send_keys(Keys.CONTROL + "a")
model_input.send_keys(Keys.DELETE)
model_input.send_keys("accord")

wait.until(EC.element_to_be_clickable((
    By.XPATH, '//*[@id="modeldt_Model"]/div/div/button/font/font'
))).click()

# Year filter
wait.until(EC.element_to_be_clickable((
    By.XPATH, '//*[@id="filter-items"]/div[5]/span/span/span'
))).click()

year_input = wait.until(EC.element_to_be_clickable((
    By.XPATH, '//*[@id="modeldt_Year"]/div/div/input'
)))
year_input.send_keys("2025")

wait.until(EC.element_to_be_clickable((
    By.XPATH, '//*[@id="modeldt_Year"]/div/div/button'
))).click()

# Wait for table reload
wait.until(EC.invisibility_of_element_located((
    By.XPATH, '//*[@id="modeldt_processing"]'
)))

# Click correct model
rows = driver.find_elements(By.XPATH, '//*[@id="modeldt"]/tbody/tr')

for row in rows:
    if "accord" in row.find_element(By.XPATH, "./td[3]").text.lower():
        row.find_element(By.XPATH, "./td[3]/a").click()
        break

# --------------------------------------------------
# STEP 4: GET MODEL DETAILS
# --------------------------------------------------
make = wait.until(EC.visibility_of_element_located((
    By.XPATH, '//*[@id="nav-7-1"]/div/div[2]/div[2]'
))).text

model = wait.until(EC.visibility_of_element_located((
    By.XPATH, '//*[@id="nav-7-1"]/div/div[3]/div[2]'
))).text

year = wait.until(EC.visibility_of_element_located((
    By.XPATH, '//*[@id="nav-7-1"]/div/div[5]/div[2]'
))).text

# ==================================================
# 🔥 ACCESSORIES EXTRACTION (Sheet 1)
# ==================================================
wait.until(EC.element_to_be_clickable((
    By.XPATH, '/html/body/main/div[2]/div[2]/div/ui-view/div/div[2]/div[2]/header/div/div/ul/li[4]/a'
))).click()

while True:
    wait.until(lambda d: len(d.find_elements(By.XPATH, '//*[@id="modelAccessories"]/tbody/tr')) > 0)
    rows = driver.find_elements(By.XPATH, '//*[@id="modelAccessories"]/tbody/tr')

    for row in rows:
        accessories_data.append({
            "Make": make,
            "Model": model,
            "Year": year,
            "PartName": row.find_element(By.XPATH, "./td[2]").text,
            "Category": row.find_element(By.XPATH, "./td[3]").text,
            "Status": row.find_element(By.XPATH, "./td[5]/span").text
        })

    try:
        next_btn = driver.find_element(By.XPATH, '//*[@id="modelAccessories_next"]')
        if "disabled" in next_btn.get_attribute("class"):
            break
        next_btn.click()
        wait.until(EC.staleness_of(rows[0]))
    except:
        break

# ==================================================
# 🔥 TRIMS EXTRACTION (Sheet 2)
# ==================================================
wait.until(EC.element_to_be_clickable((
    By.XPATH, '/html/body/main/div[2]/div[2]/div/ui-view/div/div[2]/div[2]/header/div/div/ul/li[3]/a'
))).click()

print("✅ Trims tab opened")

while True:
    wait.until(lambda d: len(d.find_elements(By.XPATH, '//*[@id="trimDt"]/tbody/tr')) > 0)
    rows = driver.find_elements(By.XPATH, '//*[@id="trimDt"]/tbody/tr')

    for row in rows:
        trim_name = row.find_element(By.XPATH, "./td[2]/a").text
        trim_status = row.find_element(By.XPATH, "./td[6]/span").text

        trims_data.append({
            "Make": make,
            "Model": model,
            "Year": year,
            "Trim": trim_name,
            "Trim_Status": trim_status
        })

    # Pagination (if exists)
    try:
        next_btn = driver.find_element(By.XPATH, '//*[@id="trimDt_next"]')
        if "disabled" in next_btn.get_attribute("class"):
            break
        next_btn.click()
        wait.until(EC.staleness_of(rows[0]))
    except:
        break

# --------------------------------------------------
# SAVE TO EXCEL (MULTIPLE SHEETS)
# --------------------------------------------------
with pd.ExcelWriter("final_output.xlsx") as writer:
    pd.DataFrame(accessories_data).to_excel(writer, sheet_name="Accessories", index=False)
    pd.DataFrame(trims_data).to_excel(writer, sheet_name="Trims", index=False)

print("✅ Excel created with 2 sheets (Accessories + Trims)")

input("Press Enter to close...")
driver.quit()
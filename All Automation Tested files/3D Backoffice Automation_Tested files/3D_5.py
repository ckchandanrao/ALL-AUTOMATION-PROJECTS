from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import pandas as pd
import time

# Setup
driver = webdriver.Chrome()
driver.maximize_window()
wait = WebDriverWait(driver, 20)

data = []

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

print("✅ Logged in")

wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))

# --------------------------------------------------
# STEP 3: Navigate to Models
# --------------------------------------------------
wait.until(EC.element_to_be_clickable((
    By.XPATH, '//*[@id="sideNavMenu"]/li[3]/a/span[2]'
))).click()

# --------------------------------------------------
# STEP 4: Open Filter
# --------------------------------------------------
wait.until(EC.element_to_be_clickable((
    By.XPATH, '//*[@id="modeldt_column_filter"]'
))).click()

# --------------------------------------------------
# STEP 5: MODEL FILTER
# --------------------------------------------------
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

# --------------------------------------------------
# STEP 6: YEAR FILTER
# --------------------------------------------------
wait.until(EC.element_to_be_clickable((
    By.XPATH, '//*[@id="filter-items"]/div[5]/span/span/span'
))).click()

year_input = wait.until(EC.element_to_be_clickable((
    By.XPATH, '//*[@id="modeldt_Year"]/div/div/input'
)))

year_input.send_keys(Keys.CONTROL + "a")
year_input.send_keys(Keys.DELETE)
year_input.send_keys("2025")

wait.until(EC.element_to_be_clickable((
    By.XPATH, '//*[@id="modeldt_Year"]/div/div/button'
))).click()

# Wait for table reload
wait.until(EC.invisibility_of_element_located((
    By.XPATH, '//*[@id="modeldt_processing"]'
)))

wait.until(lambda d: len(d.find_elements(By.XPATH, '//*[@id="modeldt"]/tbody/tr')) > 0)

# --------------------------------------------------
# STEP 7: CLICK CORRECT MODEL
# --------------------------------------------------
rows = driver.find_elements(By.XPATH, '//*[@id="modeldt"]/tbody/tr')

for row in rows:
    model_text = row.find_element(By.XPATH, "./td[3]").text.lower()

    if "accord" in model_text:
        row.find_element(By.XPATH, "./td[3]/a").click()
        break

# --------------------------------------------------
# STEP 8: FETCH MODEL DETAILS
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

status = wait.until(EC.visibility_of_element_located((
    By.XPATH, '//*[@id="nav-7-1"]/div/div[6]/div[2]'
))).text

# --------------------------------------------------
# STEP 9: ACCESSORIES TAB
# --------------------------------------------------
wait.until(EC.element_to_be_clickable((
    By.XPATH, '/html/body/main/div[2]/div[2]/div/ui-view/div/div[2]/div[2]/header/div/div/ul/li[4]/a'
))).click()

# --------------------------------------------------
# STEP 10: PAGINATION LOOP
# --------------------------------------------------
while True:

    # Wait for rows
    wait.until(lambda d: len(d.find_elements(By.XPATH, '//*[@id="modelAccessories"]/tbody/tr')) > 0)

    rows = driver.find_elements(By.XPATH, '//*[@id="modelAccessories"]/tbody/tr')

    print(f"Processing {len(rows)} rows...")

    for row in rows:
        try:
            part_name = row.find_element(By.XPATH, "./td[2]").text
            category = row.find_element(By.XPATH, "./td[3]").text
            acc_status = row.find_element(By.XPATH, "./td[5]/span").text

            data.append({
                "Make": make,
                "Model": model,
                "Year": year,
                "Status": status,
                "PartName": part_name,
                "Category": category,
                "Acc_Status": acc_status
            })

        except:
            pass

    # Try clicking NEXT page
    try:
        next_btn = driver.find_element(By.XPATH, '//*[@id="modelAccessories_next"]')

        if "disabled" in next_btn.get_attribute("class"):
            print("✅ No more pages")
            break

        next_btn.click()

        # Wait for page change
        wait.until(EC.staleness_of(rows[0]))

    except:
        break

# --------------------------------------------------
# STEP 11: SAVE TO EXCEL
# --------------------------------------------------
df = pd.DataFrame(data)
df.to_excel("final_output_all_pages.xlsx", index=False)

print("✅ Data saved to Excel")

# --------------------------------------------------
# DONE
# --------------------------------------------------
input("Press Enter to close browser...")
driver.quit()
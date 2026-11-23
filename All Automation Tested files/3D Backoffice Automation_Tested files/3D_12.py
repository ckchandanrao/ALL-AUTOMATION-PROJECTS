from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time

# ================= INPUT =================
INPUT_MODEL = input("Enter Model: ").strip().lower()
INPUT_YEAR = input("Enter Year: ").strip()

driver = webdriver.Chrome()
driver.maximize_window()
wait = WebDriverWait(driver, 25)

accessories_data = []

# ---------------- LOGIN ----------------
driver.get("https://sandbox-admin.partsgorilla.com/login")

wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='text']"))).send_keys("izmosupport")
driver.find_element(By.XPATH, "//input[@type='password']").send_keys("adm!n@supp@rt")
driver.find_element(By.XPATH, "//input[@type='submit']").click()

# ---------------- NAVIGATION ----------------
wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="sideNavMenu"]/li[3]/a/span[2]'))).click()
wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="modeldt_column_filter"]'))).click()

# ---------------- FILTER ----------------
wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="filter-items"]/div[3]/span/span/span'))).click()
driver.find_element(By.XPATH, '//*[@id="modeldt_Model"]/div/div/input').send_keys(INPUT_MODEL)
driver.find_element(By.XPATH, '//*[@id="modeldt_Model"]/div/div/button').click()

wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="filter-items"]/div[5]/span/span/span'))).click()
driver.find_element(By.XPATH, '//*[@id="modeldt_Year"]/div/div/input').send_keys(INPUT_YEAR)
driver.find_element(By.XPATH, '//*[@id="modeldt_Year"]/div/div/button').click()

wait.until(EC.invisibility_of_element_located((By.XPATH, '//*[@id="modeldt_processing"]')))

# ---------------- COUNTRY ----------------
country = driver.find_element(By.XPATH, '//*[@id="modeldt"]/tbody/tr[1]/td[2]').text.strip()

# ---------------- CLICK MODEL ----------------
rows = driver.find_elements(By.XPATH, '//*[@id="modeldt"]/tbody/tr')
for row in rows:
    try:
        name = row.find_element(By.XPATH, "./td[3]").text.lower()
        if name == INPUT_MODEL:
            row.find_element(By.XPATH, "./td[3]/a").click()
            break
    except:
        continue

# ---------------- MODEL DETAILS ----------------
make = wait.until(EC.visibility_of_element_located((By.XPATH,'//*[@id="nav-7-1"]/div/div[2]/div[2]'))).text
model = wait.until(EC.visibility_of_element_located((By.XPATH,'//*[@id="nav-7-1"]/div/div[3]/div[2]'))).text
year = wait.until(EC.visibility_of_element_located((By.XPATH,'//*[@id="nav-7-1"]/div/div[5]/div[2]'))).text

# ==================================================
# MODEL ACCESSORIES (NA → ALL)
# ==================================================
acc_tab = wait.until(EC.element_to_be_clickable((
    By.XPATH,
    '/html/body/main/div[2]/div[2]/div/ui-view/div/div[2]/div[2]/header/div/div/ul/li[4]/a'
)))
driver.execute_script("arguments[0].click();", acc_tab)

time.sleep(2)

rows = driver.find_elements(By.XPATH, '//*[@id="modelAccessories"]/tbody/tr')

valid = []
for r in rows:
    cols = r.find_elements(By.TAG_NAME, "td")
    if len(cols) >= 5:
        valid.append(cols)

if len(valid) == 0:
    accessories_data.append({
        "Country": country,
        "Make": make,
        "Model": model,
        "Year": year,
        "PartName": "NA",
        "Category": "NA",
        "Status": "NA",
        "Trim": "ALL"
    })
else:
    for cols in valid:
        accessories_data.append({
            "Country": country,
            "Make": make,
            "Model": model,
            "Year": year,
            "PartName": cols[1].text.strip(),
            "Category": cols[2].text.strip(),
            "Status": cols[4].text.strip(),
            "Trim": "ALL"
        })

print("✅ Model accessories done")

# ==================================================
# GO TO TRIMS
# ==================================================
trim_tab = wait.until(EC.element_to_be_clickable((
    By.XPATH,
    '/html/body/main/div[2]/div[2]/div/ui-view/div/div[2]/div[2]/header/div/div/ul/li[3]/a'
)))
driver.execute_script("arguments[0].click();", trim_tab)

time.sleep(2)

# ==================================================
# TRIMS LOOP WITH PAGINATION
# ==================================================
while True:
    trim_rows = driver.find_elements(By.XPATH, '//*[@id="trimDt"]/tbody/tr')

    if len(trim_rows) == 0:
        break

    i = 0
    while i < len(trim_rows):
        try:
            trim_rows = driver.find_elements(By.XPATH, '//*[@id="trimDt"]/tbody/tr')

            if i >= len(trim_rows):
                break

            row = trim_rows[i]
            trim_name = row.find_element(By.XPATH, "./td[2]/a").text.strip()

            row.find_element(By.XPATH, "./td[2]/a").click()

            # ACCESSORIES TAB
            acc_tab = wait.until(EC.element_to_be_clickable((
                By.XPATH,
                '/html/body/main/div[2]/div[2]/div/ui-view/div/div[2]/div[2]/div[1]/header/div/div/ul/li[3]/a'
            )))
            driver.execute_script("arguments[0].click();", acc_tab)

            time.sleep(2)

            rows = driver.find_elements(By.XPATH, '//*[@id="trimAccessoriesDt"]/tbody/tr')

            valid = []
            for r in rows:
                cols = r.find_elements(By.TAG_NAME, "td")
                if len(cols) >= 5:
                    valid.append(cols)

            if len(valid) == 0:
                print(f"⚠️ No accessories for {trim_name}")
            else:
                for cols in valid:
                    accessories_data.append({
                        "Country": country,
                        "Make": make,
                        "Model": model,
                        "Year": year,
                        "PartName": cols[1].text.strip(),
                        "Category": cols[2].text.strip(),
                        "Status": cols[4].text.strip(),
                        "Trim": trim_name
                    })

            # BACK
            back_btn = wait.until(EC.element_to_be_clickable((
                By.XPATH,
                '/html/body/main/div[2]/div[2]/breadcrumbs/div/ul/li[2]/span/a'
            )))
            driver.execute_script("arguments[0].click();", back_btn)

            wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="trimDt"]/tbody/tr')))

            i += 1

        except Exception as e:
            print("Trim error:", e)
            i += 1
            continue

    # -------- PAGINATION --------
    try:
        next_btn = driver.find_element(By.XPATH, '//*[@id="trimDt_next"]')

        if "disabled" not in next_btn.get_attribute("class"):
            print("➡️ Moving to next trim page...")

            old_first_row = driver.find_elements(By.XPATH, '//*[@id="trimDt"]/tbody/tr')[0]

            driver.execute_script("arguments[0].click();", next_btn)

            wait.until(EC.staleness_of(old_first_row))
            wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="trimDt"]/tbody/tr')))

        else:
            print("✅ Last trim page reached")
            break

    except Exception as e:
        print("Pagination error:", e)
        break

# ==================================================
# SAVE
# ==================================================
df = pd.DataFrame(accessories_data)
df.to_excel("final_output.xlsx", index=False)

print("✅ FINAL FILE GENERATED")
driver.quit()
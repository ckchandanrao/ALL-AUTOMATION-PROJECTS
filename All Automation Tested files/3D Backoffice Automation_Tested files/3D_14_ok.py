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
trim_data = []
model_color_set = set()

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
        if row.find_element(By.XPATH, "./td[3]").text.lower() == INPUT_MODEL:
            row.find_element(By.XPATH, "./td[3]/a").click()
            break
    except:
        continue

# ---------------- MODEL DETAILS ----------------
make = wait.until(EC.visibility_of_element_located((By.XPATH,'//*[@id="nav-7-1"]/div/div[2]/div[2]'))).text
model = wait.until(EC.visibility_of_element_located((By.XPATH,'//*[@id="nav-7-1"]/div/div[3]/div[2]'))).text
year = wait.until(EC.visibility_of_element_located((By.XPATH,'//*[@id="nav-7-1"]/div/div[5]/div[2]'))).text

# ==================================================
# MODEL COLORS
# ==================================================
driver.execute_script("arguments[0].click();", wait.until(EC.element_to_be_clickable((
    By.XPATH,'/html/body/main/div[2]/div[2]/div/ui-view/div/div[2]/div[2]/header/div/div/ul/li[2]/a'))))

time.sleep(2)

for row in driver.find_elements(By.XPATH,'//*[@id="modelColorDt"]/tbody/tr'):
    cols = row.find_elements(By.TAG_NAME,"td")
    if cols:
        model_color_set.add(cols[0].text.strip())

# ==================================================
# MODEL ACCESSORIES
# ==================================================
driver.execute_script("arguments[0].click();", wait.until(EC.element_to_be_clickable((
    By.XPATH,'/html/body/main/div[2]/div[2]/div/ui-view/div/div[2]/div[2]/header/div/div/ul/li[4]/a'))))

time.sleep(2)

rows = driver.find_elements(By.XPATH,'//*[@id="modelAccessories"]/tbody/tr')

valid = []
for r in rows:
    cols = r.find_elements(By.TAG_NAME,"td")
    if len(cols) >= 5:
        valid.append(cols)

if not valid:
    accessories_data.append({
        "Country":country,"Make":make,"Model":model,"Year":year,
        "PartName":"NA","Category":"NA","Status":"NA","Trim":"ALL"
    })
else:
    for cols in valid:
        accessories_data.append({
            "Country":country,"Make":make,"Model":model,"Year":year,
            "PartName":cols[1].text,"Category":cols[2].text,
            "Status":cols[4].text,"Trim":"ALL"
        })

print("✅ Model accessories done")

# ==================================================
# TRIMS LOOP (COLORS + ACCESSORIES)
# ==================================================
driver.execute_script("arguments[0].click();", wait.until(EC.element_to_be_clickable((
    By.XPATH,'/html/body/main/div[2]/div[2]/div/ui-view/div/div[2]/div[2]/header/div/div/ul/li[3]/a'))))

time.sleep(2)

while True:
    trim_rows = driver.find_elements(By.XPATH,'//*[@id="trimDt"]/tbody/tr')
    if not trim_rows:
        break

    i = 0
    while i < len(trim_rows):
        try:
            trim_rows = driver.find_elements(By.XPATH,'//*[@id="trimDt"]/tbody/tr')
            row = trim_rows[i]

            trim_name = row.find_element(By.XPATH,"./td[2]/a").text
            trim_status = row.find_element(By.XPATH,"./td[6]/span").text

            row.find_element(By.XPATH,"./td[2]/a").click()

            # -------- TRIM COLORS --------
            driver.execute_script("arguments[0].click();", wait.until(EC.element_to_be_clickable((
                By.XPATH,'/html/body/main/div[2]/div[2]/div/ui-view/div/div[2]/div[2]/div[1]/header/div/div/ul/li[2]/a'))))

            time.sleep(1)

            trim_colors = []
            for r in driver.find_elements(By.XPATH,'//*[@id="trimColorDt"]/tbody/tr'):
                cols = r.find_elements(By.TAG_NAME,"td")
                if cols:
                    trim_colors.append(cols[0].text.strip())

            combined = sorted(set(trim_colors).union(model_color_set))

            # -------- TRIM ACCESSORIES --------
            driver.execute_script("arguments[0].click();", wait.until(EC.element_to_be_clickable((
                By.XPATH,'/html/body/main/div[2]/div[2]/div/ui-view/div/div[2]/div[2]/div[1]/header/div/div/ul/li[3]/a'))))

            time.sleep(1)

            rows = driver.find_elements(By.XPATH,'//*[@id="trimAccessoriesDt"]/tbody/tr')

            valid = []
            for r in rows:
                cols = r.find_elements(By.TAG_NAME,"td")
                if len(cols) >= 5:
                    valid.append(cols)

            if valid:
                for cols in valid:
                    accessories_data.append({
                        "Country":country,"Make":make,"Model":model,"Year":year,
                        "PartName":cols[1].text,
                        "Category":cols[2].text,
                        "Status":cols[4].text,
                        "Trim":trim_name
                    })

            trim_data.append({
                "Country":country,"Make":make,"Model":model,"Year":year,
                "Trim":trim_name,"Trim_Status":trim_status,
                "Colors":" | ".join(combined)
            })

            # BACK
            driver.execute_script("arguments[0].click();", wait.until(EC.element_to_be_clickable((
                By.XPATH,'/html/body/main/div[2]/div[2]/breadcrumbs/div/ul/li[2]/span/a'))))

            wait.until(EC.presence_of_element_located((By.XPATH,'//*[@id="trimDt"]/tbody/tr')))

            i += 1

        except Exception as e:
            print("Trim error:", e)
            i += 1

    # PAGINATION
    try:
        next_btn = driver.find_element(By.XPATH,'//*[@id="trimDt_next"]')
        if "disabled" not in next_btn.get_attribute("class"):
            old = driver.find_elements(By.XPATH,'//*[@id="trimDt"]/tbody/tr')[0]
            driver.execute_script("arguments[0].click();", next_btn)
            wait.until(EC.staleness_of(old))
        else:
            break
    except:
        break

# ==================================================
# SAVE
# ==================================================
with pd.ExcelWriter("final_output.xlsx") as writer:
    pd.DataFrame(accessories_data).to_excel(writer,"Accessories",index=False)
    pd.DataFrame(trim_data).to_excel(writer,"Trims",index=False)

print("✅ FINAL OUTPUT READY")
driver.quit()
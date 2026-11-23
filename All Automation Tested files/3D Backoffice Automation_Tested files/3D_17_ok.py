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

# 🔥 MERGE STRUCTURE (for accessories)
accessory_map = {}
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
for row in driver.find_elements(By.XPATH, '//*[@id="modeldt"]/tbody/tr'):
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
color_tab_model = wait.until(EC.element_to_be_clickable((
    By.XPATH,
    '/html/body/main/div[2]/div[2]/div/ui-view/div/div[2]/div[2]/header/div/div/ul/li[2]/a'
)))
driver.execute_script("arguments[0].click();", color_tab_model)

time.sleep(2)

rows = driver.find_elements(By.XPATH, '//*[@id="modelColorDt"]/tbody/tr')
for r in rows:
    cols = r.find_elements(By.TAG_NAME, "td")
    if cols:
        model_color_set.add(cols[0].text.strip())

print("✅ Model colors:", model_color_set)

# ==================================================
# MODEL ACCESSORIES
# ==================================================
acc_tab_model = wait.until(EC.element_to_be_clickable((
    By.XPATH,
    '/html/body/main/div[2]/div[2]/div/ui-view/div/div[2]/div[2]/header/div/div/ul/li[4]/a'
)))
driver.execute_script("arguments[0].click();", acc_tab_model)

time.sleep(2)

rows = driver.find_elements(By.XPATH,'//*[@id="modelAccessories"]/tbody/tr')
valid = [r.find_elements(By.TAG_NAME,"td") for r in rows if len(r.find_elements(By.TAG_NAME,"td")) >= 5]

if not valid:
    accessory_map["NA"] = {"Category":"NA","Status":"NA","Trims":{"ALL"}}
else:
    for cols in valid:
        part = cols[1].text.strip()
        accessory_map[part] = {
            "Category": cols[2].text.strip(),
            "Status": cols[4].text.strip(),
            "Trims": {"ALL"}
        }

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
# TRIMS LOOP (COLORS + ACCESSORIES)
# ==================================================
while True:
    trim_rows = driver.find_elements(By.XPATH,'//*[@id="trimDt"]/tbody/tr')
    if not trim_rows:
        break

    i = 0
    while i < len(trim_rows):
        try:
            trim_rows = driver.find_elements(By.XPATH,'//*[@id="trimDt"]/tbody/tr')
            row = trim_rows[i]

            trim_name = row.find_element(By.XPATH,"./td[2]/a").text.strip()
            trim_status = row.find_element(By.XPATH,"./td[6]/span").text.strip()

            row.find_element(By.XPATH,"./td[2]/a").click()

            # -------- TRIM COLORS --------
            color_tab_trim = wait.until(EC.element_to_be_clickable((
                By.XPATH,
                '/html/body/main/div[2]/div[2]/div/ui-view/div/div[2]/div[2]/div[1]/header/div/div/ul/li[2]/a'
            )))
            driver.execute_script("arguments[0].click();", color_tab_trim)

            time.sleep(1)

            trim_colors = []
            rows = driver.find_elements(By.XPATH, '//*[@id="trimColorDt"]/tbody/tr')
            for r in rows:
                cols = r.find_elements(By.TAG_NAME, "td")
                if cols:
                    trim_colors.append(cols[0].text.strip())

            combined_colors = sorted(set(trim_colors).union(model_color_set))

            # -------- TRIM ACCESSORIES --------
            acc_tab_trim = wait.until(EC.element_to_be_clickable((
                By.XPATH,
                '/html/body/main/div[2]/div[2]/div/ui-view/div/div[2]/div[2]/div[1]/header/div/div/ul/li[3]/a'
            )))
            driver.execute_script("arguments[0].click();", acc_tab_trim)

            time.sleep(1)

            rows = driver.find_elements(By.XPATH,'//*[@id="trimAccessoriesDt"]/tbody/tr')
            for r in rows:
                cols = r.find_elements(By.TAG_NAME,"td")
                if len(cols) < 5:
                    continue

                part = cols[1].text.strip()

                if part in accessory_map:
                    accessory_map[part]["Trims"].add(trim_name)
                else:
                    accessory_map[part] = {
                        "Category": cols[2].text.strip(),
                        "Status": cols[4].text.strip(),
                        "Trims": {trim_name}
                    }

            # -------- STORE TRIM DATA --------
            trim_data.append({
                "Country": country,
                "Year": year,
                "Make": make,
                "Model": model,
                "Trim": trim_name,
                "Trim_Status": trim_status,
                "Color_Count": len(combined_colors),
                "Colors": ",".join(combined_colors)
            })

            # -------- BACK --------
            back_btn = wait.until(EC.element_to_be_clickable((
                By.XPATH,
                '/html/body/main/div[2]/div[2]/breadcrumbs/div/ul/li[2]/span/a'
            )))
            driver.execute_script("arguments[0].click();", back_btn)

            wait.until(EC.presence_of_element_located((By.XPATH,'//*[@id="trimDt"]/tbody/tr')))

            i += 1

        except Exception as e:
            print("Trim error:", e)
            i += 1

    # -------- PAGINATION --------
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
# BUILD TRIM → ACCESSORY COUNT MAP
# ==================================================
trim_count_map = {}

for part, val in accessory_map.items():

    if part=="NA":
        continue

    for t in val["Trims"]:
        if t not in trim_count_map:
            trim_count_map[t] = set()
        trim_count_map[t].add(part)

# Convert sets to counts
for t in trim_count_map:
    trim_count_map[t] = len(trim_count_map[t])

# ==================================================
# FINAL ACCESSORY OUTPUT (UPDATED)
# ==================================================
final_accessories = []

for part, val in accessory_map.items():

    trims_list = val["Trims"]
    trim_string = ",".join(sorted(trims_list))

    # ✅ Correct count logic (based on trim group)
    if "ALL" in trims_list:
        acc_count = trim_count_map.get("ALL", 0)
    else:
        acc_count = max([trim_count_map.get(t, 0) for t in trims_list])

    final_accessories.append({
        "Country": country,
        "Make": make,
        "Model": model,
        "Year": year,
        "PartName": part,
        "Category": val["Category"],
        "Status": val["Status"],
        "Trim": trim_string,
        "Accessory_Count": acc_count   # ✅ NEW COLUMN
    })

# ==================================================
# BUILD TRIM → ACCESSORY COUNT MAP (FROM ACCESSORIES)
# ==================================================
trim_accessory_count_map = {}

for row in final_accessories:
    trims = row["Trim"].split(",")

    for t in trims:
        t = t.strip()

        if t == "ALL":
            # Apply to all trims
            for tr in trim_data:
                trim_accessory_count_map[tr["Trim"]] = row["Accessory_Count"]
        else:
            trim_accessory_count_map[t] = row["Accessory_Count"]

# ==================================================
# MAIN SHEET CREATION
# ==================================================
main_data = []

for tr in trim_data:
    trim_name = tr["Trim"]

    main_data.append({
        "Country": tr["Country"],
        "Year": tr["Year"],
        "Make": tr["Make"],
        "Model": tr["Model"],
        "Trim": trim_name,
        "Accessories_Count": trim_accessory_count_map.get(trim_name, 0)
    })

# ==================================================
# SAVE (UPDATED WITH MAIN SHEET)
# ==================================================
with pd.ExcelWriter("final_output.xlsx") as writer:
    pd.DataFrame(final_accessories).to_excel(writer, "Accessories", index=False)
    pd.DataFrame(trim_data).to_excel(writer, "Trims", index=False)
    pd.DataFrame(main_data).to_excel(writer, "Main", index=False)

print("✅ FINAL OUTPUT WITH MAIN SHEET READY")
driver.quit()

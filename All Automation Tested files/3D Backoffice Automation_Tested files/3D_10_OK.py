from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import pandas as pd

# ================= USER INPUT =================
INPUT_MODEL = input("Enter Model (e.g., accord): ").strip().lower()
INPUT_YEAR = input("Enter Year (e.g., 2025): ").strip()

driver = webdriver.Chrome()
driver.maximize_window()
wait = WebDriverWait(driver, 20)

accessories_data = []
trim_color_map = {}

# ---------------- LOGIN ----------------
driver.get("https://sandbox-admin.partsgorilla.com/login")

wait.until(EC.presence_of_element_located((
    By.XPATH,
    "/html/body/main/div/div/div/ui-view/div[1]/div/div/div/div[2]/div/form/div[1]/input"
))).send_keys("izmosupport")

driver.find_element(By.XPATH,
"/html/body/main/div/div/div/ui-view/div[1]/div/div/div/div[2]/div/form/div[2]/input"
).send_keys("adm!n@supp@rt")

wait.until(EC.element_to_be_clickable((
    By.XPATH,
    "/html/body/main/div/div/div/ui-view/div[1]/div/div/div/div[2]/div/form/div[3]/input"
))).click()

# ---------------- NAVIGATION ----------------
wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="sideNavMenu"]/li[3]/a/span[2]'))).click()
wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="modeldt_column_filter"]'))).click()

# ---------------- MODEL FILTER ----------------
wait.until(EC.element_to_be_clickable((
    By.XPATH, '//*[@id="filter-items"]/div[3]/span/span/span'
))).click()

model_input = wait.until(EC.element_to_be_clickable((
    By.XPATH, '//*[@id="modeldt_Model"]/div/div/input'
)))
model_input.clear()
model_input.send_keys(INPUT_MODEL)

wait.until(EC.element_to_be_clickable((
    By.XPATH, '//*[@id="modeldt_Model"]/div/div/button/font/font'
))).click()

# ---------------- YEAR FILTER ----------------
wait.until(EC.element_to_be_clickable((
    By.XPATH, '//*[@id="filter-items"]/div[5]/span/span/span'
))).click()

year_input = wait.until(EC.element_to_be_clickable((
    By.XPATH, '//*[@id="modeldt_Year"]/div/div/input'
)))
year_input.clear()
year_input.send_keys(INPUT_YEAR)

wait.until(EC.element_to_be_clickable((
    By.XPATH, '//*[@id="modeldt_Year"]/div/div/button'
))).click()

wait.until(EC.invisibility_of_element_located((
    By.XPATH, '//*[@id="modeldt_processing"]'
)))

# ---------------- CLICK CORRECT MODEL (FIXED) ----------------
rows = driver.find_elements(By.XPATH, '//*[@id="modeldt"]/tbody/tr')

found = False

for row in rows:
    model_name = row.find_element(By.XPATH, "./td[3]").text.strip().lower()

    if model_name == INPUT_MODEL:
        row.find_element(By.XPATH, "./td[3]/a").click()
        print(f"✅ Exact match found: {model_name}")
        found = True
        break

if not found:
    raise Exception(f"❌ Exact model '{INPUT_MODEL}' not found")

# ---------------- MODEL DETAILS ----------------
make = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="nav-7-1"]/div/div[2]/div[2]'))).text
model = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="nav-7-1"]/div/div[3]/div[2]'))).text
year = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="nav-7-1"]/div/div[5]/div[2]'))).text

# ==================================================
# MODEL COLORS
# ==================================================
wait.until(EC.element_to_be_clickable((
    By.XPATH, '/html/body/main/div[2]/div[2]/div/ui-view/div/div[2]/div[2]/header/div/div/ul/li[2]/a'
))).click()

wait.until(lambda d: len(d.find_elements(By.XPATH, '//*[@id="modelColorDt"]/tbody/tr')) > 0)

model_color_set = set()
for row in driver.find_elements(By.XPATH, '//*[@id="modelColorDt"]/tbody/tr'):
    try:
        model_color_set.add(row.find_element(By.XPATH, "./td[1]/a").text.strip())
    except:
        pass

# ==================================================
# ACCESSORIES
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
# TRIMS
# ==================================================
wait.until(EC.element_to_be_clickable((
    By.XPATH, '/html/body/main/div[2]/div[2]/div/ui-view/div/div[2]/div[2]/header/div/div/ul/li[3]/a'
))).click()

while True:
    wait.until(lambda d: len(d.find_elements(By.XPATH, '//*[@id="trimDt"]/tbody/tr')) > 0)

    trim_names = [r.find_element(By.XPATH, "./td[2]/a").text for r in driver.find_elements(By.XPATH, '//*[@id="trimDt"]/tbody/tr')]

    for trim_name in trim_names:
        try:
            row = wait.until(EC.element_to_be_clickable((
                By.XPATH, f'//table[@id="trimDt"]//tr[td[2]/a[text()="{trim_name}"]]'
            )))

            trim_status = row.find_element(By.XPATH, "./td[6]/span").text
            row.find_element(By.XPATH, "./td[2]/a").click()

            wait.until(EC.element_to_be_clickable((
                By.XPATH, '/html/body/main/div[2]/div[2]/div/ui-view/div/div[2]/div[2]/div[1]/header/div/div/ul/li[2]/a'
            ))).click()

            trim_colors = []
            try:
                wait.until(lambda d: len(d.find_elements(By.XPATH, '//*[@id="trimColorDt"]/tbody/tr')) > 0)
                for crow in driver.find_elements(By.XPATH, '//*[@id="trimColorDt"]/tbody/tr'):
                    trim_colors.append(crow.find_element(By.XPATH, "./td[1]/a").text.strip())
            except:
                pass

            combined = model_color_set.union(set(trim_colors))

            trim_color_map[trim_name] = {
                "Make": make,
                "Model": model,
                "Year": year,
                "Trim": trim_name,
                "Trim_Status": trim_status,
                "Colors": " | ".join(sorted(combined))
            }

            wait.until(EC.element_to_be_clickable((
                By.XPATH, '/html/body/main/div[2]/div[2]/breadcrumbs/div/ul/li[2]/span/a'
            ))).click()

        except Exception as e:
            print("Trim error:", trim_name, e)

    try:
        next_btn = driver.find_element(By.XPATH, '//*[@id="trimDt_next"]')
        if "disabled" in next_btn.get_attribute("class"):
            break
        next_btn.click()
        wait.until(EC.staleness_of(next_btn))
    except:
        break

# ==================================================
# SAVE
# ==================================================
trims_data = list(trim_color_map.values())

with pd.ExcelWriter("final_output.xlsx") as writer:
    pd.DataFrame(accessories_data).to_excel(writer, sheet_name="Accessories", index=False)
    pd.DataFrame(trims_data).to_excel(writer, sheet_name="Trims", index=False)

print("✅ FINAL FILE GENERATED")

driver.quit()
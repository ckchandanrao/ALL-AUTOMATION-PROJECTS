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

# Store data
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

# --------------------------------------------------
# STEP 3: Wait for dashboard
# --------------------------------------------------
wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))

# --------------------------------------------------
# STEP 4: Navigate to Models
# --------------------------------------------------
wait.until(EC.element_to_be_clickable((
    By.XPATH, '//*[@id="sideNavMenu"]/li[3]/a/span[2]'
))).click()

print("✅ Opened Models section")

# --------------------------------------------------
# STEP 5: Open Filter
# --------------------------------------------------
wait.until(EC.element_to_be_clickable((
    By.XPATH, '//*[@id="modeldt_column_filter"]'
))).click()

# --------------------------------------------------
# STEP 6: MODEL FILTER
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

print("✅ Model filter applied")

# Wait for table
wait.until(EC.presence_of_all_elements_located((
    By.XPATH, '//*[@id="modeldt"]/tbody/tr'
)))

# --------------------------------------------------
# STEP 7: YEAR FILTER
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

print("✅ Year filter applied")

# Wait for table refresh
wait.until(EC.presence_of_all_elements_located((
    By.XPATH, '//*[@id="modeldt"]/tbody/tr'
)))

# --------------------------------------------------
# STEP 8: CLICK MODEL
# --------------------------------------------------
wait.until(EC.element_to_be_clickable((
    By.XPATH, '//*[@id="modeldt"]/tbody/tr[2]/td[3]/a'
))).click()

print("✅ Model opened")

# --------------------------------------------------
# STEP 9: FETCH DETAILS
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

print("Make:", make)
print("Model:", model)
print("Year:", year)
print("Status:", status)

# Store data
data.append({
    "Make": make,
    "Model": model,
    "Year": year,
    "Status": status
})

# --------------------------------------------------
# STEP 10: SAVE TO EXCEL
# --------------------------------------------------
df = pd.DataFrame(data)
df.to_excel("model_data.xlsx", index=False)

print("✅ Data saved to model_data.xlsx")

# --------------------------------------------------
# DONE
# --------------------------------------------------
input("Press Enter to close browser...")
driver.quit()
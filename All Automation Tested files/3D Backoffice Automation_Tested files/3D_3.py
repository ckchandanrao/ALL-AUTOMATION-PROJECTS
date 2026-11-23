from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys

# Setup
driver = webdriver.Chrome()
driver.maximize_window()
wait = WebDriverWait(driver, 20)

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
# STEP 6: MODEL FILTER (accord)
# --------------------------------------------------
wait.until(EC.element_to_be_clickable((
    By.XPATH, '//*[@id="filter-items"]/div[3]/span/span/span'
))).click()

model_input = wait.until(EC.element_to_be_clickable((
    By.XPATH, '//*[@id="modeldt_Model"]/div/div/input'
)))

model_input.click()
model_input.send_keys(Keys.CONTROL + "a")
model_input.send_keys(Keys.DELETE)
model_input.send_keys("accord")

wait.until(EC.element_to_be_clickable((
    By.XPATH, '//*[@id="modeldt_Model"]/div/div/button/font/font'
))).click()

print("✅ Model filter applied")

# 🔥 WAIT for table to refresh before next action
wait.until(EC.presence_of_all_elements_located((
    By.XPATH, '//*[@id="modeldt"]/tbody/tr'
)))

# --------------------------------------------------
# STEP 7: YEAR FILTER (DIRECT)
# --------------------------------------------------
wait.until(EC.element_to_be_clickable((
    By.XPATH, '//*[@id="filter-items"]/div[5]/span/span/span'
))).click()

print("✅ Year filter opened")

year_input = wait.until(EC.element_to_be_clickable((
    By.XPATH, '//*[@id="modeldt_Year"]/div/div/input'
)))

year_input.click()
year_input.send_keys(Keys.CONTROL + "a")
year_input.send_keys(Keys.DELETE)
year_input.send_keys("2025")

wait.until(EC.element_to_be_clickable((
    By.XPATH, '//*[@id="modeldt_Year"]/div/div/button'
))).click()

print("✅ Year filter applied")

# 🔥 WAIT for updated results
wait.until(EC.presence_of_all_elements_located((
    By.XPATH, '//*[@id="modeldt"]/tbody/tr'
)))

# --------------------------------------------------
# DONE
# --------------------------------------------------
input("Press Enter to close browser...")
driver.quit()
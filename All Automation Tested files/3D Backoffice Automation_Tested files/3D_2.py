from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import time

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
# STEP 3: Wait for dashboard to load
# --------------------------------------------------
wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
time.sleep(3)

# --------------------------------------------------
# STEP 4: Click Side Menu → Models
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
# STEP 6: Select Model Filter Option
# --------------------------------------------------
wait.until(EC.element_to_be_clickable((
    By.XPATH, '//*[@id="filter-items"]/div[3]/span/span/span'
))).click()

# --------------------------------------------------
# STEP 7: Enter Model Name (accord)
# --------------------------------------------------
model_input = wait.until(EC.presence_of_element_located((
    By.XPATH, '//*[@id="modeldt_Model"]/div/div/input'
)))
model_input.clear()
model_input.send_keys("accord")

# Click Search button
wait.until(EC.element_to_be_clickable((
    By.XPATH, '//*[@id="modeldt_Model"]/div/div/button/font/font'
))).click()

print("✅ Model 'accord' searched")

time.sleep(3)

# --------------------------------------------------
# STEP 8: Search MAKE (Honda)
# --------------------------------------------------
make_input = wait.until(EC.presence_of_element_located((
    By.XPATH, '//*[@id="overlaybody"]/div[2]/form/input[1]'
)))
make_input.clear()
make_input.send_keys("honda")
make_input.send_keys(Keys.ENTER)

print("✅ Make 'honda' searched")

# --------------------------------------------------
# DONE (Keep browser open)
# --------------------------------------------------
input("Press Enter to close browser...")
driver.quit()
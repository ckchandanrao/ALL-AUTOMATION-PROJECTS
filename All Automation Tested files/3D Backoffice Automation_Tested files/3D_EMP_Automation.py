from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Setup
driver = webdriver.Chrome()
driver.maximize_window()
wait = WebDriverWait(driver, 15)

# Open URL
driver.get("https://sandbox-admin.partsgorilla.com/login")

# Step 1: Enter Username
username = wait.until(EC.presence_of_element_located((
    By.XPATH,
    "/html/body/main/div/div/div/ui-view/div[1]/div/div/div/div[2]/div/form/div[1]/input"
)))
username.send_keys("izmosupport")

# Step 2: Enter Password
password = driver.find_element(
    By.XPATH,
    "/html/body/main/div/div/div/ui-view/div[1]/div/div/div/div[2]/div/form/div[2]/input"
)
password.send_keys("adm!n@supp@rt")

# Step 3: Click Login
login_btn = driver.find_element(
    By.XPATH,
    "/html/body/main/div/div/div/ui-view/div[1]/div/div/div/div[2]/div/form/div[3]/input"
)
login_btn.click()

# Wait for next page to load (adjust selector based on dashboard)
wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))

print("✅ Login successful")

# Keep browser open for inspection
input("Press Enter to close...")
driver.quit()
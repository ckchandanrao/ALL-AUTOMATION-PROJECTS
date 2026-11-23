from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time

options = Options()
options.add_experimental_option("detach", True)

driver = webdriver.Chrome(options=options)
driver.maximize_window()

driver.get("https://bmw-arles.stg-eudemosite.com/versions/bmw-fr-fr.htm")

all_models = []

# loop from page 1 to 6
for page in range(1, 7):
    print(f"\n=== Scraping page {page} ===")

    # wait for model list on this page
    WebDriverWait(driver, 10).until(
        EC.presence_of_all_elements_located((By.CLASS_NAME, "model-name"))
    )
    time.sleep(1)  # small extra safety wait

    # 1️⃣ get all model names on this page
    elements = driver.find_elements(By.CLASS_NAME, "model-name")
    page_models = [el.text.strip() for el in elements if el.text.strip()]
    all_models.extend(page_models)

    print(f"Found {len(page_models)} models on page {page}")

    # 2️⃣ click the next page number (2,3,4,5,6)
    if page < 6:
        next_page = page + 1

        # ⬇️ TRY THIS LOCATOR FIRST (very generic: any element whose text == page number)
        next_page_el = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (By.XPATH, f"//*[normalize-space(text())='{next_page}']")
            )
        )

        # safer to click with JavaScript on React/MUI UIs
        driver.execute_script("arguments[0].click();", next_page_el)
        time.sleep(2)  # wait for next page to load

# 3️⃣ save everything to Excel
df = pd.DataFrame({"Model Name": all_models})
df.to_excel("models_list.xlsx", index=False)

print("\nDone!")
print("Total models collected:", len(all_models))
print("Excel saved as models_list.xlsx")

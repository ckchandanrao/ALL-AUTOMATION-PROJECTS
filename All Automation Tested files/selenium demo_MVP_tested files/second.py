from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import pandas as pd
import time

# Keep browser open
options = Options()
options.add_experimental_option("detach", True)

driver = webdriver.Chrome(options=options)
driver.maximize_window()

driver.get("https://bmw-arles.stg-eudemosite.com/versions/bmw-fr-fr.htm")

time.sleep(3)  # wait for page to load fully

# Fetch all elements
elements = driver.find_elements(By.CLASS_NAME, "model-name")

# Extract text only (removes empty values)
model_names = [el.text.strip() for el in elements if el.text.strip()]

# Print names
for name in model_names:
    print(name)

# Save to Excel
df = pd.DataFrame({"Model Name": model_names})
df.to_excel("models_list.xlsx", index=False)

print("Excel file generated successfully: models_list.xlsx")

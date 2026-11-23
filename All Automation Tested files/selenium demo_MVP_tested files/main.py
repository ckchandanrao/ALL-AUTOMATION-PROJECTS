from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options

# Keep browser open
options = Options()
options.add_experimental_option("detach", True)

# Start Chrome
driver = webdriver.Chrome(options=options)
driver.maximize_window()
# Open website
driver.get("https://bmw-arles.stg-eudemosite.com/versions/bmw-fr-fr.htm")
model_names=driver.find_elements(By.CLASS_NAME, "model-name")

for li in model_names:
    print(li.text)

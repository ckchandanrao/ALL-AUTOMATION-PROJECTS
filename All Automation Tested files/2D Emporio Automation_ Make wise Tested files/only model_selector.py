# =====================================================
# SELENIUM ONLY — IZMO EMPORIO
# FINAL (MODEL SELECTOR + OUTSIDE URL ONLY)
# MULTI URL SUPPORT
# =====================================================

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time
from datetime import datetime

# =====================================================
# INSERT ALL URLS HERE
# =====================================================
URLS = [
    "https://abarth.izmocars.fr/emporio-fr-fr.htm",
"https://alfa-romeo.izmocars.fr/emporio-fr-fr.htm",
"https://audi.izmocars.fr/emporio-fr-fr.htm",
"https://alpineemporio.eudemosite.com/emporio-fr-fr.htm",
"https://bmw.izmocars.fr/emporio-fr-fr.htm",
"https://citroen.izmocars.fr/emporio-fr-fr.htm",
"https://cupra.izmocars.fr/emporio-fr-fr.htm",
"https://ds.izmocars.fr/emporio-fr-fr.htm",
"https://fiat.izmocars.fr/emporio-fr-fr.htm",
"https://ford.izmocars.fr/emporio-fr-fr.htm",
"https://honda.izmocars.fr/emporio-fr-fr.htm",
"https://hyundai.izmocars.fr/emporio-fr-fr.htm",
"https://isuzu.izmocars.fr/emporio-fr-fr.htm",
"https://jaguar.izmocars.fr/emporio-fr-fr.htm",
"https://jeep.izmocars.fr/emporio-fr-fr.htm",
"https://kia.izmocars.fr/emporio-fr-fr.htm",
"https://landrover.izmocars.fr/emporio-fr-fr.htm",
"https://lexus.izmocars.fr/emporio-fr-fr.htm",
"https://mazda.izmocars.fr/emporio-fr-fr.htm",
"https://mercedes-benz.izmocars.fr/emporio-fr-fr.htm",
"https://mg.izmocars.fr/emporio-fr-fr.htm",
"https://mini.izmocars.fr/emporio-fr-fr.htm",
"https://mitsubishi-motors.izmocars.fr/emporio-fr-fr.htm",
"https://nissan.izmocars.fr/emporio-fr-fr.htm",
"https://opel.izmocars.fr/emporio-fr-fr.htm",
"https://porsche.izmocars.fr/emporio-fr-fr.htm",
"https://seat.izmocars.fr/emporio-fr-fr.htm",
"https://smart.izmocars.fr/emporio-fr-fr.htm",
"https://subaru.izmocars.fr/emporio-fr-fr.htm",
"https://suzuki.izmocars.fr/emporio-fr-fr.htm",
"https://skoda.izmocars.fr/emporio-fr-fr.htm",
"https://toyota.izmocars.fr/emporio-fr-fr.htm",
"https://volkswagen.izmocars.fr/emporio-fr-fr.htm",
"https://volvo.izmocars.fr/emporio-fr-fr.htm",
"https://leapmotoremporio2d.eudemosite.com/emporio-fr-fr.htm",
"https://lanciaemporio2d.eudemosite.com/emporio-fr-fr.htm",
"https://ford.izmocars.fr/emporio-fr-fr.htm",
"https://bydemporio.eudemosite.com/emporio-fr-fr.htm",
"https://dacia2dizmoemporio.eudemosite.com/emporio-fr-fr.htm",
"http://maserati2dizmoemporio.eudemosite.com/emporio-fr-fr.htm",
"http://peugeot2dizmoemporio.eudemosite.com/emporio-fr-fr.htm",
"http://renault2dizmoemporio.eudemosite.com/emporio-fr-fr.htm",
"https://maxus2d.eudemosite.com/emporio-fr-fr.htm",
"https://cadillac2demporio.eudemosite.com/emporio-fr-fr.htm",

    # add more links freely
]

# =====================================================
# HELPERS
# =====================================================
def enter_iframe(driver, wait):
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "iframe")))
    driver.switch_to.frame(driver.find_element(By.TAG_NAME, "iframe"))

# =====================================================
# SETUP
# =====================================================
driver = webdriver.Chrome()
wait = WebDriverWait(driver, 20)

results = []

# =====================================================
# MAIN LOOP
# =====================================================
for url in URLS:
    print(f"\n🌐 Checking URL → {url}")

    try:
        driver.get(url)
        driver.maximize_window()
        time.sleep(3)

        enter_iframe(driver, wait)

        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "h4.modelTitle")))
        model_titles = driver.find_elements(By.CSS_SELECTOR, "h4.modelTitle")
        model_images = driver.find_elements(By.CSS_SELECTOR, "img")

        model_count = len(model_titles)
        print(f"🚗 Models found: {model_count}")

        for i in range(model_count):
            model_name = model_titles[i].text.strip()

            outside_url = ""
            model_selector_status = "Fail"

            try:
                outside_url = model_images[i].get_attribute("src") or ""
                if ".webp" in outside_url.lower():
                    model_selector_status = "Pass"
            except:
                outside_url = ""

            results.append({
                "Source URL": url,
                "Model Name": model_name,
                "Model Selector": model_selector_status,
                "Outside Image URL": outside_url
            })

        driver.switch_to.default_content()

    except Exception as e:
        print(f"❌ Failed URL → {url}")
        print(e)

driver.quit()

# =====================================================
# EXPORT
# =====================================================
df = pd.DataFrame(results)
output_file = f"Emporio_ModelSelector_Only_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
df.to_excel(output_file, index=False)

print(f"\n✅ EXCEL GENERATED → {output_file}")
print(df)

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time
from urllib.parse import urlparse

#Function to Extract Year & Alias from Image URL
def extract_year_alias_from_img(img_src: str):
    """
    From a URL like:
    https://s3-eu-west-1.amazonaws.com/.../FR/2025/25bmw/25bmw120dftnmsprthb4bfr/640x480/bmw_120dftnmsprthb4bfr_lowaggressive.webp

    Return:
        year  = '2025'
        alias = '120dftnmsprthb4bfr'
    """
    year = ""
    alias = ""

    path = urlparse(img_src).path                  # /toolkit/newgen/FR/2025/25bmw/...
    segments = [s for s in path.split("/") if s]   # ["toolkit","newgen","FR","2025",...]

    # 1) find a 4-digit year (e.g. 2025)
    for s in segments:
        if s.isdigit() and len(s) == 4:
            year = s
            break

    # 2) alias from filename part after "bmw_"
    filename = segments[-1]                        # bmw_120dftnmsprthb4bfr_lowaggressive.webp
    base = filename.rsplit(".", 1)[0]              # bmw_120dftnmsprthb4bfr_lowaggressive
    parts = base.split("_")                        # ["bmw","120dftnmsprthb4bfr","lowaggressive"]
    if len(parts) >= 2:
        alias = parts[1]

    return year, alias


# ---------- SELENIUM SETUP ----------
options = Options()
options.add_experimental_option("detach", True)

driver = webdriver.Chrome(options=options)
driver.maximize_window()

driver.get("https://bmw-arles.stg-eudemosite.com/versions/bmw-fr-fr.htm")

wait = WebDriverWait(driver, 10)

all_rows = []   # will collect data from all 6 pages

# ---------- LOOP THROUGH PAGES 1 TO 6 ----------
for page in range(1, 7):
    print(f"\n=== Scraping page {page} ===")

    # wait until model names are present on this page
    wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME, "model-name")))
    time.sleep(1)  # small extra wait

    # 1) SCRAPE MODEL NAMES ON THIS PAGE
    model_elems = driver.find_elements(By.CLASS_NAME, "model-name")
    model_names = [el.text.strip() for el in model_elems if el.text.strip()]

    print(f"  Models found on page {page}: {len(model_names)}")

    # 2) SCRAPE CAR IMAGES THAT MATCH THE BMW S3 PATTERN ON THIS PAGE
    all_imgs = driver.find_elements(By.TAG_NAME, "img")
    car_img_srcs = []

    for img in all_imgs:
        src = img.get_attribute("src") or ""
        if "staticeu.izmocars.com/toolkit/newgen" in src and "bmw_" in src:
            car_img_srcs.append(src)

    print(f"  Car images found on page {page}: {len(car_img_srcs)}")

    # 3) PAIR MODELS WITH IMAGES (BY INDEX) AND EXTRACT YEAR & ALIAS
    count = min(len(model_names), len(car_img_srcs))

    for i in range(count):
        name = model_names[i]
        img_src = car_img_srcs[i]

        year, alias = extract_year_alias_from_img(img_src)

        print(f"    #{i+1}: {name}")
        print(f"       IMG SRC: {img_src}")
        print(f"       YEAR   : {year}")
        print(f"       ALIAS  : {alias}")

        all_rows.append({
            "Model Name": name,
            "Year": year,
            "Alias": alias
        })

    # if there are extra model names without a matching image, add them with blank year/alias
    for j in range(count, len(model_names)):
        name = model_names[j]
        print(f"    #{j+1}: {name} (no matching image)")
        all_rows.append({
            "Model Name": name,
            "Year": "",
            "Alias": ""
        })

    # 4) CLICK NEXT PAGE (except after page 6)
    if page < 6:
        next_page = page + 1
        print(f"  Going to page {next_page}...")

        # try to click the pagination number using its visible text
        # this XPath is generic: any clickable element whose text = next_page
        next_xpath = f"//*[normalize-space(text())='{next_page}']"

        # scroll a bit down to make sure pagination is in view
        driver.execute_script("window.scrollBy(0, 800);")

        el = wait.until(EC.element_to_be_clickable((By.XPATH, next_xpath)))
        driver.execute_script("arguments[0].click();", el)
        time.sleep(2)  # wait for next page to load


# ---------- SAVE EVERYTHING TO EXCEL ----------
df = pd.DataFrame(all_rows)
df.to_excel("models_all_pages_year_alias.xlsx", index=False)

print("\nScraping finished for all 6 pages.")
print("Excel created: models_all_pages_year_alias.xlsx")
# driver.quit()  # uncomment if you want to close the browser

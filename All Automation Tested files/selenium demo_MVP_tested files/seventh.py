from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import pandas as pd
import time
from urllib.parse import urlparse


def extract_year_alias_from_img(img_src: str, make_prefix: str):
    """
    From a URL like:
    https://.../FR/2025/25bmw/25bmw120dftnmsprthb4bfr/640x480/bmw_120dftnmsprthb4bfr_lowaggressive.webp

    Return:
        year  = '2025'
        alias = '120dftnmsprthb4bfr'

    make_prefix should be e.g. "bmw", "kia", "maserati", "mini"
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

    # 2) alias from filename part after "<make_prefix>_"
    filename = segments[-1]                        # e.g. bmw_120dftnmsprthb4bfr_lowaggressive.webp
    base = filename.rsplit(".", 1)[0]              # e.g. bmw_120dftnmsprthb4bfr_lowaggressive
    parts = base.split("_")                        # e.g. ["bmw","120dftnmsprthb4bfr","lowaggressive"]

    if len(parts) >= 2 and parts[0].lower() == make_prefix.lower():
        alias = parts[1]

    return year, alias


# ---------- AVAILABLE MAKES & URLs ----------
MAKE_URLS = {
    "bmw": "https://bmw-arles.stg-eudemosite.com/versions/bmw-fr-fr.htm",
    "kia": "https://kia-arras.stg-eudemosite.com/versions/kia-fr-fr.htm",
    "maserati": "https://maserati-annecy.stg-eudemosite.com/versions/maserati-fr-fr.htm",
    "mini": "https://mini-dax.stg-eudemosite.com/versions/mini-fr-fr.htm",
    "Skoda": "https://skoda-chateauroux.stg-eudemosite.com/versions/%C5%A1koda-fr-fr.htm",
}

# ---------- ASK USER WHICH MAKE TO TEST ----------
print("Available makes:")
for make in MAKE_URLS.keys():
    print(" -", make)

selected_make = input("\nEnter the make you want to test: ").strip().lower()

while selected_make not in MAKE_URLS:
    print("Invalid make! Please choose from:", ", ".join(MAKE_URLS.keys()))
    selected_make = input("Enter the make you want to test: ").strip().lower()

start_url = MAKE_URLS[selected_make]
print(f"\nScraping: {selected_make.upper()} → {start_url}")

# used to detect images like 'bmw_', 'kia_', 'maserati_', 'mini_'
file_prefix = selected_make.lower() + "_"


# ---------- SELENIUM SETUP ----------
options = Options()
options.add_experimental_option("detach", True)

driver = webdriver.Chrome(options=options)
driver.maximize_window()
driver.get(start_url)

wait = WebDriverWait(driver, 10)

all_rows = []   # will collect data from all pages

# ---------- LOOP THROUGH PAGES USING NEXT ARROW ----------
page = 1

while True:
    print(f"\n=== Scraping page {page} ===")

    # wait until model names are present on this page
    wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME, "model-name")))
    time.sleep(1)  # small extra wait

    # 1) SCRAPE MODEL NAMES ON THIS PAGE
    model_elems = driver.find_elements(By.CLASS_NAME, "model-name")
    model_names = [el.text.strip() for el in model_elems if el.text.strip()]

    print(f"  Models found on page {page}: {len(model_names)}")

    # 2) SCRAPE CAR IMAGES THAT MATCH THE PATTERN ON THIS PAGE
    all_imgs = driver.find_elements(By.TAG_NAME, "img")
    car_img_srcs = []

    for img in all_imgs:
        src = img.get_attribute("src") or ""
        # generic for all makes: image is hosted in toolkit/newgen and filename contains "<make>_"
        if "staticeu.izmocars.com/toolkit/newgen" in src and file_prefix in src:
            car_img_srcs.append(src)

    print(f"  Car images found on page {page}: {len(car_img_srcs)}")

    # 3) PAIR MODELS WITH IMAGES (BY INDEX) AND EXTRACT YEAR & ALIAS
    count = min(len(model_names), len(car_img_srcs))

    for i in range(count):
        name = model_names[i]
        img_src = car_img_srcs[i]

        year, alias = extract_year_alias_from_img(img_src, selected_make)

        print(f"    #{i+1}: {name}")
        print(f"       IMG SRC: {img_src}")
        print(f"       YEAR   : {year}")
        print(f"       ALIAS  : {alias}")

        all_rows.append({
            "Make": selected_make.upper(),
            "Model Name": name,
            "Year": year,
            "Alias": alias
        })

    # if there are extra model names without a matching image, add them with blank year/alias
    for j in range(count, len(model_names)):
        name = model_names[j]
        print(f"    #{j+1}: {name} (no matching image)")
        all_rows.append({
            "Make": selected_make.upper(),
            "Model Name": name,
            "Year": "",
            "Alias": ""
        })

    # 4) FIND NEXT ARROW; IF DISABLED OR MISSING → STOP
    try:
        next_arrow = driver.find_element(By.CSS_SELECTOR, ".paginationPageNumber.nextArrow")
    except NoSuchElementException:
        print("  No next arrow found. Stopping pagination.")
        break

    classes = (next_arrow.get_attribute("class") or "").lower()
    aria_disabled = (next_arrow.get_attribute("aria-disabled") or "").lower()

    if "disabled" in classes or aria_disabled == "true":
        print("  Next arrow is disabled. Reached last page.")
        break

    print("  Clicking next arrow to go to next page...")

    # scroll into view & click
    driver.execute_script("arguments[0].scrollIntoView(true);", next_arrow)
    try:
        wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".paginationPageNumber.nextArrow")))
        driver.execute_script("arguments[0].click();", next_arrow)
    except TimeoutException:
        print("  Next arrow not clickable. Stopping pagination.")
        break

    time.sleep(2)  # wait for next page to load
    page += 1


# ---------- SAVE EVERYTHING TO EXCEL ----------
df = pd.DataFrame(all_rows)

excel_filename = f"{selected_make}_models_all_pages_year_alias.xlsx"
df.to_excel(excel_filename, index=False)

print("\nScraping finished for all pages.")
print(f"Excel created: {excel_filename}")

# driver.quit()  # uncomment if you want to close the browser at the end

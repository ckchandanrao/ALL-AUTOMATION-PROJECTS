from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
    ElementNotInteractableException,
    ElementClickInterceptedException,
)
import pandas as pd
import time
from urllib.parse import urlparse


def extract_year_alias_from_img(img_src: str, make_prefix: str):
    year = ""
    alias = ""

    path = urlparse(img_src).path
    segments = [s for s in path.split("/") if s]

    # Extract 4-digit year
    for s in segments:
        if s.isdigit() and len(s) == 4:
            year = s
            break

    # Extract alias
    filename = segments[-1]
    base = filename.rsplit(".", 1)[0]
    parts = base.split("_")

    if len(parts) >= 2 and parts[0].lower() == make_prefix.lower():
        alias = parts[1]

    return year, alias


# ---------- BASE PATH FOR PASS/FAIL CHECK ----------
BASE_IMG_PATH = "https://s3-eu-west-1.amazonaws.com/staticeu.izmocars.com/toolkit/newgen/"


# ---------- MAKE LIST ----------
MAKE_URLS = {
    "bmw": "https://bmw-arles.stg-eudemosite.com/versions/bmw-fr-fr.htm",
    "kia": "https://kia-arras.stg-eudemosite.com/versions/kia-fr-fr.htm",
    "maserati": "https://maserati-annecy.stg-eudemosite.com/versions/maserati-fr-fr.htm",
    "mini": "https://mini-dax.stg-eudemosite.com/versions/mini-fr-fr.htm",
    "skoda": "https://skoda-chateauroux.stg-eudemosite.com/versions/%C5%A1koda-fr-fr.htm",
}


# ---------- ASK USER FOR MAKE ----------
print("Available makes:")
for make in MAKE_URLS:
    print(" -", make)

selected_make = input("\nEnter the make you want to test: ").strip().lower()

while selected_make not in MAKE_URLS:
    print("Invalid make! Choose from:", ", ".join(MAKE_URLS.keys()))
    selected_make = input("Enter make: ").strip().lower()

start_url = MAKE_URLS[selected_make]
file_prefix = selected_make.lower() + "_"

print(f"\nScraping {selected_make.upper()} → {start_url}\n")


# ---------- SELENIUM ----------
options = Options()
options.add_experimental_option("detach", True)

driver = webdriver.Chrome(options=options)
driver.maximize_window()
driver.get(start_url)

wait = WebDriverWait(driver, 10)

all_rows = []


# ---------- LOOP PAGES UNTIL NEXT ARROW DISABLED ----------
page = 1

while True:
    print(f"\n=== Scraping Page {page} ===")

    wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME, "model-name")))
    time.sleep(1)

    # Extract model names
    model_elems = driver.find_elements(By.CLASS_NAME, "model-name")
    model_names = [el.text.strip() for el in model_elems if el.text.strip()]

    print(f"  Models found: {len(model_names)}")

    # Extract images
    all_imgs = driver.find_elements(By.TAG_NAME, "img")
    car_img_srcs = []
    for img in all_imgs:
        src = img.get_attribute("src") or ""
        if "staticeu.izmocars.com/toolkit/newgen" in src and file_prefix in src.lower():
            car_img_srcs.append(src)

    print(f"  Car images found: {len(car_img_srcs)}")

    count = min(len(model_names), len(car_img_srcs))

    # Pair models with images
    for i in range(count):
        name = model_names[i]
        img_src = car_img_srcs[i]

        year, alias = extract_year_alias_from_img(img_src, selected_make)

        # PASS/FAIL LOGIC
        model_selector_status = "Pass" if BASE_IMG_PATH in img_src else "Fail"

        all_rows.append({
            "Make": selected_make.upper(),
            "Model Name": name,
            "Year": year,
            "Alias": alias,
            "Model selector": model_selector_status
        })

    # Handle models without images
    for j in range(count, len(model_names)):
        name = model_names[j]
        all_rows.append({
            "Make": selected_make.upper(),
            "Model Name": name,
            "Year": "",
            "Alias": "",
            "Model selector": "Fail"
        })

    # ---------- PAGINATION LOGIC (FIXED CLICK) ----------
    try:
        next_arrow = driver.find_element(By.CSS_SELECTOR, ".paginationPageNumber.nextArrow")
    except NoSuchElementException:
        print("  No next arrow. Stopping.")
        break

    classes = (next_arrow.get_attribute("class") or "").lower()
    disabled = (next_arrow.get_attribute("aria-disabled") or "").lower()

    if "disabled" in classes or disabled == "true":
        print("  Next arrow disabled → last page reached.")
        break

    print("  Moving to next page...")

    try:
        # scroll into view
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", next_arrow)
        # wait until clickable
        wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".paginationPageNumber.nextArrow")))
        # use JS click to avoid ElementNotInteractable
        driver.execute_script("arguments[0].click();", next_arrow)
    except (TimeoutException, ElementNotInteractableException, ElementClickInterceptedException) as e:
        print(f"  Could not click next arrow ({type(e).__name__}). Stopping pagination.")
        break

    time.sleep(2)
    page += 1


# ---------- SAVE EXCEL ----------
df = pd.DataFrame(all_rows)
excel_filename = f"{selected_make}_models_all_pages_year_alias.xlsx"
df.to_excel(excel_filename, index=False)

print("\nScraping complete!")
print(f"Excel saved: {excel_filename}")
# driver.quit()

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


def extract_year_alias_from_img(img_src: str):
    """
    Try to extract year and alias from an image URL like:
    https://.../FR/2025/25bmw/25bmw120dftnmsprthb4bfr/640x480/bmw_120dftnmsprthb4bfr_lowaggressive.webp

    Returns:
        year  -> '2025'
        alias -> '120dftnmsprthb4bfr'

    If it cannot find them, returns empty strings.
    """
    year = ""
    alias = ""

    if not img_src:
        return year, alias

    path = urlparse(img_src).path
    segments = [s for s in path.split("/") if s]

    # 1) Extract a 4-digit year from the path
    for s in segments:
        if s.isdigit() and len(s) == 4:
            year = s
            break

    # 2) Alias from filename part after "<brand>_"
    if segments:
        filename = segments[-1]                      # e.g. bmw_120dftnmsprthb4bfr_lowaggressive.webp
        base = filename.rsplit(".", 1)[0]           # e.g. bmw_120dftnmsprthb4bfr_lowaggressive
        parts = base.split("_")                     # e.g. ["bmw","120dftnmsprthb4bfr","lowaggressive"]
        if len(parts) >= 2:
            alias = parts[1]

    return year, alias


# ---------- RAW MAKE LIST (original mapping) ----------
RAW_MAKE_URLS = {
    "bmw": "https://bmw-arles.stg-eudemosite.com/versions/bmw-fr-fr.htm",
    "kia": "https://kia-arras.stg-eudemosite.com/versions/kia-fr-fr.htm",
    "maserati": "https://maserati-annecy.stg-eudemosite.com/versions/maserati-fr-fr.htm",
    "mini": "https://mini-dax.stg-eudemosite.com/versions/mini-fr-fr.htm",
    "skoda": "https://skoda-chateauroux.stg-eudemosite.com/versions/%C5%A1koda-fr-fr.htm",
    "suzuki": "https://suzuki-bagneux.stg-eudemosite.com/versions/suzuki-fr-fr.htm",
    "Alpine": "https://alpine-frejus.stg-eudemosite.com/versions/alpine-fr-fr.htm",
    "byd": "https://byd-chartres.stg-eudemosite.com/versions/byd-fr-fr.htm",
    "mitsubishi": "https://mitsubishi-motors-miho.stg-eudemosite.com/versions/mitsubishi-fr-fr.htm",
    "seat": "https://seat-charleville.stg-eudemosite.com/versions/seat-fr-fr.htm",
    "smart": "- Smart MELUN",
    "Ford": "https://ford-arras.stg-eudemosite.com/versions/ford-fr-fr.htm",
    "opel": "https://opel-tours.stg-eudemosite.com/versions/opel-fr-fr.htm",
    "lexus": "https://lexus-chambourcy.stg-eudemosite.com/versions/lexus-fr-fr.htm",
    "fiat": "https://fiat-aixenprovence.stg-eudemosite.com/versions/fiat-fr-fr.htm",
    "dS": "https://dsautomobiles-salondeprovence.stg-eudemosite.com/versions/ds-fr-fr.htm",
    "toyota": "https://toyota-asnieres.stg-eudemosite.com/versions/toyota-fr-fr.htm",
    "peugeot": "https://peugeot-saintmaximin.stg-eudemosite.com/versions/peugeot-fr-fr.htm",
    "landrover": "https://landrovermodelpages.stg-eudemosite.com/versions/land-rover-fr-fr.htm",
    "cupra": "https://cupra-chateauroux.stg-eudemosite.com/versions/cupra-fr-fr.htm",
    "citroen": "https://citroen-arras.stg-eudemosite.com/versions/citroen-fr-fr.htm",
    "fiat professional": "https://fiat-aixenprovence.stg-eudemosite.com/versions/fiat-professional-fr-fr.htm",
    "Dacia": "- Dacia Arras",
    "Renault": "https://renault-annemasse.stg-eudemosite.com/versions/renault-fr-fr.htm",
    "Jeep": "https://jeepmodelpages.stg-eudemosite.com/versions/jeep-fr-fr.htm",
    "Jaguar": "https://jaguarmodelpages.stg-eudemosite.com/versions/jaguar-fr-fr.htm",
    "Volkswagen": "|Volkswagen Model Pages",
    "Volvo": "- Volvo Charleville-Mézières",
    "Audi": "https://audi-angouleme.stg-eudemosite.com/versions/audi-fr-fr.htm",
    "Mercedes Benz": "|Mercedes-Benz Model Pages",
    "Honda": "Les dernières actualités chez Honda Lyon Sud",
    "Hyundai": "https://hyundai-amiens-poulainville.stg-eudemosite.com/versions/hyundai-fr-fr.htm",
    "Infiniti": "|Infiniti Viry Châtillon",
    "Subaru": "https://subarumodelpages.stg-eudemosite.com/versions/subaru-outback-fr-fr.htm",
    "MINI": "- MINI DAX",
    "alfa Romeo": "|AVICARS AIX EN PROVENCE",
    "nissan": "- Nissan Cambrai",
    "Mercedes-Benz Utilitaires": "https://mercedes-benz-beauvais-utilitaires.stg-eudemosite.com/versions/mercedes-benz-utilitaires-fr-fr.htm",
    "Seat": "https://seat-charleville.stg-eudemosite.com/versions/seat-fr-fr.htm",
    "Mazda": "|MAZDA CANNES",
    "MG": "|MG Car East France",
    "Porsche": "Porsche LILLE : Concessionnaire & garage - Nord 59",
    "Volkswagen Utilitaires": "https://volkswagen-mantes.stg-eudemosite.com/versions/volkswagen-utilitaires-fr-fr.htm",
    "Leap": "https://leapmotor-generic-demosite.stg-eudemosite.com/versions/leapmotor-fr-fr.htm",
}

# ---------- CLEANED MAKE LIST (lowercase keys, only valid URLs) ----------
MAKE_URLS = {}
for name, url in RAW_MAKE_URLS.items():
    if isinstance(url, str) and url.lower().startswith("http"):
        key = name.strip().lower()
        MAKE_URLS[key] = url

# ---------- ASK USER FOR MAKE ----------
print("Available makes:")
for make in MAKE_URLS:
    print(" -", make)

selected_make = input("\nEnter the make you want to test: ").strip().lower()

while selected_make not in MAKE_URLS:
    print("Invalid make! Choose from:", ", ".join(MAKE_URLS.keys()))
    selected_make = input("Enter make: ").strip().lower()

start_url = MAKE_URLS[selected_make]
print(f"\nScraping {selected_make.upper()} → {start_url}\n")


# ---------- SELENIUM ----------
options = Options()
options.add_experimental_option("detach", True)

driver = webdriver.Chrome(options=options)
driver.maximize_window()
driver.get(start_url)

wait = WebDriverWait(driver, 10)
all_rows = []


# ---------- LOOP PAGES USING NEXT ARROW (works even if no page numbers) ----------
page_index = 1

while True:
    print(f"\n=== Scraping Page {page_index} ===")

    # Wait for at least some models
    wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME, "model-name")))
    time.sleep(1)

    # ---------- SCROLL THROUGH PAGE TO LOAD ALL LAZY MODELS ----------
    last_height = driver.execute_script("return document.body.scrollHeight")
    while True:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(2)
        new_height = driver.execute_script("return document.body.scrollHeight")
        if new_height == last_height:
            break
        last_height = new_height

    # 1) Extract model names (after full scroll)
    model_elems = driver.find_elements(By.CLASS_NAME, "model-name")
    model_names = [el.text.strip() for el in model_elems if el.text.strip()]
    print(f"  Models found after scrolling: {len(model_names)}")

    # 2) Extract model images (same class you gave)
    img_elems = driver.find_elements(By.CSS_SELECTOR, "img.img-responsive.model-img-t")
    print(f"  Model images found after scrolling: {len(img_elems)}")

    # 3) Pair models with images by index
    for idx, name in enumerate(model_names):
        img_src = ""
        if idx < len(img_elems):
            img_src = img_elems[idx].get_attribute("src") or ""

        img_src_lower = img_src.lower() if img_src else ""

        # ---------- MODEL SELECTOR RULE ----------
        # If .jpg (coming soon) OR no image → Fail/Fail/Fail
        # If .webp → Pass + extract year & alias
        if (not img_src) or (".jpg" in img_src_lower):
            model_selector_status = "Fail"
            year_out = "Fail"
            alias_out = "Fail"
        elif ".webp" in img_src_lower:
            model_selector_status = "Pass"
            year, alias = extract_year_alias_from_img(img_src)
            year_out = year if year else ""
            alias_out = alias if alias else ""
        else:
            # Any other (png, etc.) → treat as Fail
            model_selector_status = "Fail"
            year_out = "Fail"
            alias_out = "Fail"

        print(f"  #{idx+1}: {name}")
        print(f"     IMG  : {img_src}")
        print(f"     YEAR : {year_out}")
        print(f"     ALIAS: {alias_out}")
        print(f"     Model selector: {model_selector_status}")

        all_rows.append({
            "Make": selected_make.upper(),
            "Model Name": name,
            "Year": year_out,
            "Alias": alias_out,
            "Model selector": model_selector_status
        })

    # ---------- TRY NEXT ARROW ----------
    try:
        next_arrow = driver.find_element(By.CSS_SELECTOR, ".paginationPageNumber.nextArrow")
    except NoSuchElementException:
        print("  No next arrow. Stopping pagination.")
        break

    classes = (next_arrow.get_attribute("class") or "").lower()
    disabled = (next_arrow.get_attribute("aria-disabled") or "").lower()

    if "disabled" in classes or disabled == "true":
        print("  Next arrow disabled → last page reached.")
        break

    print("  Going to next page...")

    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", next_arrow)
        wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".paginationPageNumber.nextArrow")))
        driver.execute_script("arguments[0].click();", next_arrow)
    except (TimeoutException, ElementNotInteractableException, ElementClickInterceptedException) as e:
        print(f"  Cannot click next arrow: {e}. Stopping pagination.")
        break

    time.sleep(2)
    page_index += 1


# ---------- SAVE EXCEL ----------
df = pd.DataFrame(all_rows)
excel_filename = f"{selected_make}_models_all_pages_year_alias.xlsx"
df.to_excel(excel_filename, index=False)

print("\nScraping complete!")
print(f"Excel saved: {excel_filename}")
# driver.quit()

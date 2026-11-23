from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import pandas as pd
import time
from urllib.parse import urlparse


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
time.sleep(3)   # wait for page 1 to load


# ---------- 1) SCRAPE ALL MODEL NAMES ON PAGE 1 ----------
model_elems = driver.find_elements(By.CLASS_NAME, "model-name")
model_names = [el.text.strip() for el in model_elems if el.text.strip()]

print("Models on page 1:")
for m in model_names:
    print(" -", m)

# rows for Excel – Year/Alias will be empty first, then we fill first row
rows = [{"Model Name": name, "Year": "", "Alias": ""} for name in model_names]


# ---------- 2) FIND FIRST CAR IMAGE, OPEN IN NEW TAB, EXTRACT YEAR & ALIAS ----------
car_img_src = ""

# find first <img> that looks like a car image (matches your S3 pattern)
for img in driver.find_elements(By.TAG_NAME, "img"):
    src = img.get_attribute("src") or ""
    if "staticeu.izmocars.com/toolkit/newgen" in src and "bmw_" in src:
        car_img_src = src
        break

print("\nCar image src found on listing page:", car_img_src)

year = ""
alias = ""

if car_img_src:
    # open image in new tab (as you requested)
    main_window = driver.current_window_handle
    driver.switch_to.new_window("tab")
    driver.get(car_img_src)
    time.sleep(2)

    # try to find <img> in the new tab (if there is HTML wrapper)
    try:
        img_in_tab = driver.find_element(By.TAG_NAME, "img")
        final_src = img_in_tab.get_attribute("src") or car_img_src
    except Exception:
        final_src = car_img_src

    print("Final img src used for parsing:", final_src)

    year, alias = extract_year_alias_from_img(final_src)

    print("Parsed YEAR:", year)
    print("Parsed ALIAS:", alias)

    # close image tab and go back to main page
    driver.close()
    driver.switch_to.window(main_window)

# put year & alias into first row only
if rows and year and alias:
    rows[0]["Year"] = year
    rows[0]["Alias"] = alias


# ---------- 3) SAVE TO EXCEL ----------
df = pd.DataFrame(rows)
df.to_excel("models_page1_year_alias_from_img.xlsx", index=False)

print("\nExcel created: models_page1_year_alias_from_img.xlsx")
# driver.quit()  # uncomment if you want to close browser when done

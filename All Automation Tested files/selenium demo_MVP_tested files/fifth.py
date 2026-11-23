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

print("Models on page 1:", len(model_names))
for m in model_names:
    print(" -", m)


# ---------- 2) SCRAPE ALL CAR IMAGES THAT MATCH BMW PATTERN ----------
all_imgs = driver.find_elements(By.TAG_NAME, "img")
car_img_srcs = []

for img in all_imgs:
    src = img.get_attribute("src") or ""
    # match only your BMW car images on S3
    if "staticeu.izmocars.com/toolkit/newgen" in src and "bmw_" in src:
        car_img_srcs.append(src)

print("\nCar images found on page 1:", len(car_img_srcs))
for s in car_img_srcs:
    print(" -", s)

# make sure we don't crash if counts differ
count = min(len(model_names), len(car_img_srcs))

rows = []

# ---------- 3) FOR EACH MODEL + IMAGE: EXTRACT YEAR & ALIAS ----------
for i in range(count):
    name = model_names[i]
    img_src = car_img_srcs[i]

    year, alias = extract_year_alias_from_img(img_src)

    print(f"\nModel {i+1}: {name}")
    print("  IMG SRC:", img_src)
    print("  YEAR   :", year)
    print("  ALIAS  :", alias)

    rows.append({
        "Model Name": name,
        "Year": year,
        "Alias": alias
    })

# If there are more models than images, add remaining with blank year/alias
for j in range(count, len(model_names)):
    rows.append({
        "Model Name": model_names[j],
        "Year": "",
        "Alias": ""
    })


# ---------- 4) SAVE TO EXCEL ----------
df = pd.DataFrame(rows)
df.to_excel("models_page1_year_alias_all_models.xlsx", index=False)

print("\nExcel created: models_page1_year_alias_all_models.xlsx")

# driver.quit()   # uncomment if you want to close browser at the end

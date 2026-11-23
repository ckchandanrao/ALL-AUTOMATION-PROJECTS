from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import re
import time
import asyncio
import aiohttp
import unicodedata
from datetime import datetime
from urllib.parse import unquote
from collections import Counter
from openpyxl.chart import BarChart, Reference
from openpyxl.utils.dataframe import dataframe_to_rows
import requests

# =====================================================
# 🔥 ADD MULTIPLE URLS HERE
# =====================================================

URLS = [
    "https://acura.izmocars.com/emporio-en-us.htm",
    "https://volvo.izmocars.fr/emporio-fr-fr.htm",
    "https://skodaemporio2d.eudemosite.com/emporio-es-es.htm",
]

PARALLEL_REQUESTS = 100
RETRIES = 2


# =====================================================
# HELPERS
# =====================================================

def extract_year_make_alias_from_s3(url):
    if not url:
        return "NOT_CAPTURED", "NOT_CAPTURED", "NOT_CAPTURED"

    parts = url.split("/")
    try:
        year = next(p for p in parts if re.fullmatch(r"\d{4}", p))
        idx = parts.index(year)
        make = re.sub(r"^\d+", "", parts[idx + 1])
        alias = parts[idx + 2]

        return year, unquote(make), unquote(alias)
    except:
        return "NOT_CAPTURED", "NOT_CAPTURED", "NOT_CAPTURED"


def enter_iframe(driver, wait):
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "iframe")))
    driver.switch_to.frame(driver.find_element(By.TAG_NAME, "iframe"))


def check_ext_hotspot(driver, timeout=12):
    end = time.time() + timeout
    while time.time() < end:
        elements = driver.find_elements(By.CSS_SELECTOR, ".control-block.eyeButton")
        for el in elements:
            if el.is_displayed() and el.is_enabled():
                return "PASS"
        time.sleep(0.5)
    return "FAIL"


# =====================================================
# REGION DETECTION
# =====================================================

def get_region_from_url(url):
    url = url.lower()
    if "-fr" in url:
        return "FR"
    elif "-es" in url:
        return "ES"
    elif "-us" in url:
        return "US"
    return "FR"


# =====================================================
# NORMALIZATION
# =====================================================

def normalize_text(text):
    normalized = unicodedata.normalize("NFKD", text)
    return normalized.encode("ASCII", "ignore").decode("ASCII").strip().lower()


def normalize_color_name(name):
    return normalize_text(name).replace(" ", "-")


# =====================================================
# COLOR CHECK
# =====================================================

def get_color_names_and_duplicate_status(driver):
    color_names = []

    try:
        color_tiles = driver.find_elements(
            By.XPATH,
            "//*[@id='color-container']//*[@ng-click=\"selectColorTlc(color,color.colorName);\"]"
        )

        last_text = ""

        for tile in color_tiles:
            try:
                driver.execute_script("arguments[0].click();", tile)

                WebDriverWait(driver, 8).until(
                    lambda d: d.find_element(
                        By.CSS_SELECTOR,
                        ".fontFamilyBold.fontSize16.ng-binding.ng-scope"
                    ).text.strip() != last_text
                )

                name = driver.find_element(
                    By.CSS_SELECTOR,
                    ".fontFamilyBold.fontSize16.ng-binding.ng-scope"
                ).text.strip()

                if name:
                    color_names.append(name)
                    last_text = name

            except:
                pass
    except:
        pass

    normalized = [normalize_text(c) for c in color_names]
    counts = Counter(normalized)

    row_status = [
        "FAIL" if counts[color] > 1 else "PASS"
        for color in normalized
    ]

    return color_names, row_status


# =====================================================
# INTERIOR PANO CHECK
# =====================================================

def check_interior_pano(year, make, alias, region):
    year_short = year[-2:]
    year_make = f"{year_short}{make.lower()}"

    if region == "US":
        base = "https://s3-us-west-2.amazonaws.com/static.izmocars.com/toolkit/newgen/"
    else:
        base = "https://s3-eu-west-1.amazonaws.com/staticeu.izmocars.com/toolkit/newgen/"

    pano_url = f"{base}{year}/{year_make}/{alias}/{alias}_panovr/"

    try:
        r = requests.head(pano_url, timeout=10)
        return "PASS" if r.status_code == 200 else "FAIL"
    except:
        return "FAIL"


# =====================================================
# BLUE HOTSPOT CHECK
# =====================================================

def check_blue_hotspot(year, make, alias, region):
    year_short = year[-2:]
    year_make = f"{year_short}{make.lower()}"

    if region == "US":
        base = "https://s3-us-west-2.amazonaws.com/static.izmocars.com/toolkit/newgen/"
    else:
        base = "https://s3-eu-west-1.amazonaws.com/staticeu.izmocars.com/toolkit/newgen/"

    hotspot_url = f"{base}{year}/{year_make}/{alias}/{alias}_panovr/skin/hs.webp"

    try:
        r = requests.head(hotspot_url, timeout=10)
        return "PASS" if r.status_code == 200 else "FAIL"
    except:
        return "FAIL"


# =====================================================
# S3 COLOR PATH GENERATION
# =====================================================

def check_colors_broken(model_name, year, make, alias, color_list, s3_results, region):

    if region == "US":
        base_url = f"https://static-h2.izmocars.com/toolkit/newgen/{region}"
    else:
        base_url = f"https://staticeu-h2.izmocars.com/toolkit/newgen/{region}"

    year_short = year[-2:]
    year_make = f"{year_short}{make.lower()}"

    base_prefix = (
        f"{base_url}/{year}/{year_make}/{alias}/"
        f"{alias}_animations/_360color/_660x440"
    )

    for color_name in color_list:
        color_folder = normalize_color_name(color_name)

        for angle in range(1, 37):
            angle_url = f"{base_prefix}/{color_folder}/angle{angle}.webp"

            s3_results.append({
                "Model Name": model_name,
                "Year": year,
                "Make": make,
                "Alias": alias,
                "Color Name": color_name,
                "Angle": angle,
                "S3 URL": angle_url
            })


# =====================================================
# ASYNC VALIDATION
# =====================================================

async def check_url(session, url):
    for attempt in range(RETRIES):
        try:
            async with session.head(url) as resp:
                return "PASS" if resp.status == 200 else f"BROKEN ({resp.status})"
        except:
            if attempt == RETRIES - 1:
                return "BROKEN"


async def validate_s3_links_async(s3_df):
    connector = aiohttp.TCPConnector(limit=PARALLEL_REQUESTS)

    async with aiohttp.ClientSession(connector=connector) as session:
        tasks = [check_url(session, url) for url in s3_df["S3 URL"].tolist()]
        statuses = await asyncio.gather(*tasks)

    s3_df["S3 Status"] = statuses
    return s3_df


# =====================================================
# MODEL SELECTOR CHECK
# =====================================================

def check_model_selector(driver, model_index):
    try:
        images = driver.find_elements(By.CSS_SELECTOR, "img")
        outside_s3 = images[model_index].get_attribute("src") or ""

        return "PASS" if ".webp" in outside_s3.lower() else "FAIL"
    except:
        return "FAIL"


# =====================================================
# MAIN EXECUTION
# =====================================================

results = []
s3_results = []

for URL in URLS:
    region = get_region_from_url(URL)
    print(f"\n🌍 Processing URL: {URL}")

    driver = webdriver.Chrome()
    wait = WebDriverWait(driver, 25)

    driver.get(URL)
    driver.maximize_window()

    enter_iframe(driver, wait)
    time.sleep(3)

    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "h4.modelTitle")))
    model_count = len(driver.find_elements(By.CSS_SELECTOR, "h4.modelTitle"))

    for i in range(model_count):
        model_titles = driver.find_elements(By.CSS_SELECTOR, "h4.modelTitle")
        model_name = model_titles[i].text.strip()

        model_selector_status = check_model_selector(driver, i)

        driver.execute_script(
            "arguments[0].scrollIntoView({block:'center'}); arguments[0].click();",
            model_titles[i]
        )

        time.sleep(3)

        inside_s3 = ""
        try:
            img2 = WebDriverWait(driver, 6).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "img[src*='toolkit']"))
            )
            inside_s3 = img2.get_attribute("src")
        except:
            pass

        year, make, alias = extract_year_make_alias_from_s3(inside_s3)

        color_list, duplicate_status = get_color_names_and_duplicate_status(driver)
        color_count = len(color_list)

        check_colors_broken(
            model_name, year, make, alias,
            color_list, s3_results, region
        )

        ext_hotspot_status = check_ext_hotspot(driver)
        interior_pano_status = check_interior_pano(year, make, alias, region)
        blue_hotspot_status = check_blue_hotspot(year, make, alias, region)

        for idx, color in enumerate(color_list):
            results.append({
                "Year": year,
                "Make": make,
                "Model Name": model_name,
                "Alias": alias,
                "Model Selector Broken": model_selector_status,
                "Color Name": color,
                "Color Count": color_count,
                "Duplicate Colors": duplicate_status[idx],
                "Ext Hotspot": ext_hotspot_status,
                "Interior Pano": interior_pano_status,
                "Blue Hotspot": blue_hotspot_status,
                "Inside S3 URL": inside_s3
            })

        driver.switch_to.default_content()
        driver.back()
        time.sleep(2)
        enter_iframe(driver, wait)

    driver.quit()


# =====================================================
# VALIDATE S3 LINKS
# =====================================================

s3_df = pd.DataFrame(s3_results)
print("🔎 Validating S3 URLs...")
s3_df = asyncio.run(validate_s3_links_async(s3_df))


# =====================================================
# SUMMARY + GRAPH
# =====================================================

summary_df = s3_df.groupby("Model Name").agg(
    Total_URLs=("S3 URL", "count"),
    PASS_Count=("S3 Status", lambda x: (x == "PASS").sum()),
    BROKEN_Count=("S3 Status", lambda x: (x != "PASS").sum())
).reset_index()

summary_df["PASS %"] = round(
    (summary_df["PASS_Count"] / summary_df["Total_URLs"]) * 100, 2
)

summary_df["BROKEN %"] = round(
    (summary_df["BROKEN_Count"] / summary_df["Total_URLs"]) * 100, 2
)

output_file = f"Emporio_Final_Output_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    pd.DataFrame(results).to_excel(writer, sheet_name="Main Report", index=False)
    s3_df.to_excel(writer, sheet_name="Color Broken Status", index=False)
    summary_df.to_excel(writer, sheet_name="Model Summary", index=False)

    workbook = writer.book
    graph_sheet = workbook.create_sheet("Graph Summary")

    for r in dataframe_to_rows(summary_df, index=False, header=True):
        graph_sheet.append(r)

    bar = BarChart()
    bar.title = "Model-wise PASS vs BROKEN"

    bar.add_data(
        Reference(graph_sheet, min_col=2, max_col=3,
                  min_row=1, max_row=len(summary_df) + 1),
        titles_from_data=True
    )

    bar.set_categories(
        Reference(graph_sheet, min_col=1,
                  min_row=2, max_row=len(summary_df) + 1)
    )

    graph_sheet.add_chart(bar, "H2")

print("✅ Excel Generated Successfully")

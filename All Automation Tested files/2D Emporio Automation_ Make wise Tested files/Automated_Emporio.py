# =====================================================
# 🔥 EMPORIO AUTOMATION WITH EXCEL + UI
# =====================================================

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
import tkinter as tk
from tkinter import messagebox

INPUT_FILE = "Emporio_Input.xlsx"
PARALLEL_REQUESTS = 100
RETRIES = 2

# =====================================================
# LOAD URLS FROM EXCEL (US / FR / ES)
# =====================================================

def load_urls_from_excel():
    regions = ["US", "FR", "ES"]
    all_data = []

    for region in regions:
        try:
            df = pd.read_excel(INPUT_FILE, sheet_name=region)
            df.columns = df.columns.str.strip().str.upper()

            if not {"MAKE", "URL"}.issubset(df.columns):
                continue

            df["Region"] = region
            df = df[df["URL"].str.startswith("http")]
            all_data.append(df)

        except:
            continue

    if not all_data:
        raise Exception("No valid sheets (US, FR, ES) found in Excel.")

    return pd.concat(all_data).reset_index(drop=True)

# =====================================================
# UI SELECTION (Region Wise Select / Deselect)
# =====================================================
def select_urls_ui(execution_df):

    selected_indices = []

    root = tk.Tk()
    root.title("Emporio Execution Panel")
    root.geometry("1100x700")

    tk.Label(root,
             text="EMPORIO EXECUTION DASHBOARD",
             font=("Arial", 16, "bold"),
             anchor="w").pack(pady=10, anchor="w", padx=20)

    canvas = tk.Canvas(root)
    scrollbar = tk.Scrollbar(root, orient="vertical", command=canvas.yview)
    frame = tk.Frame(canvas)

    frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )

    canvas.create_window((0, 0), window=frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    region_groups = execution_df.groupby("Region")
    region_vars = {}

    for region, region_df in region_groups:

        # ===== REGION HEADER =====
        tk.Label(frame,
                 text=f"REGION : {region}",
                 font=("Arial", 14, "bold"),
                 fg="blue",
                 anchor="w").pack(anchor="w", padx=15, pady=15)

        region_vars[region] = []

        # ===== BUTTON FRAME =====
        btn_frame = tk.Frame(frame)
        btn_frame.pack(anchor="w", padx=20, pady=5)

        def select_all(r=region):
            for var in region_vars[r]:
                var.set(1)

        def deselect_all(r=region):
            for var in region_vars[r]:
                var.set(0)

        tk.Button(btn_frame,
                  text="Select All",
                  bg="green",
                  fg="white",
                  command=select_all).pack(side="left", padx=5)

        tk.Button(btn_frame,
                  text="Deselect All",
                  bg="red",
                  fg="white",
                  command=deselect_all).pack(side="left", padx=5)

        # ===== GRID FOR MAKES =====
        grid_frame = tk.Frame(frame)
        grid_frame.pack(anchor="w", padx=5, pady=5)
        


        row_num = 0
        col_num = 0

        for idx, row in region_df.iterrows():

            var = tk.IntVar(value=1)
            region_vars[region].append(var)

            chk = tk.Checkbutton(
                grid_frame,
                text=row["MAKE"],
                variable=var,
                width=18,
                anchor="w",
                justify="left"
            )

            chk.grid(row=row_num,
                     column=col_num,
                     padx=0,
                     pady=0,
                     sticky="w")

            selected_indices.append((var, idx))

            col_num += 1

            if col_num == 8:  # 8 per row
                col_num = 0
                row_num += 1

        

    # ===== RUN BUTTON =====
    def run_selected():
        final = []
        for var, idx in selected_indices:
            if var.get() == 1:
                final.append(idx)

        if not final:
            messagebox.showerror("Error", "No Make Selected!")
            return

        root.selected_result = final
        root.destroy()

    tk.Button(root,
              text="RUN EXECUTION",
              bg="#2980b9",
              fg="white",
              width=25,
              height=2,
              command=run_selected).pack(pady=15, anchor="w", padx=20)

    root.mainloop()

    if hasattr(root, "selected_result"):
        return root.selected_result
    else:
        return []


# =====================================================
# ALL YOUR ORIGINAL FUNCTIONS (UNCHANGED)
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
        make = unquote(make)
        alias = unquote(alias)
        return year, make, alias
    except:
        return "NOT_CAPTURED", "NOT_CAPTURED", "NOT_CAPTURED"

def enter_iframe(driver, wait):
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "iframe")))
    driver.switch_to.frame(driver.find_element(By.TAG_NAME, "iframe"))

def check_ext_hotspot(driver, timeout=12):
    end = time.time() + timeout
    while time.time() < end:
        for el in driver.find_elements(By.CSS_SELECTOR, ".control-block.eyeButton"):
            if el.is_displayed() and el.is_enabled():
                return "PASS"
        time.sleep(0.5)
    return "FAIL"

def get_region_from_url(url):
    url = url.lower()
    if "-fr" in url: return "FR"
    if "-es" in url: return "ES"
    if "-us" in url: return "US"
    return "FR"

def normalize_text(text):
    normalized = unicodedata.normalize('NFKD', text)
    ascii_text = normalized.encode('ASCII', 'ignore').decode('ASCII')
    return ascii_text.strip().lower()

def normalize_color_name(name):
    return normalize_text(name).replace(" ", "-")

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
    row_status = ["Duplicate Exist" if counts[c] > 1 else "NO Duplicate" for c in normalized]
    return color_names, row_status

def check_interior_pano(year, make, alias, region):
    year_short = year[-2:]
    year_make = f"{year_short}{make.lower()}"

    if region == "US":
        base = "https://app.izmoemporio.com/resources/izmo-emporio-hybrid/templates/templateone/views/visualizer/interiorpano.html?dataPath=//s3-us-west-2.amazonaws.com/static.izmocars.com/toolkit/newgen/"
    else:
        base = "https://app.izmoemporio.com/resources/izmo-emporio-hybrid/templates/templateone/views/visualizer/interiorpano.html?dataPath=//s3-eu-west-1.amazonaws.com/staticeu.izmocars.com/toolkit/newgen/"

    pano_url = f"{base}{year}/{year_make}/{alias}/{alias}_panovr/&country={region}"

    try:
        r = requests.head(pano_url, timeout=10)
        return "PASS" if r.status_code == 200 else "FAIL"
    except:
        return "FAIL"

def check_blue_hotspot(year, make, alias, region):
    year_short = year[-2:]
    year_make = f"{year_short}{make.lower()}"

    base = (
        "https://s3-us-west-2.amazonaws.com/static.izmocars.com/toolkit/newgen/"
        if region == "US"
        else "https://s3-eu-west-1.amazonaws.com/staticeu.izmocars.com/toolkit/newgen/"
    )

    hotspot_url = f"{base}{year}/{year_make}/{alias}/{alias}_panovr/skin/hs.webp"

    try:
        r = requests.head(hotspot_url, timeout=10)
        return "PASS" if r.status_code == 200 else "FAIL"
    except:
        return "FAIL"

def check_colors_broken(model_name, year, make, alias, color_list, s3_results, region):
    base_url = (
        f"https://static-h2.izmocars.com/toolkit/newgen/{region}"
        if region == "US"
        else f"https://staticeu-h2.izmocars.com/toolkit/newgen/{region}"
    )

    year_short = year[-2:]
    year_make = f"{year_short}{make.lower()}"

    base_prefix = f"{base_url}/{year}/{year_make}/{alias}/{alias}_animations/_360color/_660x440"

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

execution_df = load_urls_from_excel()
selected_indices = select_urls_ui(execution_df)

if not selected_indices:
    print("❌ No Make Selected. Exiting.")
    exit()

results = []
s3_results = []

for idx in selected_indices:
    URL = execution_df.loc[idx]["URL"]
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
        color_list, row_duplicate_status = get_color_names_and_duplicate_status(driver)
        color_count = len(color_list)

        check_colors_broken(model_name, year, make, alias, color_list, s3_results, region)

        ext_hotspot_status = check_ext_hotspot(driver)
        interior_pano_status = check_interior_pano(year, make, alias, region)
        blue_hotspot_status = check_blue_hotspot(year, make, alias, region)

        for idx2, color in enumerate(color_list):
            results.append({
                "Year": year,
                "Make": make,
                "Model Name": model_name,
                "Alias": alias,
                "Model Selector Broken": model_selector_status,
                "Color Name": color,
                "Color Count": color_count,
                "Duplicate Colors": row_duplicate_status[idx2],
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

# VALIDATE
s3_df = pd.DataFrame(s3_results)
if not s3_df.empty:
    s3_df = asyncio.run(validate_s3_links_async(s3_df))

summary_df = s3_df.groupby("Model Name").agg(
    Total_URLs=("S3 URL", "count"),
    PASS_Count=("S3 Status", lambda x: (x == "PASS").sum()),
    BROKEN_Count=("S3 Status", lambda x: (x != "PASS").sum())
).reset_index()

summary_df["PASS %"] = round((summary_df["PASS_Count"] / summary_df["Total_URLs"]) * 100, 2)
summary_df["BROKEN %"] = round((summary_df["BROKEN_Count"] / summary_df["Total_URLs"]) * 100, 2)

output_file = f"Emporio_Final_Output_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    pd.DataFrame(results).to_excel(writer, sheet_name="Main Report", index=False)
    s3_df.to_excel(writer, sheet_name="Color Broken Status", index=False)
    summary_df.to_excel(writer, sheet_name="Model Summary", index=False)

print("✅ Excel Generated Successfully")

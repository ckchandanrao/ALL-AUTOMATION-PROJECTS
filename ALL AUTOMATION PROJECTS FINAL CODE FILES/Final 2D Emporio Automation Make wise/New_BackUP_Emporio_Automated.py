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
from selenium.common.exceptions import TimeoutException

INPUT_FILE = "Emporio_Input.xlsx"
PARALLEL_REQUESTS = 10
RETRIES = 3
TIMEOUT = 30

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
             anchor="w").pack(pady=20, anchor="w", padx=20)

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
                     padx=1,
                     pady=0,
                     sticky="w")

            selected_indices.append((var, idx))

            col_num += 1

            if col_num == 6:  # 8 per row
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

def open_url(driver, url):

    try:
        driver.set_page_load_timeout(300)
        driver.get(url)

        print(f"✅ Loaded: {url}")
        return "PASS"

    except TimeoutException:
        print(f"❌ LINK NOT LOADING: {url}")

        # stop page loading
        driver.execute_script("window.stop();")

        return "LINK NOT LOADING"



def get_base_trim(driver):
    try:
        element = WebDriverWait(driver, 12).until(
            EC.presence_of_element_located((
                By.XPATH,
                "/html/body/div/div[2]/div/ng-view/div/div/div[4]/div[1]/div/div/p[2]"
            ))
        )

        text = element.text.strip()
        return text if text else ""

    except Exception as e:
        print("Base Trim not found:", e)
        return ""



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

def check_model_loading(driver):
    try:
        elements = driver.find_elements(By.CLASS_NAME, "fancybox-skin")

        # If any visible fancybox-skin exists → FAIL
        for el in elements:
            if el.is_displayed():
                return "FAIL"

        return "PASS"

    except:
        return "PASS"

# def enter_iframe(driver, wait):
#     wait.until(EC.presence_of_element_located((By.TAG_NAME, "iframe")))
#     driver.switch_to.frame(driver.find_element(By.TAG_NAME, "iframe"))


def enter_iframe(driver, wait):

    try:
        iframe = WebDriverWait(driver, 40).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "iframe"))
        )

        driver.switch_to.frame(iframe)

    except TimeoutException:
        print("❌ Iframe not found on page")
        raise



#old logic
# def check_ext_hotspot(driver, timeout=15):
#     end = time.time() + timeout
#     while time.time() < end:
#         for el in driver.find_elements(By.CSS_SELECTOR, ".control-block.eyeButton"):
#             if el.is_displayed() and el.is_enabled():
#                 return "PASS"
#         time.sleep(0.5)
#     return "FAIL"


#New Logic
def check_ext_hotspot(driver, timeout=15):

    end = time.time() + timeout

    while time.time() < end:

        try:

            elements = driver.find_elements(By.CSS_SELECTOR, ".control-block.eyeButton")

            for el in elements:

                visible = driver.execute_script("""
                    const el = arguments[0];
                    if (!el) return false;
                    const style = window.getComputedStyle(el);
                    return style && style.display !== 'none' && style.visibility !== 'hidden' && el.offsetHeight > 0;
                """, el)

                if visible:
                    return "PASS"

        except:
            pass

        time.sleep(0.5)

    return "FAIL"

#Getting region from emporio URL
# def get_region_from_url(url):
#     url = url.lower()
#     if "-fr" in url: return "FR"
#     if "-es" in url: return "ES"
#     if "-us" in url: return "US"
#     return "FR"

def normalize_text(text):
    normalized = unicodedata.normalize('NFKD', text)
    ascii_text = normalized.encode('ASCII', 'ignore').decode('ASCII')
    return ascii_text.strip().lower()

def normalize_for_s3(text):
    if not text:
        return ""

    # Remove accents (Š → S, é → e, etc.)
    normalized = unicodedata.normalize('NFKD', text)
    ascii_text = normalized.encode('ASCII', 'ignore').decode('ASCII')

    # Lowercase
    ascii_text = ascii_text.lower()

    # Replace special characters with space
    ascii_text = re.sub(r'[^a-z0-9\s-]', ' ', ascii_text)

    # Remove multiple spaces
    ascii_text = re.sub(r'\s+', ' ', ascii_text)

    # Replace spaces with hyphen
    ascii_text = ascii_text.strip().replace(' ', '-')

    # Remove double hyphens if any
    ascii_text = re.sub(r'-+', '-', ascii_text)

    return ascii_text


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


#OLD WORKING LOGIC
# def check_interior_pano(year, make, alias, region):
#     year_short = year[-2:]
#     year_make = f"{year_short}{make.lower()}"

#     if region == "US":
#         base = "https://app.izmoemporio.com/resources/izmo-emporio-hybrid/templates/templateone/views/visualizer/interiorpano.html?dataPath=//s3-us-west-2.amazonaws.com/static.izmocars.com/toolkit/newgen/"
#     else:
#         base = "https://app.izmoemporio.com/resources/izmo-emporio-hybrid/templates/templateone/views/visualizer/interiorpano.html?dataPath=//s3-eu-west-1.amazonaws.com/staticeu.izmocars.com/toolkit/newgen/"

#     pano_url = f"{base}{year}/{year_make}/{alias}/{alias}_panovr/&country={region}"

#     try:
#          #OLD LOGIC
#         r = requests.get(pano_url, timeout=10)
#         return "PASS" if r.status_code == 200 else "FAIL"
#     except:
#         return "FAIL"


# NEW OLD LOGIC
# def check_interior_pano(year, make, alias, region):

#     # Normalize make and alias (if you're using normalization)
#     normalized_make = normalize_for_s3(make)
#     normalized_alias = normalize_for_s3(alias)

#     year_short = year[-2:]
#     year_make = f"{year_short}{normalized_make}"

#     if region == "US":
#         base = "https://s3-us-west-2.amazonaws.com/static.izmocars.com/toolkit/newgen/"
#     else:
#         base = "https://s3-eu-west-1.amazonaws.com/staticeu.izmocars.com/toolkit/newgen/"

#     # 🔥 Build XML URL directly
#     xml_url = f"{base}{year}/{year_make}/{normalized_alias}/{normalized_alias}_panovr/izmoPano.xml"

#     try:
#         r = requests.head(xml_url, timeout=10)

#         print("Interior XML URL:", xml_url)
#         print("Interior XML Status:", r.status_code)

#         if r.status_code == 200:
#             return "PASS"
#         else:
#             return "FAIL"

#     except Exception as e:
#         print("Interior Pano Exception:", e)
#         return "FAIL"


def check_interior_pano(year, make, alias, region):

    # Normalize values for S3 path
    normalized_make = normalize_for_s3(make)
    normalized_alias = normalize_for_s3(alias)

    year_short = year[-2:]
    year_make = f"{year_short}{normalized_make}"

    # Select correct S3 base based on region
    if region == "US":
        base = "https://s3-us-west-2.amazonaws.com/static.izmocars.com/toolkit/newgen/"
    else:
        base = "https://s3-eu-west-1.amazonaws.com/staticeu.izmocars.com/toolkit/newgen/"

    # Build pano base folder
    pano_base = f"{base}{year}/{year_make}/{normalized_alias}/{normalized_alias}_panovr/"

    # Possible pano XML filenames used in Emporio
    pano_files = [
        "izmoPano.xml",
        "pano.xml",
        "panovr.xml",
        "config.xml"
    ]

    for file in pano_files:

        xml_url = pano_base + file

        try:
            print("Checking Interior Pano:", xml_url)

            r = requests.head(xml_url, timeout=10)

            if r.status_code == 200:
                print("Interior Pano Found:", xml_url)
                return "PASS"

        except Exception as e:
            print("Error checking pano:", e)

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

#OLD LOGIC
# def check_colors_broken(model_name, year, make, alias, color_list, s3_results, region):
#     base_url = (
#         f"https://static-h2.izmocars.com/toolkit/newgen/{region}"
#         if region == "US"
#         else f"https://staticeu-h2.izmocars.com/toolkit/newgen/{region}"
#     )

#     year_short = year[-2:]
#     #year_make = f"{year_short}{make.lower()}"
#     normalized_make = normalize_for_s3(make)
#     normalized_alias = normalize_for_s3(alias)
#     year_make = f"{year_short}{normalized_make}"

#     #base_prefix = f"{base_url}/{year}/{year_make}/{alias}/{alias}_animations/_360color/_660x440"
#     base_prefix = f"{base_url}/{year}/{year_make}/{normalized_alias}/{normalized_alias}_animations/_360color/_660x440"

#     for color_name in color_list:
#         #color_folder = normalize_color_name(color_name)
#         color_folder = normalize_for_s3(color_name)
#         for angle in range(1, 37):
#             angle_url = f"{base_prefix}/{color_folder}/angle{angle}.webp"
#             s3_results.append({
#                 "Model Name": model_name,
#                 "Year": year,
#                 "Make": make,
#                 "Alias": alias,
#                 "Color Name": normalize_for_s3(color_name),
#                 "Angle": angle,
#                 "S3 URL": angle_url
#             })

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

#OLD last working LOGIC
# async def check_url(session, url):
#     for attempt in range(RETRIES):
#         try:
#             async with session.get(url) as resp:
#                 return "PASS" if resp.status == 200 else f"BROKEN ({resp.status})"
#         except:
#             if attempt == RETRIES - 1:
#                 return "BROKEN"


#NEW LOGIC
async def check_url(session, url):

    headers = {
        "User-Agent": "Mozilla/5.0",
        "Accept": "image/webp,image/*,*/*;q=0.8"
    }

    for attempt in range(RETRIES):

        try:
            async with session.get(
                url,
                headers=headers,
                allow_redirects=True
            ) as resp:

                #await asyncio.sleep(0.03)

                if resp.status == 404:
                    return "FAIL"

                if resp.status in [200, 206]:
                    return "PASS"

                if resp.status in [403, 429, 500]:
                    await asyncio.sleep(1)
                    continue

        except:
            await asyncio.sleep(1)

    return "PASS"



#New old Logic
# async def check_url(session, row):

#     urls_to_try = [row["S3 URL"], row["Alt URL"]]

#     for url in urls_to_try:

#         try:
#             async with session.get(url, timeout=TIMEOUT) as resp:

#                 if resp.status in [200, 206]:
#                     return "PASS"

#         except:
#             pass

#     return "BROKEN"


#OLD LOGIC
# async def check_url(session, url):
    
#     headers = {
#         "User-Agent": "Mozilla/5.0",
#         "Accept": "image/webp,image/*,*/*;q=0.8"
#     }

#     for attempt in range(RETRIES):
#         try:
#             async with session.get(url, headers=headers, allow_redirects=True) as resp:

#                 if resp.status in [200, 206]:
#                     return "PASS"
#                 else:
#                     return f"BROKEN ({resp.status})"

#         except:
#             if attempt == RETRIES - 1:
#                 return "BROKEN"


# #old LOGIC
# async def check_url(session, url):

#     for attempt in range(RETRIES):

#         try:
#             async with session.get(url, allow_redirects=True) as resp:

#                 if resp.status in [200, 206]:
#                     return "PASS"
#                 else:
#                     return f"BROKEN ({resp.status})"

#         except Exception:

#             if attempt == RETRIES - 1:
#                 return "BROKEN"

#OLD LOGIC
# async def validate_s3_links_async(s3_df):
#     connector = aiohttp.TCPConnector(limit=PARALLEL_REQUESTS)
#     async with aiohttp.ClientSession(connector=connector) as session:
#         tasks = [check_url(session, url) for url in s3_df["S3 URL"].tolist()]
#         statuses = await asyncio.gather(*tasks)
#     s3_df["S3 Status"] = statuses
#     return s3_df

#NEW LOGIC
async def validate_s3_links_async(s3_df):

    connector = aiohttp.TCPConnector(
        limit=PARALLEL_REQUESTS,
        ttl_dns_cache=300,
        keepalive_timeout=30
    )

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
    #region = get_region_from_url(URL)  
    region = execution_df.loc[idx]["Region"]

    print(f"\n🌍 Processing URL: {URL}")

    driver = webdriver.Chrome()
    driver.set_page_load_timeout(300)
    wait = WebDriverWait(driver, 25)

    #driver.get(URL)

    status = open_url(driver, URL)

    if status == "LINK NOT LOADING":
        results.append({
            "URL": URL,
            "Status": "LINK NOT LOADING"
        })
        driver.quit() 
        continue


    driver.maximize_window()

    enter_iframe(driver, wait)
    time.sleep(3)

    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "h4.modelTitle")))
    model_count = len(driver.find_elements(By.CSS_SELECTOR, "h4.modelTitle"))

    # for i in range(model_count):
    #     model_titles = driver.find_elements(By.CSS_SELECTOR, "h4.modelTitle")
    #     model_name = model_titles[i].text.strip()
    #     model_selector_status = check_model_selector(driver, i)

    #     #CLICK MODEL
    #     driver.execute_script(
    #         "arguments[0].scrollIntoView({block:'center'}); arguments[0].click();",
    #         model_titles[i]
    #     )

    #     time.sleep(3)  # allow overlay if appears

    #     # 🔽 STEP 1: CHECK MODEL LOADING FIRST
    #     model_loading_status = check_model_loading(driver)
    #     # time.sleep(3)

    #     # 🔽 STEP 2: IF LOADING FAIL → SKIP BASE TRIM
    #     if model_loading_status == "FAIL":
    #         base_trim = ""
    #         print("⚠ Model Loading Failed - Skipping Base Trim")
    #     else:
    #         try:
    #             base_trim = get_base_trim(driver)
    #         except:
    #             base_trim=""
    #     print("DEBUG Base Trim:", base_trim)

    #     # base_trim = get_base_trim(driver)
    #     # print("DEBUG Base Trim:", base_trim)
    #     # Wait until header text is visible and not empty


    #     # WebDriverWait(driver, 10).until(
    #     #     lambda d: d.find_element(
    #     #         By.XPATH,
    #     #         "/html/body/div/div[2]/div/ng-view/div/div/div[4]/div[1]/div/div/p[2]"
    #     #     ).text.strip() != ""
    #     # )

    #     # base_trim = get_base_trim(driver)
    #     # print("DEBUG Base Trim:", base_trim)



    #     inside_s3 = ""
    #     try:
    #         img2 = WebDriverWait(driver, 6).until(
    #             EC.presence_of_element_located((By.CSS_SELECTOR, "img[src*='toolkit']"))
    #         )
    #         inside_s3 = img2.get_attribute("src")
    #     except:
    #         pass

    #     year, make, alias = extract_year_make_alias_from_s3(inside_s3)
    #     base_trim = get_base_trim(driver)

    #     color_list, row_duplicate_status = get_color_names_and_duplicate_status(driver)
    #     color_count = len(color_list)

    #     check_colors_broken(model_name, year, make, alias, color_list, s3_results, region)

    #     ext_hotspot_status = check_ext_hotspot(driver)
    #     interior_pano_status = check_interior_pano(year, make, alias, region)
    #     blue_hotspot_status = check_blue_hotspot(year, make, alias, region)

    #     for idx2, color in enumerate(color_list):
    #         results.append({
    #             "Year": year,
    #             "Make": make,
    #             "Model Name": model_name,
    #             "Alias": alias,
    #             "Base Trim": base_trim,
    #             "Model Selector Broken": model_selector_status,
    #             "Model Loading": model_loading_status,
    #             "Color Name": color,
    #             "Color Count": color_count,
    #             "Duplicate Colors": row_duplicate_status[idx2],
    #             "Ext Hotspot": ext_hotspot_status,
    #             "Interior Pano": interior_pano_status,
    #             "Blue Hotspot": blue_hotspot_status,
    #             "Inside S3 URL": inside_s3
    #         })

    #     driver.switch_to.default_content()
    #     driver.back()
    #     time.sleep(2)
    #     enter_iframe(driver, wait)


    for i in range(model_count):

        
        model_titles = driver.find_elements(By.CSS_SELECTOR, "h4.modelTitle")
        model_name = model_titles[i].text.strip()
        model_selector_status = check_model_selector(driver, i)

        # 🔽 CLICK MODEL
        driver.execute_script(
            "arguments[0].scrollIntoView({block:'center'}); arguments[0].click();",
            model_titles[i]
        )

        # Extra wait for first model only
        if i == 0:
            time.sleep(5)   # increase to 6–8 if needed
        else:
            time.sleep(2)

        time.sleep(2)

        # 🔽 CHECK MODEL LOADING FIRST
        model_loading_status = check_model_loading(driver)

        # 🔽 DEFAULT VALUES
        base_trim = ""
        color_list = []
        row_duplicate_status = []
        color_count = 0
        inside_s3 = ""
        year = make = alias = "NOT_CAPTURED"
        ext_hotspot_status = "FAIL"
        interior_pano_status = "FAIL"
        blue_hotspot_status = "FAIL"

        # 🔽 ONLY IF MODEL LOADED PROPERLY
        if model_loading_status == "PASS":

            try:
                base_trim = get_base_trim(driver)
            except:
                base_trim = ""

            try:
                img2 = WebDriverWait(driver, 6).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "img[src*='toolkit']"))
                )
                inside_s3 = img2.get_attribute("src")
                year, make, alias = extract_year_make_alias_from_s3(inside_s3)
            except:
                pass

            color_list, row_duplicate_status = get_color_names_and_duplicate_status(driver)
            color_count = len(color_list)

            check_colors_broken(model_name, year, make, alias, color_list, s3_results, region)
            
            

            ext_hotspot_status = check_ext_hotspot(driver)
            interior_pano_status = check_interior_pano(year, make, alias, region)
            blue_hotspot_status = check_blue_hotspot(year, make, alias, region)

        # 🔽 WRITE RESULT EVEN IF FAIL
        if not color_list:
            results.append({
                "Year": year,
                "Make": make,
                "Model Name": model_name,
                "Alias": alias,
                "Base Trim": base_trim,
                "Model Selector Broken": model_selector_status,
                "Model Loading": model_loading_status,
                "Color Name": "",
                "Color Count": 0,
                "Duplicate Colors": "",
                "Ext Hotspot": ext_hotspot_status,
                "Interior Pano": interior_pano_status,
                "Blue Hotspot": blue_hotspot_status,
                "Inside S3 URL": inside_s3
            })
        else:
            for idx2, color in enumerate(color_list):
                results.append({
                    "Year": year,
                    "Make": make,
                    "Model Name": model_name,
                    "Alias": alias,
                    "Base Trim": base_trim,
                    "Model Selector Broken": model_selector_status,
                    "Model Loading": model_loading_status,
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
    time.sleep(2)
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

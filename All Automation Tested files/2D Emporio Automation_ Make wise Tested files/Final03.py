# =====================================================
# SELENIUM ONLY — IZMO EMPORIO (FINAL – LOCKED LOGIC)
# =====================================================

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import re
import time
import os
from datetime import datetime

URL = "https://mercedes-benz.izmocars.fr/emporio-fr-fr.htm"

# =====================================================
# HELPERS (UNCHANGED)
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
        return year, make, alias
    except:
        return "NOT_CAPTURED", "NOT_CAPTURED", "NOT_CAPTURED"

def enter_iframe(driver, wait):
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "iframe")))
    driver.switch_to.frame(driver.find_element(By.TAG_NAME, "iframe"))

def check_ext_hotspot(driver, timeout=7):
    end = time.time() + timeout
    while time.time() < end:
        for el in driver.find_elements(By.CSS_SELECTOR, ".control-block.eyeButton"):
            if el.is_displayed() and el.is_enabled():
                return "PASS"
        time.sleep(0.3)
    return "FAIL"

def is_interior_tab_available(driver):
    return len(driver.find_elements(
        By.XPATH, "//*[@ng-click=\"maintainOpen('#INT360');\"]"
    )) > 0

def capture_screenshot(driver, folder, model_index, model_name, reason):
    os.makedirs(folder, exist_ok=True)
    safe = re.sub(r"[^\w\-]", "_", model_name)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    path = f"{folder}/{model_index}_{safe}_{reason}_{ts}.png"
    driver.save_screenshot(path)
    print(f"📸 Screenshot saved → {path}")

# =====================================================
# COLOR UNIQUENESS — IMAGE BASED (UNCHANGED)
# =====================================================
def check_color_uniqueness(driver):
    color_elements = driver.find_elements(
        By.CSS_SELECTOR, "#color-container .color-container.ng-scope"
    )

    image_ids = []
    for el in color_elements:
        try:
            img = el.find_element(By.CSS_SELECTOR, "img")
            src = img.get_attribute("src")
            if src:
                image_ids.append(src.split("?")[0])
        except:
            pass

    if len(image_ids) == len(set(image_ids)) and len(image_ids) > 0:
        return "PASS"
    return "FAIL"

# =====================================================
# COLOR NAMES + DUPLICATE CHECK (UNCHANGED)
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
                    lambda d: (
                        d.find_element(
                            By.CSS_SELECTOR,
                            ".fontFamilyBold.fontSize16.ng-binding.ng-scope"
                        ).text.strip() != last_text
                    )
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

    normalized = [c.lower() for c in color_names]

    duplicate_status = (
        "PASS"
        if len(normalized) > 0 and len(normalized) == len(set(normalized))
        else "FAIL"
    )

    return color_names, ", ".join(color_names), duplicate_status

# =====================================================
# COLORS BROKEN — S3 PATH CHECK (UNCHANGED)
# =====================================================
def check_colors_broken(driver, color_names):
    broken_colors = []

    try:
        color_tiles = driver.find_elements(
            By.XPATH,
            "//*[@id='color-container']//*[@ng-click=\"selectColorTlc(color,color.colorName);\"]"
        )

        for idx, tile in enumerate(color_tiles):
            try:
                driver.execute_script("arguments[0].click();", tile)
                time.sleep(0.6)

                img = driver.find_elements(By.CSS_SELECTOR, "img[src*='toolkit']")
                if not img or not img[0].get_attribute("src"):
                    broken_colors.append(f"{color_names[idx]} (BROKEN)")
            except:
                broken_colors.append(f"{color_names[idx]} (BROKEN)")
    except:
        pass

    return "PASS" if not broken_colors else ", ".join(broken_colors)

# =====================================================
# SETUP
# =====================================================
driver = webdriver.Chrome()
wait = WebDriverWait(driver, 25)

driver.get(URL)
driver.maximize_window()

results = []

# =====================================================
# INITIAL IFRAME
# =====================================================
enter_iframe(driver, wait)
time.sleep(3)

wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "h4.modelTitle")))
model_count = len(driver.find_elements(By.CSS_SELECTOR, "h4.modelTitle"))

print(f"🚗 Models found: {model_count}")

# =====================================================
# MAIN LOOP
# =====================================================
for i in range(model_count):
    print(f"\n👉 Processing model index {i}")

    model_titles = driver.find_elements(By.CSS_SELECTOR, "h4.modelTitle")
    model_name = model_titles[i].text.strip()

    outside_s3 = ""
    model_selector_status = "Fail"

    try:
        img = driver.find_elements(By.CSS_SELECTOR, "img")[i]
        outside_s3 = img.get_attribute("src") or ""
        if ".webp" in outside_s3.lower():
            model_selector_status = "Pass"
        else:
            capture_screenshot(driver, "screenshots/ModelSelector_FAIL", i, model_name, "MODEL_SELECTOR")
    except:
        capture_screenshot(driver, "screenshots/ModelSelector_FAIL", i, model_name, "MODEL_SELECTOR")

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

    try:
        WebDriverWait(driver, 8).until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, "#color-container .color-container.ng-scope")
            )
        )
        color_count = len(driver.find_elements(By.CSS_SELECTOR, "#color-container .color-container.ng-scope"))
    except:
        color_count = 0

    color_uniqueness_status = check_color_uniqueness(driver) if color_count else "FAIL"

    color_list, colors_name, duplicate_colors_status = get_color_names_and_duplicate_status(driver)
    colors_broken_status = check_colors_broken(driver, color_list)

    ext_hotspot_status = check_ext_hotspot(driver)
    if ext_hotspot_status == "FAIL":
        capture_screenshot(driver, "screenshots/ExtHotspot_FAIL", i, model_name, "EXT_HOTSPOT")

    if is_interior_tab_available(driver):
        interior_tab_status = "PASS"
        int_hotspot_status = "PASS"
    else:
        interior_tab_status = "NA"
        int_hotspot_status = "FAIL"
        capture_screenshot(driver, "screenshots/InteriorTab_FAIL", i, model_name, "INTERIOR_TAB")
        capture_screenshot(driver, "screenshots/InteriorHotspot_FAIL", i, model_name, "INTERIOR_HOTSPOT")

    results.append({
        "Model Name": model_name,
        "Model Selector": model_selector_status,
        "Outside S3 URL": outside_s3,
        "Year": year,
        "Make": make,
        "Alias": alias,
        "Color Count": color_count,
        "Color Uniqueness (Visual)": color_uniqueness_status,
        "Colors Name": colors_name,
        "Duplicate Colors": duplicate_colors_status,
        "Colors Broken": colors_broken_status,
        "Ext Hotspot": ext_hotspot_status,
        "Interior Tab": interior_tab_status,
        "Int Hotspot": int_hotspot_status,
        "Inside S3 URL": inside_s3
    })

    driver.switch_to.default_content()
    driver.back()
    time.sleep(2)
    enter_iframe(driver, wait)

driver.quit()

# =====================================================
# EXPORT
# =====================================================
df = pd.DataFrame(results)
output_file = f"Emporio_Final_Output_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
df.to_excel(output_file, index=False)

print(f"\n✅ FINAL EXCEL GENERATED → {output_file}")
print(df)

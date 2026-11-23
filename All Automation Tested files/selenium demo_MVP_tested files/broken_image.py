from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

import time
import csv
import pandas as pd
import requests

URL = "https://lexusmvp.stg-eudemosite.com/versions/lexus-es-es.htm"


def find_model_name_for_image(driver, img):
    """
    Try to find a model name near the image by looking for headings
    in ancestor <div> containers. Fallback to alt/src if nothing found.
    """
    # 1) Try heading inside nearest ancestors
    for depth in range(1, 6):  # up to 5 levels up
        try:
            ancestor = img.find_element(By.XPATH, f"./ancestor::div[{depth}]")
        except Exception:
            continue

        try:
            heading = ancestor.find_element(
                By.CSS_SELECTOR,
                "h1, h2, h3, h4, .version-name, .model-name, .grade-name, .version-title"
            )
            text = heading.text.strip()
            if text:
                return text
        except Exception:
            pass

    # 2) Fallback: use alt attribute if present
    alt = img.get_attribute("alt") or ""
    if alt.strip():
        return alt.strip()

    # 3) Fallback: derive something from src
    src = img.get_attribute("src") or ""
    if src:
        last_part = src.split("/")[-1]
        name_part = last_part.split(".")[0]
        return name_part

    return "UNKNOWN"


def image_status(driver, img, src):
    """
    Decide if image is loading or broken using:
      - JS naturalWidth after scrolling into view
      - HTTP status code
    """
    js_loaded = False
    http_ok = False

    # --- JS check with scrollIntoView + small wait ---
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", img)
        time.sleep(0.8)  # give browser time to actually load lazy images

        js_loaded = driver.execute_script("""
            var img = arguments[0];
            if (!img) return false;
            return img.complete &&
                   typeof img.naturalWidth != 'undefined' &&
                   img.naturalWidth > 0;
        """, img)
    except Exception:
        js_loaded = False

    # --- HTTP status check ---
    try:
        # HEAD is lighter; if some servers block HEAD, use GET instead
        resp = requests.head(src, timeout=5)
        if resp.status_code == 405:  # method not allowed, try GET
            resp = requests.get(src, stream=True, timeout=5)
        http_ok = (200 <= resp.status_code < 300)
    except Exception:
        http_ok = False

    if js_loaded or http_ok:
        return "loading"
    else:
        return "broken"


def main():
    driver = webdriver.Chrome()
    driver.maximize_window()
    driver.get(URL)

    wait = WebDriverWait(driver, 25)

    # Wait until at least one staticeu image appears
    try:
        wait.until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, "img[src*='staticeu.izmocars.com']")
            )
        )
    except TimeoutException:
        print("No images from staticeu.izmocars.com found. Check the page or selectors.")
        driver.quit()
        return

    all_imgs = driver.find_elements(By.TAG_NAME, "img")
    results = []

    for img in all_imgs:
        src = img.get_attribute("src") or ""
        if not src:
            continue

        # Only consider car images from staticeu.izmocars.com
        if "staticeu.izmocars.com" not in src:
            continue

        model_name = find_model_name_for_image(driver, img)
        status = image_status(driver, img, src)

        results.append({
            "Model Name": model_name,
            "Outside Image URL": src,
            "Status": status
        })

    driver.quit()

    if not results:
        print("No matching images collected. You may need to adjust the domain filter.")
        return

    # ---------- Print table ----------
    print("\nTABLE OUTPUT:\n")
    header = "| {:<30} | {:<80} | {:<8} |".format(
        "Model Name", "Outside Image URL", "Status"
    )
    sep = "|" + "-"*32 + "|" + "-"*82 + "|" + "-"*10 + "|"
    print(header)
    print(sep)

    for row in results:
        print(
            "| {:<30} | {:<80} | {:<8} |".format(
                (row["Model Name"] or "")[:30],
                (row["Outside Image URL"] or "")[:80],
                row["Status"]
            )
        )

    # ---------- Save CSV ----------
    csv_file = "lexus_versions_image_status.csv"
    with open(csv_file, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["Model Name", "Outside Image URL", "Status"])
        for row in results:
            writer.writerow([row["Model Name"], row["Outside Image URL"], row["Status"]])

    print(f"\nSaved CSV file: {csv_file}")

    # ---------- Save Excel ----------
    excel_file = "lexus_versions_image_status.xlsx"
    df = pd.DataFrame(results, columns=["Model Name", "Outside Image URL", "Status"])
    df.to_excel(excel_file, index=False)
    print(f"Saved Excel file: {excel_file}")


if __name__ == "__main__":
    main()

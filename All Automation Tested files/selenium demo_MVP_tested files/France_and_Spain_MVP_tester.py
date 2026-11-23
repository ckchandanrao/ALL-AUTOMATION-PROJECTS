import threading
import time
import json
from pathlib import Path
from urllib.parse import urlparse

import pandas as pd
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

import tkinter as tk
from tkinter import messagebox


# ============================================================
#          LOAD MAKES FROM EXTERNAL JSON FILES (FR + ES)
# ============================================================

MAKES_FILE_FR = Path("makes_FR.json")          # FRANCE MVP config
MAKES_FILE_ES = Path("makes_spain.json")    # SPAIN MVP config


def load_make_urls_from_file(path: Path):
    """Load make → URL mappings from JSON, keeping original case."""
    with path.open("r", encoding="utf-8") as f:
        raw = json.load(f)

    makes = {}
    for name, url in raw.items():
        if isinstance(url, str) and url.lower().startswith(("http://", "https://")):
            makes[name.strip()] = url
    return makes


def load_grouped_make_urls():
    """Force order: FRANCE MVP first, SPAIN MVP second."""
    grouped = {}

    if MAKES_FILE_FR.exists():
        grouped["FRANCE MVP"] = load_make_urls_from_file(MAKES_FILE_FR)

    if MAKES_FILE_ES.exists():
        grouped["SPAIN MVP"] = load_make_urls_from_file(MAKES_FILE_ES)

    if not grouped:
        raise FileNotFoundError("No valid JSON files found (makes.json / makes_spain.json).")

    return grouped


GROUPED_MAKE_URLS = load_grouped_make_urls()

MAKE_URLS = {}
for market, makes in GROUPED_MAKE_URLS.items():
    for make_name, url in makes.items():
        MAKE_URLS[f"{market}::{make_name}"] = url


# ============================================================
#                        COOKIE HANDLING
# ============================================================

COOKIES_DIR = Path("cookies")
COOKIES_DIR.mkdir(exist_ok=True)


def _cookie_file_for_make(composite_key: str):
    safe_key = composite_key.replace("::", "__")
    return COOKIES_DIR / f"cookies_{safe_key.lower()}.json"


def save_cookies(driver, composite_key):
    file_path = _cookie_file_for_make(composite_key)
    cookies = driver.get_cookies()
    file_path.write_text(json.dumps(cookies, indent=2))
    print(f"Cookies saved for {composite_key} → {file_path}")


def load_cookies(driver, url, composite_key):
    file = _cookie_file_for_make(composite_key)
    if not file.exists():
        print(f"No cookies for {composite_key}, first run.")
        driver.get(url)
        return False

    try:
        cookies = json.loads(file.read_text())
    except Exception as e:
        print(f"Failed to read cookies for {composite_key}: {e}")
        driver.get(url)
        return False

    driver.get(url)
    for c in cookies:
        c.pop("sameSite", None)
        try:
            driver.add_cookie(c)
        except Exception:
            pass

    driver.refresh()
    print(f"Cookies loaded for {composite_key}")
    return True


# ============================================================
#                       MODEL EXPANDERS
# ============================================================

def expand_all_models_on_page(driver):
    while True:
        try:
            btn = driver.find_element(
                By.XPATH,
                "//button[contains(translate(.,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'voir plus') "
                "or contains(translate(.,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'load more')]"
            )
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
            driver.execute_script("arguments[0].click();", btn)
            print("Clicked 'Voir plus' / 'Load more'")
            time.sleep(2)
        except Exception:
            break


def extract_year_alias_from_img(img_src: str):
    year, alias = "", ""
    if not img_src:
        return year, alias

    path = urlparse(img_src).path
    segments = [s for s in path.split("/") if s]

    for seg in segments:
        if seg.isdigit() and len(seg) == 4:
            year = seg
            break

    filename = segments[-1]
    base = filename.split(".", 1)[0]
    parts = base.split("_")
    if len(parts) >= 2:
        alias = parts[1]

    return year, alias


# ============================================================
#               MAIN SCRAPER LOGIC
# ============================================================

def run_test_for_make(composite_key, status_callback=None):

    def update(msg):
        print(msg)
        if status_callback:
            status_callback(msg)

    market, make_name = composite_key.split("::", 1)
    url = MAKE_URLS[composite_key]

    update(f"[{market}] Starting {make_name.upper()} → {url}")

    options = Options()
    options.add_experimental_option("detach", True)
    driver = webdriver.Chrome(options=options)
    driver.maximize_window()

    cookies_loaded = load_cookies(driver, url, composite_key)
    if not cookies_loaded:
        update(f"[{market}] First load – accept cookies manually.")
        time.sleep(8)

    wait = WebDriverWait(driver, 10)
    rows = []

    time.sleep(2)
    nums = []
    for e in driver.find_elements(By.CSS_SELECTOR, ".paginationPageNumber"):
        t = e.text.strip()
        if t.isdigit():
            nums.append(int(t))
    pages = sorted(nums) if nums else [1]

    update(f"[{market}] Pages detected: {pages}")

    for p in pages:
        if p != 1:
            try:
                link = wait.until(EC.element_to_be_clickable(
                    (By.XPATH, f"//*[contains(@class,'paginationPageNumber')][text()='{p}']")
                ))
                driver.execute_script("arguments[0].click();", link)
                time.sleep(2)
            except Exception:
                continue

        wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME, "model-name")))

        while True:
            last = driver.execute_script("return document.body.scrollHeight")
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(2)
            new = driver.execute_script("return document.body.scrollHeight")
            if new == last:
                break

        expand_all_models_on_page(driver)

        names = [e.text.strip() for e in driver.find_elements(By.CLASS_NAME, "model-name") if e.text.strip()]
        imgs = driver.find_elements(By.CSS_SELECTOR, "img.img-responsive.model-img-t")

        for idx, model in enumerate(names):
            src = imgs[idx].get_attribute("src") if idx < len(imgs) else ""
            if ".webp" in (src.lower() if src else ""):
                status = "Pass"
                year, alias = extract_year_alias_from_img(src)
            else:
                status = "Fail"
                year = alias = "Fail"

            rows.append({
                "Market": market,
                "Make": make_name.upper(),
                "Model Name": model,
                "Year": year,
                "Alias": alias,
                "Model selector": status,
                "Page": p
            })

            update(f"[{market}] {model} → {status}")

    save_cookies(driver, composite_key)
    driver.quit()
    update(f"[{market}] Finished. Rows collected: {len(rows)}")
    return rows


# ============================================================
# GUI (Start Test Button at TOP + Group Select-All)
# ============================================================

def main():
    root = tk.Tk()
    root.title("Model Image Tester (FRANCE MVP + SPAIN MVP)")
    root.geometry("900x950")

    # ------------------------------------------------------
    # START TEST BUTTON AT THE TOP
    # ------------------------------------------------------
    def on_start():
        selected = [key for key, v in make_vars.items() if v.get()]
        if not selected:
            messagebox.showwarning("Warning", "Select at least one make.")
            return

        def worker():
            all_rows = []
            total = len(selected)

            for i, comp in enumerate(selected, start=1):
                set_status(f"Running {i}/{total}: {comp}")
                rows = run_test_for_make(comp, set_status)
                all_rows.extend(rows)

            if all_rows:
                df = pd.DataFrame(all_rows)
                df.to_excel("models_all_makes_year_alias.xlsx", index=False)
                set_status("Done. Excel saved: models_all_makes_year_alias.xlsx")
            else:
                set_status("No data collected.")

        threading.Thread(target=worker, daemon=True).start()

    start_btn = tk.Button(root, text="START TEST", font=("Segoe UI", 12, "bold"), command=on_start, bg="#4CAF50", fg="white")
    start_btn.pack(pady=15)

    tk.Label(root, text="Select Makes to Test:", font=("Segoe UI", 12, "bold")).pack()

    # ------------------------------------------------------
    # MAIN CONTAINER
    # ------------------------------------------------------
    container = tk.Frame(root)
    container.pack(fill="both", expand=True, padx=15, pady=5)

    make_vars = {}
    group_keys = {"FRANCE MVP": [], "SPAIN MVP": []}

    ordered_groups = ["FRANCE MVP", "SPAIN MVP"]

    for market in ordered_groups:
        if market not in GROUPED_MAKE_URLS:
            continue

        lf = tk.LabelFrame(container, text=market, font=("Segoe UI", 11, "bold"), padx=10, pady=8)
        lf.pack(fill="x", pady=10)

        makes = GROUPED_MAKE_URLS[market]
        columns = 7

        # Select All button for each group
        select_all_var = tk.BooleanVar()

        def make_toggle(market_name, v=select_all_var):
            def _toggle():
                val = v.get()
                for key in group_keys[market_name]:
                    make_vars[key].set(val)
            return _toggle

        tk.Checkbutton(
            lf,
            text=f"Select all {market}",
            variable=select_all_var,
            command=make_toggle(market)
        ).grid(row=0, column=0, sticky="w", padx=5, pady=5)

        # Checkboxes
        for i, make_name in enumerate(sorted(makes.keys())):
            comp = f"{market}::{make_name}"
            var = tk.BooleanVar()
            make_vars[comp] = var
            group_keys[market].append(comp)

            tk.Checkbutton(lf, text=make_name, variable=var).grid(
                row=1 + i // columns,
                column=i % columns,
                sticky="w",
                padx=10,
                pady=3
            )

    # ------------------------------------------------------
    # STATUS LABEL AT BOTTOM
    # ------------------------------------------------------
    status = tk.StringVar(value="Idle")

    tk.Label(root, textvariable=status, anchor="w", justify="left", wraplength=820).pack(
        pady=10, fill="x", padx=10
    )

    def set_status(msg):
        status.set(msg)
        root.update_idletasks()

    root.mainloop()


if __name__ == "__main__":
    main()

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


# ---------- COOKIES SAVE / LOAD (MAIN LOGIC FOR REMEMBERING CONSENT) ----------

COOKIES_DIR = Path("cookies")
COOKIES_DIR.mkdir(exist_ok=True)


def _cookie_file_for_make(make: str) -> Path:
    return COOKIES_DIR / f"cookies_{make.lower()}.json"


def save_cookies(driver, make: str):
    """
    Save all cookies for this session into cookies/cookies_<make>.json
    Call this at the end of a successful run.
    """
    file_path = _cookie_file_for_make(make)
    cookies = driver.get_cookies()
    file_path.write_text(json.dumps(cookies, indent=2))
    print(f"Cookies saved for {make.upper()} → {file_path}")


def load_cookies(driver, url: str, make: str) -> bool:
    """
    If cookie file exists for this make, load cookies, refresh the page,
    and return True.
    If not, JUST open the page and return False, so you can accept manually.
    """
    file_path = _cookie_file_for_make(make)
    if not file_path.exists():
        print(f"No saved cookies for {make.upper()} yet → first run, manual accept needed.")
        driver.get(url)
        return False

    try:
        cookies = json.loads(file_path.read_text())
    except Exception as e:
        print(f"Failed to read cookies for {make.upper()}: {e}. Opening page without cookies.")
        driver.get(url)
        return False

    # Must open the domain before adding cookies
    driver.get(url)
    for c in cookies:
        # Some fields like 'sameSite' can break add_cookie in some setups
        c.pop("sameSite", None)
        try:
            driver.add_cookie(c)
        except Exception:
            # If any cookie can't be added, just skip it
            continue

    driver.refresh()
    print(f"Cookies loaded for {make.upper()} from {file_path}")
    return True


# ---------- COOKIE / POPUP HANDLER: ALWAYS ACCEPT COOKIES (BACKUP, UNUSED BY DEFAULT) ----------

def accept_cookies(driver, timeout: int = 5):
    """
    Backup cookie handler in case popup still appears.
    Priority:
    1) <span id="acceptSpan">TOUT ACCEPTER</span>   (Skoda-style)
    2) Any element with text 'TOUT ACCEPTER' / 'TOUT ACCEPTOR'
    3) Generic 'ACCEPTER' / 'ACCEPT' buttons (OneTrust, etc.)
    """
    wait = WebDriverWait(driver, timeout)

    # 1) Direct Skoda popup accept: <span id="acceptSpan">TOUT ACCEPTER</span>
    try:
        span = wait.until(EC.element_to_be_clickable((By.ID, "acceptSpan")))
        if span.is_displayed():
            driver.execute_script("arguments[0].click();", span)
            print("Cookies accepted → <span id='acceptSpan'>TOUT ACCEPTER</span>")
            return
    except Exception:
        pass

    # 2) Generic TOUT ACCEPTER / TOUT ACCEPTOR (buttons, links, spans, divs)
    tout_xpath = (
        # buttons
        "//button[contains(translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'), 'TOUT ACCEPTER') "
        "or contains(translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'), 'TOUT ACCEPTOR')]"
        # links
        " | //a[contains(translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'), 'TOUT ACCEPTER') "
        "or contains(translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'), 'TOUT ACCEPTOR')]"
        # spans/divs
        " | //*[(self::div or self::span) and ("
        "contains(translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'), 'TOUT ACCEPTER') "
        "or contains(translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'), 'TOUT ACCEPTOR'))]"
    )

    try:
        btn = wait.until(EC.element_to_be_clickable((By.XPATH, tout_xpath)))
        driver.execute_script("arguments[0].click();", btn)
        print("Cookies accepted → TOUT ACCEPTER / TOUT ACCEPTOR (generic)")
        return
    except Exception:
        pass

    # 3) Fallback generic accept buttons (OneTrust, Accept, Accepter, etc.)
    fallback_selectors = [
        "#onetrust-accept-btn-handler",
        ".onetrust-accept-btn-handler",
        "//button[contains(translate(., 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'),'ACCEPTER')]",
        "//button[contains(translate(., 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'),'ACCEPT')]",
    ]

    for sel in fallback_selectors:
        try:
            if sel.startswith("//"):
                el = driver.find_element(By.XPATH, sel)
            else:
                el = driver.find_element(By.CSS_SELECTOR, sel)

            if el.is_displayed():
                driver.execute_script("arguments[0].click();", el)
                print(f"Cookies accepted → Fallback clicked: {sel}")
                return
        except Exception:
            continue

    print("No cookie popup found or nothing clickable.")


# ---------- EXPAND "LOAD MORE" / "VOIR PLUS" BUTTONS ----------

def expand_all_models_on_page(driver):
    """
    Some makes (e.g. Skoda) only show a few models at first and
    load more when you click a 'Voir plus' / 'Plus de résultats' /
    'Afficher plus' / 'Load more' button.
    This function tries to click those buttons until none remain.
    """
    while True:
        try:
            # Case-insensitive French + English text matching
            btn = driver.find_element(
                By.XPATH,
                "//button[contains(translate(normalize-space(.),"
                "'ABCDEFGHIJKLMNOPQRSTUVWXYZÉÈÊÀÂÎÏÔÛÙÇ',"
                "'abcdefghijklmnopqrstuvwxyzéèêàâîïôûùç'),'voir plus')"
                " or contains(translate(normalize-space(.),"
                "'ABCDEFGHIJKLMNOPQRSTUVWXYZÉÈÊÀÂÎÏÔÛÙÇ',"
                "'abcdefghijklmnopqrstuvwxyzéèêàâîïôûùç'),'plus de résultats')"
                " or contains(translate(normalize-space(.),"
                "'ABCDEFGHIJKLMNOPQRSTUVWXYZÉÈÊÀÂÎÏÔÛÙÇ',"
                "'abcdefghijklmnopqrstuvwxyzéèêàâîïôûùç'),'afficher plus')"
                " or contains(translate(normalize-space(.),"
                "'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'load more')]"
            )

            if not btn.is_displayed():
                break

            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
            driver.execute_script("arguments[0].click();", btn)
            print("Clicked 'show more' / 'voir plus' button to load more models.")
            time.sleep(2)  # wait for extra models to load, then loop again

        except NoSuchElementException:
            # No more 'see more' buttons
            break
        except Exception:
            # Anything weird, just stop trying
            break


# ---------- CORE SCRAPING LOGIC ----------

def extract_year_alias_from_img(img_src: str):
    """
    Try to extract year and alias from an image URL like:
    https://.../FR/2025/25bmw/25bmw120dftnmsprthb4bfr/640x480/bmw_120dftnmsprthb4bfr_lowaggressive.webp

    Returns:
        year  -> '2025'
        alias -> '120dftnmsprthb4bfr'
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


def run_test_for_make(selected_make: str, status_callback=None):
    """
    Runs the scraping for a given make (lowercased).
    First run:
      - No cookies file → open page → YOU accept cookies manually.
      - At the end, cookies are saved.
    Next runs:
      - Cookies are loaded automatically → no manual click needed.
    Returns: list[dict] rows for this make.
    """

    def update_status(msg):
        print(msg)
        if status_callback:
            status_callback(msg)

    selected_make = selected_make.strip().lower()
    if selected_make not in MAKE_URLS:
        update_status(f"Invalid make: {selected_make}")
        return []

    start_url = MAKE_URLS[selected_make]
    update_status(f"Starting test for {selected_make.upper()} → {start_url}")

    # ---------- SELENIUM ----------
    options = Options()
    options.add_experimental_option("detach", True)

    driver = webdriver.Chrome(options=options)
    driver.maximize_window()

    # Try to load cookies (second run onwards)
    cookies_loaded = load_cookies(driver, start_url, selected_make)

    if not cookies_loaded:
        # First run case: page is opened, you need to click cookies manually.
        update_status(
            f"{selected_make.upper()} – First run: please accept cookies manually in the opened browser window."
        )
        # Give a few seconds for you to click, or you can just wait while script continues.
        time.sleep(8)

    wait = WebDriverWait(driver, 10)
    all_rows = []

    # ---------- DETECT PAGINATION PAGES (numbers 1,2,3,...) ----------
    time.sleep(2)
    pagination_elems = driver.find_elements(By.CSS_SELECTOR, ".paginationPageNumber")
    page_numbers = []

    for el in pagination_elems:
        txt = el.text.strip()
        if txt.isdigit():
            page_numbers.append(int(txt))

    if page_numbers:
        page_numbers = sorted(set(page_numbers))
    else:
        page_numbers = [1]  # no pagination, single page

    update_status(f"{selected_make.upper()} – detected pages: {page_numbers}")

    # ---------- LOOP OVER EACH PAGE NUMBER ----------
    for idx_page, page in enumerate(page_numbers):
        if idx_page > 0:
            # click page number (pagination) to navigate
            update_status(f"{selected_make.upper()} – going to page {page}...")
            try:
                page_xpath = (
                    "//*[contains(@class,'paginationPageNumber') and normalize-space(text())='%s']"
                    % page
                )
                page_el = wait.until(EC.element_to_be_clickable((By.XPATH, page_xpath)))
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", page_el)
                driver.execute_script("arguments[0].click();", page_el)
                time.sleep(2)
            except (TimeoutException, ElementNotInteractableException,
                    ElementClickInterceptedException, NoSuchElementException) as e:
                update_status(f"  Could not click page {page}: {e}. Skipping this page.")
                continue

        update_status(f"=== {selected_make.upper()} – Scraping Page {page} ===")

        # Wait for at least some models
        wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME, "model-name")))
        time.sleep(1)

        # Scroll to load lazy models
        last_height = driver.execute_script("return document.body.scrollHeight")
        while True:
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(2)
            new_height = driver.execute_script("return document.body.scrollHeight")
            if new_height == last_height:
                break
            last_height = new_height

        # Click any "voir plus" / "load more" buttons to reveal hidden models
        expand_all_models_on_page(driver)

        # Model names
        model_elems = driver.find_elements(By.CLASS_NAME, "model-name")
        model_names = [el.text.strip() for el in model_elems if el.text.strip()]
        update_status(f"{selected_make.upper()} – models found on page {page}: {len(model_names)}")

        # Images
        img_elems = driver.find_elements(By.CSS_SELECTOR, "img.img-responsive.model-img-t")
        update_status(f"{selected_make.upper()} – images found on page {page}: {len(img_elems)}")

        # Pair models + images
        for idx, name in enumerate(model_names):
            img_src = ""
            if idx < len(img_elems):
                img_src = img_elems[idx].get_attribute("src") or ""

            img_src_lower = img_src.lower() if img_src else ""

            # MODEL SELECTOR RULE:
            # .jpg or no image -> Fail
            # .webp -> Pass + extract year/alias
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
                model_selector_status = "Fail"
                year_out = "Fail"
                alias_out = "Fail"

            update_status(
                f"{selected_make.upper()} | page {page} | {idx+1}. {name} → {model_selector_status}"
            )

            all_rows.append({
                "Make": selected_make.upper(),
                "Model Name": name,
                "Year": year_out,
                "Alias": alias_out,
                "Model selector": model_selector_status,
                "Page": page
            })

    # ---------- SAVE COOKIES FOR NEXT RUN ----------
    save_cookies(driver, selected_make)

    driver.quit()
    update_status(f"{selected_make.upper()} – browser closed")
    update_status(f"{selected_make.upper()} done. Total rows: {len(all_rows)}")
    return all_rows


# ---------- TKINTER GUI WITH SIDE-BY-SIDE CHECKBOXES ----------

def main():
    root = tk.Tk()
    root.title("Model Image Tester (Checkbox Version)")
    root.geometry("700x500")

    tk.Label(root, text="Select Makes to Test:", font=("Segoe UI", 11, "bold")).pack(pady=(10, 5))

    # Frame to hold checkboxes in grid
    checkbox_frame = tk.Frame(root)
    checkbox_frame.pack(fill="both", expand=True, padx=10, pady=10)

    make_vars = {}
    make_list = sorted(MAKE_URLS.keys())

    # how many columns of checkboxes you want
    columns = 3

    for index, make in enumerate(make_list):
        var = tk.BooleanVar(value=False)
        cb = tk.Checkbutton(checkbox_frame, text=make, variable=var, anchor="w", justify="left")

        row = index // columns
        col = index % columns

        cb.grid(row=row, column=col, sticky="w", padx=15, pady=4)
        make_vars[make] = var

    # Select all / deselect all
    select_all_var = tk.BooleanVar(value=False)

    def on_select_all():
        val = select_all_var.get()
        for v in make_vars.values():
            v.set(val)

    select_all_cb = tk.Checkbutton(
        root,
        text="Select all",
        variable=select_all_var,
        command=on_select_all
    )
    select_all_cb.pack(pady=(0, 5))

    # Status label
    status_var = tk.StringVar(value="Idle")
    status_label = tk.Label(
        root,
        textvariable=status_var,
        anchor="w",
        justify="left",
        wraplength=660
    )
    status_label.pack(pady=(5, 5), fill="x", padx=10)

    def update_status(text):
        status_var.set(text)
        root.update_idletasks()

    def on_start():
        # Collect selected makes
        selected_makes = [m for m, v in make_vars.items() if v.get()]
        if not selected_makes:
            messagebox.showwarning("No make selected", "Please select at least one make.")
            return

        def worker():
            try:
                total = len(selected_makes)
                all_results = []

                for idx, make in enumerate(selected_makes, start=1):
                    update_status(f"Running {idx}/{total}: {make.upper()}...")
                    rows = run_test_for_make(make, status_callback=update_status)
                    all_results.extend(rows)

                if all_results:
                    df = pd.DataFrame(all_results)
                    excel_filename = "models_all_makes_year_alias.xlsx"
                    df.to_excel(excel_filename, index=False)
                    update_status(f"All selected makes completed. Excel saved: {excel_filename}")
                else:
                    update_status("No data collected. Nothing to save.")
            except Exception as e:
                update_status(f"Error: {e}")
                messagebox.showerror("Error", str(e))

        threading.Thread(target=worker, daemon=True).start()

    start_btn = tk.Button(root, text="Start Test", command=on_start, font=("Segoe UI", 10, "bold"))
    start_btn.pack(pady=(5, 10))

    root.mainloop()


if __name__ == "__main__":
    main()

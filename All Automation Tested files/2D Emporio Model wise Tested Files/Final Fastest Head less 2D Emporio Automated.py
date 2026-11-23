# =====================================================
# 🔥 EMPORIO AUTOMATION WITH EXCEL + UI (MULTI-MAKE PARALLEL VERSION)
# =====================================================
#
# WHAT CHANGED vs your original script:
#   1. When you select multiple Makes in the first screen, all of them
#      now load AT THE SAME TIME (parallel Chrome windows via ThreadPoolExecutor)
#      instead of one after another.
#   2. Instead of a separate "select models" popup per make, you now get
#      ONE combined window showing models from every selected make,
#      grouped under that make's name, with a single search box that
#      filters across all of them together.
#   3. Once you hit "RUN SELECTED", the script goes back to each make's
#      already-open driver and processes only the models you picked for
#      that make - same checks as before (colors, hotspots, pano, etc.)
#   4. Everything still lands in one combined Excel workbook at the end,
#      same 3 sheets: Main Report / Color Broken Status / Model Summary.
#
# Tune MAX_PARALLEL_MAKES below if opening too many Chrome windows at once
# is too heavy for your machine.
# =====================================================

from selenium import webdriver
from selenium.webdriver.chrome.options import Options as ChromeOptions
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
from urllib.parse import unquote, quote
from collections import Counter
import requests
import tkinter as tk
from tkinter import ttk
from selenium.common.exceptions import TimeoutException
from concurrent.futures import ThreadPoolExecutor, as_completed
import concurrent.futures
import queue

INPUT_FILE = r"C:\Users\chandan.rao\Desktop\Backoffice\2D Emporio Model wise\Emporio_Input.xlsx"
import os

print("Current Working Directory:", os.getcwd())
print("Excel Exists:", os.path.exists(INPUT_FILE))

PARALLEL_REQUESTS = 20  # concurrent open connections per batch during S3 link validation
RETRIES = 3
TIMEOUT = 30

# How many makes to load in parallel (i.e. how many Chrome windows open at once).
# Set this to len(selected_indices) if you want ALL selected makes to load
# simultaneously no matter how many there are, or cap it to be gentler on RAM/CPU.
MAX_PARALLEL_MAKES = 5

# How many makes to PROCESS (click through models, check colors/hotspots) at
# once, in parallel. Each make already has its own independent Chrome driver
# from the loading step, so this is safe - it turns the old "one make at a
# time" processing loop into "several makes at once", which is the biggest
# speed win in the whole script.
MAX_PARALLEL_PROCESSING = 5

# Set to True to run Chrome without any visible windows (recommended when
# loading several makes in parallel - avoids a wall of browser windows
# popping up on screen). Set to False if you ever need to visually watch
# what's happening for debugging.
HEADLESS = True


def get_chrome_options():
    options = ChromeOptions()

    if HEADLESS:
        options.add_argument("--headless=new")
        options.add_argument("--window-size=1920,1080")
        options.add_argument("--disable-gpu")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--log-level=3")

    return options


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

        except Exception as e:
            print(f"Error reading sheet {region}: {e}")
            continue

    if not all_data:
        raise Exception("No valid sheets (US, FR, ES) found in Excel.")

    return pd.concat(all_data).reset_index(drop=True)


# =====================================================
# UI SELECTION (Region Wise Select / Deselect) - unchanged
# =====================================================
def select_urls_ui(execution_df):

    # ---- Color palette ----
    COLOR_BG        = "#f4f6f9"   # page background
    COLOR_HEADER_BG = "#1f2a44"   # header strip
    COLOR_HEADER_FG = "#ffffff"
    COLOR_CARD_BG   = "#ffffff"
    COLOR_CARD_EDGE = "#dde3ec"
    COLOR_REGION_FG = "#1f2a44"
    COLOR_ACCENT    = "#2980b9"
    COLOR_GREEN     = "#27ae60"
    COLOR_RED       = "#e74c3c"
    COLOR_FOOTER_BG = "#ffffff"
    COLOR_MUTED     = "#6b7686"

    root = tk.Tk()
    root.title("Emporio Execution Dashboard")
    root.geometry("950x750")
    root.configure(bg=COLOR_BG)
    root.minsize(700, 500)

    # ---------- Header ----------
    header = tk.Frame(root, bg=COLOR_HEADER_BG, height=70)
    header.pack(fill="x", side="top")
    header.pack_propagate(False)

    tk.Label(
        header,
        text="EMPORIO EXECUTION DASHBOARD",
        font=("Segoe UI", 17, "bold"),
        bg=COLOR_HEADER_BG,
        fg=COLOR_HEADER_FG
    ).pack(side="left", padx=20)

    tk.Label(
        header,
        text="Select the makes you want to run, grouped by region",
        font=("Segoe UI", 10),
        bg=COLOR_HEADER_BG,
        fg="#c7cfe0"
    ).pack(side="left", padx=(0, 20))

    # ---------- Scrollable body ----------
    body_container = tk.Frame(root, bg=COLOR_BG)
    body_container.pack(fill="both", expand=True)

    canvas = tk.Canvas(body_container, bg=COLOR_BG, highlightthickness=0)
    scrollbar = ttk.Scrollbar(body_container, orient="vertical", command=canvas.yview)

    frame = tk.Frame(canvas, bg=COLOR_BG)

    frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )

    canvas.create_window((0, 0), window=frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side="left", fill="both", expand=True, padx=(15, 0), pady=15)
    scrollbar.pack(side="right", fill="y")

    # Mouse-wheel scrolling (Windows/Mac)
    def _on_mousewheel(event):
        canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    canvas.bind_all("<MouseWheel>", _on_mousewheel)

    selected = {}
    all_vars_flat = []

    footer_count_var = tk.StringVar(value="0 make(s) selected")

    def refresh_footer_count():
        total = sum(v.get() for v in all_vars_flat)
        footer_count_var.set(f"{total} make(s) selected")

    for region in execution_df["Region"].unique():

        region_df = execution_df[execution_df["Region"] == region]

        # ---- Region header row ----
        region_header = tk.Frame(frame, bg=COLOR_BG)
        region_header.pack(fill="x", pady=(18, 6), padx=5)

        tk.Label(
            region_header,
            text=f"● {region}",
            font=("Segoe UI", 13, "bold"),
            bg=COLOR_BG,
            fg=COLOR_REGION_FG
        ).pack(side="left")

        tk.Label(
            region_header,
            text=f"{len(region_df)} make(s)",
            font=("Segoe UI", 9),
            bg=COLOR_BG,
            fg=COLOR_MUTED
        ).pack(side="left", padx=(8, 0))

        # ---- Card holding the checkboxes ----
        card = tk.Frame(frame, bg=COLOR_CARD_BG, highlightbackground=COLOR_CARD_EDGE,
                         highlightthickness=1, bd=0)
        card.pack(fill="x", padx=5)

        chk_area = tk.Frame(card, bg=COLOR_CARD_BG)
        chk_area.pack(fill="x", padx=15, pady=(12, 4))

        region_vars = []
        row = 0
        col = 0

        for idx, rowdata in region_df.iterrows():

            var = tk.IntVar(value=1)
            var.trace_add("write", lambda *a: refresh_footer_count())

            selected[idx] = var
            region_vars.append(var)
            all_vars_flat.append(var)

            chk = tk.Checkbutton(
                chk_area,
                text=rowdata["MAKE"],
                variable=var,
                anchor="w",
                width=25,
                bg=COLOR_CARD_BG,
                activebackground=COLOR_CARD_BG,
                font=("Segoe UI", 10),
                selectcolor="#eaf3fb"
            )

            chk.grid(row=row, column=col, sticky="w", padx=10, pady=4)

            col += 1
            if col == 3:
                col = 0
                row += 1

        # ---- Select all / deselect all row for this card ----
        card_btn_row = tk.Frame(card, bg=COLOR_CARD_BG)
        card_btn_row.pack(fill="x", padx=15, pady=(4, 12))

        def select_all(lst=region_vars):
            for v in lst:
                v.set(1)

        def deselect_all(lst=region_vars):
            for v in lst:
                v.set(0)

        tk.Button(
            card_btn_row,
            text="Select All",
            bg=COLOR_GREEN,
            fg="white",
            font=("Segoe UI", 9, "bold"),
            relief="flat",
            padx=10,
            pady=3,
            command=select_all
        ).pack(side="left", padx=(0, 6))

        tk.Button(
            card_btn_row,
            text="Deselect All",
            bg=COLOR_RED,
            fg="white",
            font=("Segoe UI", 9, "bold"),
            relief="flat",
            padx=10,
            pady=3,
            command=deselect_all
        ).pack(side="left")

    refresh_footer_count()

    # ---------- Sticky footer ----------
    footer = tk.Frame(root, bg=COLOR_FOOTER_BG, height=64,
                       highlightbackground=COLOR_CARD_EDGE, highlightthickness=1)
    footer.pack(fill="x", side="bottom")
    footer.pack_propagate(False)

    tk.Label(
        footer,
        textvariable=footer_count_var,
        font=("Segoe UI", 10),
        bg=COLOR_FOOTER_BG,
        fg=COLOR_MUTED
    ).pack(side="left", padx=20)

    def run():

        root.selected = [
            idx
            for idx, var in selected.items()
            if var.get()
        ]

        root.destroy()

    tk.Button(
        footer,
        text="RUN EXECUTION",
        bg=COLOR_ACCENT,
        fg="white",
        font=("Segoe UI", 11, "bold"),
        relief="flat",
        width=22,
        height=1,
        pady=8,
        command=run
    ).pack(side="right", padx=20, pady=10)

    root.mainloop()

    if hasattr(root, "selected"):
        return root.selected

    return []


# =====================================================
# NEW: UNIFIED MULTI-MAKE MODEL SELECTION UI
# Shows models from every loaded make in one scrollable window,
# grouped under that make's heading, with ONE search box that
# filters across all of them at once.
# =====================================================
def select_models_ui_multi(loaded_makes, parent=None):
    """
    loaded_makes: list of dicts, each like:
        {
            "idx": <original execution_df index>,
            "make": <make name string>,
            "region": <region string>,
            "model_titles": [<model name strings>, ...],
            ...
        }

    parent: the existing Tk root (e.g. status.root) this window should
            be attached to. IMPORTANT: this must be a Toplevel of the
            existing root, not a second tk.Tk() - Tkinter does not
            support two independent Tk() roots running at once, which
            was causing this window to fail to show/respond properly.

    Returns: dict { idx: [selected model position indexes] }
    """

    # ---- Color palette (matches select_urls_ui) ----
    COLOR_BG        = "#f4f6f9"
    COLOR_HEADER_BG = "#1f2a44"
    COLOR_HEADER_FG = "#ffffff"
    COLOR_CARD_BG   = "#ffffff"
    COLOR_CARD_EDGE = "#dde3ec"
    COLOR_MAKE_FG   = "#1f2a44"
    COLOR_ACCENT    = "#2980b9"
    COLOR_GREEN     = "#27ae60"
    COLOR_RED       = "#e74c3c"
    COLOR_FOOTER_BG = "#ffffff"
    COLOR_MUTED     = "#6b7686"
    COLOR_SEARCH_BG = "#eef2f7"

    root = tk.Toplevel(parent) if parent is not None else tk.Tk()
    root.title("Select Models (All Selected Makes)")
    root.geometry("800x800")
    root.configure(bg=COLOR_BG)
    root.minsize(600, 500)

    # Make sure this window actually comes to the front and grabs focus
    root.lift()
    root.attributes("-topmost", True)
    root.after(200, lambda: root.attributes("-topmost", False))
    root.focus_force()

    # ---------- Header ----------
    header = tk.Frame(root, bg=COLOR_HEADER_BG, height=64)
    header.pack(fill="x", side="top")
    header.pack_propagate(False)

    tk.Label(
        header,
        text="SELECT MODELS",
        font=("Segoe UI", 16, "bold"),
        bg=COLOR_HEADER_BG,
        fg=COLOR_HEADER_FG
    ).pack(side="left", padx=20)

    tk.Label(
        header,
        text="Across every make you loaded",
        font=("Segoe UI", 10),
        bg=COLOR_HEADER_BG,
        fg="#c7cfe0"
    ).pack(side="left")

    # ---------- Prominent search bar (front and center, auto-focused) ----------
    search_bar = tk.Frame(root, bg=COLOR_SEARCH_BG, height=56)
    search_bar.pack(fill="x", side="top")
    search_bar.pack_propagate(False)

    search_inner = tk.Frame(search_bar, bg=COLOR_SEARCH_BG)
    search_inner.pack(fill="both", expand=True, padx=20, pady=10)

    tk.Label(
        search_inner,
        text="🔍",
        font=("Segoe UI", 12),
        bg=COLOR_SEARCH_BG,
        fg=COLOR_MUTED
    ).pack(side="left")

    search_var = tk.StringVar()

    search = tk.Entry(
        search_inner,
        textvariable=search_var,
        font=("Segoe UI", 12),
        relief="flat",
        bg="white",
        highlightthickness=1,
        highlightbackground=COLOR_CARD_EDGE,
        highlightcolor=COLOR_ACCENT
    )
    search.pack(side="left", fill="x", expand=True, padx=(8, 0), ipady=6)

    match_count_var = tk.StringVar(value="")
    tk.Label(
        search_inner,
        textvariable=match_count_var,
        font=("Segoe UI", 9),
        bg=COLOR_SEARCH_BG,
        fg=COLOR_MUTED
    ).pack(side="left", padx=10)

    # Search box gets focus the instant the window pops up, so you can
    # start typing immediately without clicking into it first.
    root.after(50, search.focus_set)
    # Ctrl+F also jumps back to search from anywhere in the window
    root.bind_all("<Control-f>", lambda e: search.focus_set())

    # ---------- Scrollable body ----------
    body_container = tk.Frame(root, bg=COLOR_BG)
    body_container.pack(fill="both", expand=True)

    canvas = tk.Canvas(body_container, bg=COLOR_BG, highlightthickness=0)
    scrollbar = ttk.Scrollbar(body_container, orient="vertical", command=canvas.yview)
    model_frame = tk.Frame(canvas, bg=COLOR_BG)

    model_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )

    canvas.create_window((0, 0), window=model_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side="left", fill="both", expand=True, padx=(15, 0), pady=(10, 10))
    scrollbar.pack(side="right", fill="y")

    def _on_mousewheel(event):
        canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    canvas.bind_all("<MouseWheel>", _on_mousewheel)

    # each entry: (IntVar, make_idx, model_position, Checkbutton widget, lowercase label for search)
    all_vars = []
    # per-make bookkeeping so we can collapse a whole section when it has no matches
    make_sections = []  # list of dicts: {header_frame, card, count_label, chk_list: [(chk, name)]}

    footer_count_var = tk.StringVar(value="0 model(s) selected")

    def refresh_footer_count():
        total = sum(v.get() for v, _, _, _, _ in all_vars)
        footer_count_var.set(f"{total} model(s) selected")

    for make_data in loaded_makes:
        make_idx = make_data["idx"]
        make_name = make_data["make"]
        region = make_data["region"]
        model_titles = make_data["model_titles"]

        section_header = tk.Frame(model_frame, bg=COLOR_BG)
        section_header.pack(fill="x", pady=(16, 6), padx=5)

        tk.Label(
            section_header,
            text=f"● {make_name}",
            font=("Segoe UI", 12, "bold"),
            bg=COLOR_BG,
            fg=COLOR_MAKE_FG
        ).pack(side="left")

        tk.Label(
            section_header,
            text=f"[{region}]",
            font=("Segoe UI", 9, "bold"),
            bg=COLOR_BG,
            fg=COLOR_ACCENT
        ).pack(side="left", padx=(6, 0))

        count_label = tk.Label(
            section_header,
            text=f"{len(model_titles)} model(s)",
            font=("Segoe UI", 9),
            bg=COLOR_BG,
            fg=COLOR_MUTED
        )
        count_label.pack(side="left", padx=(8, 0))

        card = tk.Frame(model_frame, bg=COLOR_CARD_BG, highlightbackground=COLOR_CARD_EDGE,
                         highlightthickness=1, bd=0)
        card.pack(fill="x", padx=5)

        chk_area = tk.Frame(card, bg=COLOR_CARD_BG)
        chk_area.pack(fill="x", padx=15, pady=10)

        section_chks = []

        for model_pos, model_text in enumerate(model_titles):
            var = tk.IntVar(value=1)
            var.trace_add("write", lambda *a: refresh_footer_count())

            chk = tk.Checkbutton(
                chk_area,
                text=model_text,
                variable=var,
                anchor="w",
                width=60,
                bg=COLOR_CARD_BG,
                activebackground=COLOR_CARD_BG,
                font=("Segoe UI", 10),
                selectcolor="#eaf3fb"
            )
            chk.pack(anchor="w", pady=1)

            name_lower = model_text.lower()
            all_vars.append((var, make_idx, model_pos, chk, name_lower))
            section_chks.append((chk, name_lower))

        make_sections.append({
            "header": section_header,
            "card": card,
            "count_label": count_label,
            "total": len(model_titles),
            "chks": section_chks
        })

    refresh_footer_count()

    # ---------- Select all / deselect all (global, applies to currently visible matches) ----------
    global_btn_row = tk.Frame(root, bg=COLOR_BG)
    global_btn_row.pack(fill="x", padx=20, pady=(0, 6))

    def select_all():
        for v, _, _, _, _ in all_vars:
            v.set(1)

    def deselect_all():
        for v, _, _, _, _ in all_vars:
            v.set(0)

    tk.Button(
        global_btn_row, text="Select All", bg=COLOR_GREEN, fg="white",
        font=("Segoe UI", 9, "bold"), relief="flat", padx=10, pady=3,
        command=select_all
    ).pack(side="left", padx=(0, 6))

    tk.Button(
        global_btn_row, text="Deselect All", bg=COLOR_RED, fg="white",
        font=("Segoe UI", 9, "bold"), relief="flat", padx=10, pady=3,
        command=deselect_all
    ).pack(side="left")

    # ---------- Search behavior ----------
    # Matching models stay visible (in their original order, so matches
    # naturally sit at the top of each section); non-matches are hidden.
    # A make's whole card collapses out of view if none of its models match.
    def search_models(*args):
        text = search_var.get().lower()
        total_matches = 0

        for section in make_sections:
            section_match_count = 0

            for chk, name in section["chks"]:
                if text in name:
                    chk.pack(anchor="w", pady=1)
                    section_match_count += 1
                    total_matches += 1
                else:
                    chk.pack_forget()

            section["count_label"].config(
                text=f"{section_match_count} / {section['total']} model(s)"
            )

            if text and section_match_count == 0:
                section["header"].pack_forget()
                section["card"].pack_forget()
            else:
                # ensure it's visible (in case a previous search hid it)
                if not section["header"].winfo_ismapped():
                    section["header"].pack(fill="x", pady=(16, 6), padx=5)
                    section["card"].pack(fill="x", padx=5)

        if text:
            match_count_var.set(f"{total_matches} match(es)")
        else:
            match_count_var.set("")

        canvas.configure(scrollregion=canvas.bbox("all"))

    search_var.trace_add("write", search_models)

    # ---------- Sticky footer ----------
    footer = tk.Frame(root, bg=COLOR_FOOTER_BG, height=64,
                       highlightbackground=COLOR_CARD_EDGE, highlightthickness=1)
    footer.pack(fill="x", side="bottom")
    footer.pack_propagate(False)

    tk.Label(
        footer,
        textvariable=footer_count_var,
        font=("Segoe UI", 10),
        bg=COLOR_FOOTER_BG,
        fg=COLOR_MUTED
    ).pack(side="left", padx=20)

    def run():
        selection = {}
        for var, make_idx, model_pos, _, _ in all_vars:
            if var.get():
                selection.setdefault(make_idx, []).append(model_pos)
        root.selection = selection
        root.destroy()

    tk.Button(
        footer,
        text="RUN SELECTED",
        bg=COLOR_ACCENT,
        fg="white",
        font=("Segoe UI", 11, "bold"),
        relief="flat",
        width=22,
        pady=8,
        command=run
    ).pack(side="right", padx=20, pady=10)

    if parent is not None:
        root.grab_set()
        root.wait_window()
    else:
        root.mainloop()

    if hasattr(root, "selection"):
        return root.selection

    return {}


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

        for el in elements:
            if el.is_displayed():
                return "FAIL"

        return "PASS"

    except:
        return "PASS"


def enter_iframe(driver, wait):

    try:
        iframe = WebDriverWait(driver, 40).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "iframe"))
        )

        driver.switch_to.frame(iframe)

    except TimeoutException:
        print("❌ Iframe not found on page")
        raise


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


def normalize_text(text):
    normalized = unicodedata.normalize('NFKD', text)
    ascii_text = normalized.encode('ASCII', 'ignore').decode('ASCII')
    return ascii_text.strip().lower()


def normalize_for_s3(text):
    if not text:
        return ""

    normalized = unicodedata.normalize('NFKD', text)
    ascii_text = normalized.encode('ASCII', 'ignore').decode('ASCII')

    ascii_text = ascii_text.lower()

    ascii_text = re.sub(r'[^a-z0-9\s-]', ' ', ascii_text)

    ascii_text = re.sub(r'\s+', ' ', ascii_text)

    ascii_text = ascii_text.strip().replace(' ', '-')

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


def check_interior_pano(year, make, alias, region):

    normalized_make = normalize_for_s3(make)
    normalized_alias = normalize_for_s3(alias)

    year_short = year[-2:]
    year_make = f"{year_short}{normalized_make}"

    if region == "US":
        base = "https://s3-us-west-2.amazonaws.com/static.izmocars.com/toolkit/newgen/"
    else:
        base = "https://s3-eu-west-1.amazonaws.com/staticeu.izmocars.com/toolkit/newgen/"

    pano_base = f"{base}{year}/{year_make}/{normalized_alias}/{normalized_alias}_panovr/"

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


# =====================================================
# NEW: ASYNC / BULK versions of Interior Pano + Blue Hotspot checks.
#
# The original check_interior_pano() / check_blue_hotspot() above make
# blocking `requests.head()` calls one model at a time, inside the
# Selenium loop - every model pays that network latency serially.
#
# These do the exact same checks (same URLs, same PASS/FAIL logic) but
# run for every model CONCURRENTLY in one batch via aiohttp, after all
# the Selenium clicking is done. Same accuracy, far less wall-clock time.
# =====================================================

async def check_interior_pano_async(session, year, make, alias, region):

    normalized_make = normalize_for_s3(make)
    normalized_alias = normalize_for_s3(alias)

    year_short = year[-2:]
    year_make = f"{year_short}{normalized_make}"

    if region == "US":
        base = "https://s3-us-west-2.amazonaws.com/static.izmocars.com/toolkit/newgen/"
    else:
        base = "https://s3-eu-west-1.amazonaws.com/staticeu.izmocars.com/toolkit/newgen/"

    pano_base = f"{base}{year}/{year_make}/{normalized_alias}/{normalized_alias}_panovr/"

    pano_files = ["izmoPano.xml", "pano.xml", "panovr.xml", "config.xml"]

    for file in pano_files:
        url = pano_base + file
        try:
            async with session.head(url, timeout=aiohttp.ClientTimeout(total=10)) as resp:
                if resp.status == 200:
                    return "PASS"
        except:
            pass

    return "FAIL"


async def check_blue_hotspot_async(session, year, make, alias, region):

    year_short = year[-2:]
    year_make = f"{year_short}{make.lower()}"

    base = (
        "https://s3-us-west-2.amazonaws.com/static.izmocars.com/toolkit/newgen/"
        if region == "US"
        else "https://s3-eu-west-1.amazonaws.com/staticeu.izmocars.com/toolkit/newgen/"
    )

    hotspot_url = f"{base}{year}/{year_make}/{alias}/{alias}_panovr/skin/hs.webp"

    try:
        async with session.head(hotspot_url, timeout=aiohttp.ClientTimeout(total=10)) as resp:
            return "PASS" if resp.status == 200 else "FAIL"
    except:
        return "FAIL"


async def run_pano_hotspot_checks_async(unique_keys):
    """
    unique_keys: list of (region, year, make, alias) tuples.
    Returns: dict { (region, year, make, alias): (interior_pano_status, blue_hotspot_status) }
    """

    connector = aiohttp.TCPConnector(limit=PARALLEL_REQUESTS, ttl_dns_cache=300, keepalive_timeout=30)

    async with aiohttp.ClientSession(connector=connector) as session:

        async def check_one(key):
            region, year, make, alias = key
            pano_status, hotspot_status = await asyncio.gather(
                check_interior_pano_async(session, year, make, alias, region),
                check_blue_hotspot_async(session, year, make, alias, region)
            )
            return key, (pano_status, hotspot_status)

        gathered = await asyncio.gather(*(check_one(k) for k in unique_keys))

    return dict(gathered)


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


def _looks_like_image(chunk: bytes) -> bool:
    """Check the first few bytes of a response against known image file
    signatures. A 200 status code alone doesn't prove the URL actually
    serves an image - some misconfigured/redirected endpoints return 200
    with an HTML error page or a placeholder. This confirms real image
    bytes are present."""
    if not chunk:
        return False
    if chunk[:4] == b"RIFF" and chunk[8:12] == b"WEBP":   # WebP
        return True
    if chunk[:3] == b"\xff\xd8\xff":                       # JPEG
        return True
    if chunk[:8] == b"\x89PNG\r\n\x1a\n":                  # PNG
        return True
    if chunk[:6] in (b"GIF87a", b"GIF89a"):                # GIF
        return True
    return False


async def _confirm_is_image(session, url, timeout):
    """Pull only the first 16 bytes (via a Range request) and confirm a
    real image signature is present. Cheap - no full image download."""
    headers = {"Range": "bytes=0-15", "User-Agent": "Mozilla/5.0"}
    try:
        async with session.get(url, headers=headers, timeout=timeout, allow_redirects=True) as resp:
            if resp.status not in (200, 206):
                return False
            chunk = await resp.content.read(16)
            return _looks_like_image(chunk)
    except Exception:
        return False


async def check_url(session, url):

    headers = {
        "User-Agent": "Mozilla/5.0",
        "Accept": "image/webp,image/*,*/*;q=0.8"
    }

    # Explicit per-request timeout. Without this, a request stuck behind a
    # long connection queue (common once you have thousands of URLs sharing
    # only PARALLEL_REQUESTS open connections) can hang until aiohttp's
    # default timeout, which used to get misreported as PASS below.
    request_timeout = aiohttp.ClientTimeout(total=15, connect=8)

    # Properly percent-encode special characters (apostrophes, accented
    # letters, etc. that can end up in color-name folders like
    # "reflet-d'argent") so the request reaches the intended resource
    # instead of failing on a malformed URL.
    safe_url = quote(url, safe=":/?#[]@!$&'()*+,;=~._-")

    for attempt in range(RETRIES):

        try:
            # Fast path: HEAD request - confirms the resource exists and is
            # an image WITHOUT downloading any image bytes.
            async with session.head(
                safe_url,
                headers=headers,
                allow_redirects=True,
                timeout=request_timeout
            ) as resp:

                if resp.status == 404:
                    return "FAIL"

                if resp.status in (200, 206):
                    content_type = resp.headers.get("Content-Type", "")
                    content_length = resp.headers.get("Content-Length")

                    is_image_type = content_type.lower().startswith("image/")
                    has_bytes = content_length is None or (content_length.isdigit() and int(content_length) > 0)

                    if is_image_type and has_bytes:
                        return "PASS"

                    # HEAD was ambiguous (missing/incorrect Content-Type, or
                    # zero-length response) - confirm with an actual byte
                    # signature check before deciding.
                    verified = await _confirm_is_image(session, safe_url, request_timeout)
                    return "PASS" if verified else "FAIL"

                if resp.status == 405:
                    # Some endpoints don't support HEAD - fall back straight
                    # to the signature-verified ranged GET.
                    verified = await _confirm_is_image(session, safe_url, request_timeout)
                    return "PASS" if verified else "FAIL"

                if resp.status in (403, 429, 500, 502, 503, 504):
                    # Transient / rate-limited - worth retrying with backoff
                    await asyncio.sleep(1 + attempt)
                    continue

                # Any other status code is not a "link is fine" signal -
                # treat it as broken rather than silently retrying/passing.
                return "FAIL"

        except (asyncio.TimeoutError, aiohttp.ClientError):
            await asyncio.sleep(1 + attempt)
        except Exception:
            await asyncio.sleep(1 + attempt)

    # IMPORTANT: if we exhausted every retry without a confirmed image
    # response, the link must be reported as FAIL, not PASS.
    return "FAIL"


async def validate_s3_links_async(s3_df):

    connector = aiohttp.TCPConnector(
        limit=PARALLEL_REQUESTS,
        ttl_dns_cache=300,
        keepalive_timeout=30
    )

    urls = s3_df["S3 URL"].tolist()
    statuses = [None] * len(urls)

    # Process in batches instead of firing every URL into one giant
    # asyncio.gather(). At 10k-11k+ URLs, queuing everything behind
    # PARALLEL_REQUESTS connections at once makes individual requests wait
    # long enough to hit the per-request timeout even when the link is
    # fine, which would show up as a false FAIL. Batches keep queue depth
    # bounded so requests get serviced promptly.
    BATCH_SIZE = 500

    async with aiohttp.ClientSession(connector=connector) as session:

        for start in range(0, len(urls), BATCH_SIZE):
            batch = urls[start:start + BATCH_SIZE]
            batch_results = await asyncio.gather(*[check_url(session, u) for u in batch])
            statuses[start:start + BATCH_SIZE] = batch_results

    s3_df["S3 Status"] = statuses

    return s3_df


def check_model_selector(driver, model_index):
    try:
        images = driver.find_elements(By.CSS_SELECTOR, "img")
        outside_s3 = images[model_index].get_attribute("src") or ""
        return "PASS" if ".webp" in outside_s3.lower() else "FAIL"
    except:
        return "FAIL"


class StatusWindow:

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Emporio Automation Status")
        self.root.geometry("700x500")

        tk.Label(
            self.root,
            text="EMPORIO AUTOMATION",
            font=("Arial", 16, "bold")
        ).pack(pady=10)

        self.status = tk.Label(
            self.root,
            text="Waiting...",
            font=("Arial", 12),
            fg="blue"
        )
        self.status.pack()

        tk.Label(
            self.root,
            text="Excel Generation"
        ).pack(pady=(20, 5))

        self.progress = ttk.Progressbar(
            self.root,
            orient="horizontal",
            length=500,
            mode="determinate",
            maximum=100
        )
        self.progress.pack()

        self.log = tk.Text(
            self.root,
            height=18,
            width=80
        )
        self.log.pack(pady=10)

        self.root.update()

    def update_status(self, msg):
        self.status.config(text=msg)
        self.root.update()

    def write(self, msg):
        self.log.insert(tk.END, msg + "\n")
        self.log.see(tk.END)
        self.root.update()

    def excel_progress(self, val, msg):
        self.progress["value"] = val
        self.status.config(text=msg)
        self.root.update()


# =====================================================
# NEW: LOAD A SINGLE MAKE (runs inside a worker thread)
# Opens its own Chrome driver, loads the URL, enters the
# iframe, and grabs the model title texts. Leaves the
# driver OPEN (does not quit) so it can be reused later
# for the actual model processing.
# =====================================================
def load_make(idx, row):

    url = row["URL"]
    region = row["Region"]
    make_name = row["MAKE"]

    driver = webdriver.Chrome(options=get_chrome_options())

    try:
        driver.set_page_load_timeout(300)
        wait = WebDriverWait(driver, 25)

        page_status = open_url(driver, url)

        if page_status == "LINK NOT LOADING":
            driver.quit()
            return {
                "idx": idx,
                "make": make_name,
                "region": region,
                "url": url,
                "driver": None,
                "wait": None,
                "model_titles": [],
                "status": "LINK NOT LOADING"
            }

        if HEADLESS:
            driver.set_window_size(1920, 1080)
        else:
            driver.maximize_window()

        enter_iframe(driver, wait)
        time.sleep(3)

        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "h4.modelTitle")))
        model_title_elems = driver.find_elements(By.CSS_SELECTOR, "h4.modelTitle")
        model_titles_text = [m.text.strip() for m in model_title_elems]

        return {
            "idx": idx,
            "make": make_name,
            "region": region,
            "url": url,
            "driver": driver,
            "wait": wait,
            "model_titles": model_titles_text,
            "status": "OK"
        }

    except Exception as e:
        print(f"❌ Error loading make '{make_name}': {e}")
        try:
            driver.quit()
        except:
            pass
        return {
            "idx": idx,
            "make": make_name,
            "region": region,
            "url": url,
            "driver": None,
            "wait": None,
            "model_titles": [],
            "status": f"ERROR: {e}"
        }


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

status = StatusWindow()
status.write("Automation Started")
status.update_status(f"Loading {len(selected_indices)} make(s) in parallel...")

# ---------------------------------------------------------
# STEP 1: Load ALL selected makes in parallel (each gets its
# own Chrome window loading simultaneously).
# ---------------------------------------------------------
makes_data = []

workers = min(MAX_PARALLEL_MAKES, len(selected_indices))

with ThreadPoolExecutor(max_workers=workers) as executor:
    futures = {
        executor.submit(load_make, idx, execution_df.loc[idx]): idx
        for idx in selected_indices
    }

    pending = set(futures.keys())

    while pending:
        # Wait in short bursts instead of blocking indefinitely, so we can
        # keep pumping the Tkinter event loop below and the window never
        # freezes / shows "(Not Responding)" while makes are still loading.
        done, pending = concurrent.futures.wait(
            pending, timeout=0.25, return_when=concurrent.futures.FIRST_COMPLETED
        )

        for future in done:
            result = future.result()
            makes_data.append(result)
            status.write(f"Loaded: {result['make']} [{result['region']}] -> {result['status']}")

        # Keep the Status window responsive even when nothing finished yet
        status.root.update()

    status.update_status(f"Loaded {len(makes_data)} make(s). Preparing model list...")

# Preserve original selection order for a nicer UI
makes_data.sort(key=lambda m: selected_indices.index(m["idx"]))

# Record failures directly into the results (same as your original "LINK NOT LOADING" handling)
for m in makes_data:
    if m["status"] != "OK":
        results.append({
            "URL": m["url"],
            "Status": m["status"]
        })

loaded_makes = [m for m in makes_data if m["status"] == "OK" and m["model_titles"]]

if not loaded_makes:
    status.write("❌ No makes loaded successfully / no models found. Exiting.")
else:
    # ---------------------------------------------------------
    # STEP 2: One combined model search/select window across
    # every successfully loaded make.
    # ---------------------------------------------------------
    status.update_status("Waiting for model selection...")
    selected_models_map = select_models_ui_multi(loaded_makes, parent=status.root)

    # ---------------------------------------------------------
    # STEP 3: Process only the models picked for each make,
    # reusing that make's already-open driver.
    # ---------------------------------------------------------
    for make_data in loaded_makes:

        make_idx = make_data["idx"]
        driver = make_data["driver"]
        wait = make_data["wait"]
        region = make_data["region"]

        selected_model_indexes = selected_models_map.get(make_idx, [])

        if not selected_model_indexes:
            driver.quit()
            continue

        status.update_status(f"Processing Make : {make_data['make']}")
        status.write("=" * 60)
        status.write(f"MAKE : {make_data['make']}")

        for i in selected_model_indexes:

            model_titles = driver.find_elements(By.CSS_SELECTOR, "h4.modelTitle")
            model_name = model_titles[i].text.strip()
            status.update_status(f"Processing Model : {model_name}")
            status.write(model_name)
            model_selector_status = check_model_selector(driver, i)

            # 🔽 CLICK MODEL
            driver.execute_script(
                "arguments[0].scrollIntoView({block:'center'}); arguments[0].click();",
                model_titles[i]
            )

            if i == 0:
                time.sleep(5)
            else:
                time.sleep(2)

            time.sleep(2)

            # 🔽 CHECK MODEL LOADING FIRST
            model_loading_status = check_model_loading(driver)

            base_trim = ""
            color_list = []
            row_duplicate_status = []
            color_count = 0
            inside_s3 = ""
            year = make = alias = "NOT_CAPTURED"
            ext_hotspot_status = "FAIL"
            interior_pano_status = "FAIL"
            blue_hotspot_status = "FAIL"

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

                status.update_status("Checking Colors...")
                color_list, row_duplicate_status = get_color_names_and_duplicate_status(driver)

                color_count = len(color_list)

                check_colors_broken(model_name, year, make, alias, color_list, s3_results, region)

                status.update_status("Checking External Hotspots...")
                ext_hotspot_status = check_ext_hotspot(driver)
                status.update_status("Checking Interior Panorama...")
                interior_pano_status = check_interior_pano(year, make, alias, region)
                status.update_status("Checking Blue Hotspots...")
                blue_hotspot_status = check_blue_hotspot(year, make, alias, region)

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

            status.write(f"✔ Completed : {model_name}")
            driver.switch_to.default_content()
            driver.back()
            time.sleep(2)
            enter_iframe(driver, wait)

        driver.quit()
        time.sleep(2)

# =====================================================
# VALIDATE + BUILD COMBINED EXCEL OUTPUT (unchanged)
# =====================================================
status.excel_progress(10, "Preparing Excel...")
s3_df = pd.DataFrame(s3_results)

if not s3_df.empty:
    s3_df = asyncio.run(validate_s3_links_async(s3_df))
    status.excel_progress(40, "Validating S3 Links...")

    status.excel_progress(60, "Creating Summary...")
    summary_df = s3_df.groupby("Model Name").agg(
        Total_URLs=("S3 URL", "count"),
        PASS_Count=("S3 Status", lambda x: (x == "PASS").sum()),
        BROKEN_Count=("S3 Status", lambda x: (x != "PASS").sum())
    ).reset_index()

    summary_df["PASS %"] = round((summary_df["PASS_Count"] / summary_df["Total_URLs"]) * 100, 2)
    summary_df["BROKEN %"] = round((summary_df["BROKEN_Count"] / summary_df["Total_URLs"]) * 100, 2)
else:
    summary_df = pd.DataFrame(columns=["Model Name", "Total_URLs", "PASS_Count", "BROKEN_Count", "PASS %", "BROKEN %"])

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
output_file = os.path.join(
    SCRIPT_DIR,
    f"Emporio_Final_Output_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
)

with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    pd.DataFrame(results).to_excel(writer, sheet_name="Main Report", index=False)
    status.excel_progress(75, "Writing Main Report...")
    s3_df.to_excel(writer, sheet_name="Color Broken Status", index=False)
    status.excel_progress(90, "Writing Color Broken Status...")
    summary_df.to_excel(writer, sheet_name="Model Summary", index=False)
    status.excel_progress(100, "Excel Generated Successfully")

print(f"✅ Excel Generated Successfully: {output_file}")
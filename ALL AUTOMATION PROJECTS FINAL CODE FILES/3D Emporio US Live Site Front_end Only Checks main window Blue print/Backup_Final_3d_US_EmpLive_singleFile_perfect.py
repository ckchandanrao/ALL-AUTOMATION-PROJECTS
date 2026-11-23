"""
izmo Emporio validation script - UNIVERSAL (works for any manufacturer's
izmo-emporio site, e.g. Chrysler, Volvo, etc.)

Usage:
    python emporio_validator.py [URL]

If no URL is given, defaults to the Chrysler Pacifica/Voyager site.

Writes results to an Excel file with columns:
    Sl.no, Year, Make, Model, Base Trim, Number of Trims, Default Color,
    Banner, IzmoEmporio Text, Int Hotspot Button, Ext Hotspot Button,
    BG Change Button, Number of BGs, Footer Buttons, Make Logo,
    Drag Text Logo

CHANGE FROM PREVIOUS VERSION:
    banner_status was already being computed (watch_banner_during_load)
    but was only ever printed to the console, never stored in the
    `results` dict or written to Excel. Added a "Banner" column and
    a "banner" key everywhere a result row is built (both FAIL
    shortcuts and the full success row), so it now shows up in the
    output file like every other check.
"""

import sys
import time
import re
from urllib.parse import urlparse

import requests

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter


# =========================================================
# CONFIGURATION
# =========================================================

DEFAULT_URL = "https://dodge3d.izmocars.com/emporio-en-us.htm"
URL = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_URL

INITIAL_LOAD_WAIT = 15
MODEL_LOAD_TIMEOUT = 60
BANNER_POLL_TIMEOUT = 15
FULL_CAR_TIMEOUT = 60
CAR_RENDER_WAIT = 3
DEFAULT_COLOR_TIMEOUT = 15
POPUP_OPEN_TIMEOUT = 15

REQUIRED_FOOTER_BUTTONS = 2

EXPECTED_BANNER_WIDTH = 1920
EXPECTED_BANNER_HEIGHT = 274

_site_slug = re.sub(r"[^a-zA-Z0-9]+", "_", urlparse(URL).netloc).strip("_") or "emporio"
OUTPUT_XLSX_PATH = f"emporio_validation_results_{_site_slug}.xlsx"

SEL_BANNER_IMG     = "izmo-emporio-div.visualizerLoader.popupbanner bannerloader img"
SEL_EDITION_INFO   = "izmo-emporio-div.izmo-emporio-visualizer-container > izmo-emporio-div.visHeader > izmo-emporio-div.edition_info > p:nth-child(1)"
SEL_BASE_TRIM      = "izmo-emporio-div.izmo-emporio-visualizer-container > izmo-emporio-div.visHeader > izmo-emporio-div.edition_info > p:nth-child(2)"
SEL_IZMO_SVG       = "izmo-emporio-div.izmo-emporio-visualizer-container izmo-emporio-div.regtextsm.action svg"
SEL_FOOTER_BTN     = "#footer3d > izmo-emporio-div.footer-button"
SEL_MAKE_LOGO      = "#footer3d > izmo-emporio-div.logo_group > img"
SEL_EXT_HOTSPOT    = "izmo-emporio-div.izmo-emporio-visualizer-container > izmo-emporio-div.fullScreenEyeConatiner > izmo-emporio-div.eyeButton > button"
SEL_INT_HOTSPOT    = "izmo-emporio-div.izmo-emporio-visualizer-container > izmo-emporio-div.pannel_section.fontFamilyRegular > izmo-emporio-div.pannel.interior-pano-button.action"
SEL_BG_CHANGE      = "izmo-emporio-div.izmo-emporio-visualizer-container > izmo-emporio-div.pannel_section.fontFamilyRegular > izmo-emporio-div:nth-child(2)"
# Accessory tab - same container/pattern as SEL_BG_CHANGE (nth-child(2)),
# just the next sibling (nth-child(3)). Deliberately NOT using the
# random session IDs from DevTools (#izmoe709-..., #izmoa1c4-...) -
# those are regenerated per page load and would break on the next run.
SEL_ACCESSORY_TAB  = "izmo-emporio-div.izmo-emporio-visualizer-container > izmo-emporio-div.pannel_section.fontFamilyRegular > izmo-emporio-div:nth-child(3)"
SEL_DEFAULT_COLOR  = "izmo-emporio-div.izmo-emporio-visualizer-container > izmo-emporio-div.color-section > p"
SEL_DRAG_TEXT_LOGO = "izmo-emporio-div.handICon"

# Candidate selectors for the trim button, outermost (most likely to
# carry the real click listener) to innermost.
TRIM_BUTTON_CANDIDATES = [
    ("trimInfo.action container",
     "izmo-emporio-div.izmo-emporio-visualizer-container > izmo-emporio-div.visHeader > izmo-emporio-div.trimInfo.action"),
    ("a.dropdown-toggle",
     "izmo-emporio-div.izmo-emporio-visualizer-container > izmo-emporio-div.visHeader > izmo-emporio-div.trimInfo.action > a.dropdown-toggle"),
    ("buttonTrimText",
     "izmo-emporio-div.izmo-emporio-visualizer-container > izmo-emporio-div.visHeader > izmo-emporio-div.trimInfo.action > a.dropdown-toggle > izmo-emporio-div.buttonTrimText"),
]

BACKGROUND_BUTTON_CANDIDATES = [
    ("pannel_section nth-child(2)", SEL_BG_CHANGE),
]

ACCESSORY_BUTTON_CANDIDATES = [
    ("pannel_section nth-child(3)", SEL_ACCESSORY_TAB),
]


# =========================================================
# CHROME SETUP
# =========================================================

options = Options()
options.add_argument("--disable-notifications")
options.add_argument("--start-maximized")
options.add_argument("--disk-cache-size=0")
options.add_argument("--media-cache-size=0")
options.add_argument("--headless=new")
options.add_argument("--window-size=1920,1080")
options.add_argument("--no-sandbox")

driver = webdriver.Chrome(options=options)
wait = WebDriverWait(driver, 60)

driver.execute_cdp_cmd("Network.enable", {})
driver.execute_cdp_cmd("Network.setCacheDisabled", {"cacheDisabled": True})


# =========================================================
# JS: FIND ALL MODELS (pierces shadow DOM)
# =========================================================

find_models_script = """
function findAllModels(root) {
    let models = [];
    function search(node) {
        if (!node) return;
        if (node.nodeType === Node.ELEMENT_NODE) {
            const id = node.id || "";
            const cls = typeof node.className === "string" ? node.className : "";
            if (id === "model-name" || cls.includes("model-name")) {
                if (!models.includes(node)) models.push(node);
            }
        }
        if (node.shadowRoot) search(node.shadowRoot);
        const children = node.children || [];
        for (const child of children) search(child);
    }
    search(root);
    return models;
}
return findAllModels(document);
"""


# =========================================================
# JS: GENERIC SHADOW-DOM-AWARE HELPERS
# =========================================================

get_text_by_selector_script = """
function findText(root, selector) {
    if (!root) return null;
    try {
        const el = root.querySelector(selector);
        if (el) {
            const t = (el.innerText || el.textContent || "").trim();
            if (t) return t;
        }
    } catch (e) {}
    const elements = root.querySelectorAll ? root.querySelectorAll("*") : [];
    for (const element of elements) {
        if (element.shadowRoot) {
            const result = findText(element.shadowRoot, selector);
            if (result) return result;
        }
    }
    return null;
}
return findText(document, arguments[0]);
"""

get_element_by_selector_script = """
function findElement(root, selector) {
    if (!root) return null;
    try {
        const el = root.querySelector(selector);
        if (el) return el;
    } catch (e) {}
    const elements = root.querySelectorAll ? root.querySelectorAll("*") : [];
    for (const element of elements) {
        if (element.shadowRoot) {
            const found = findElement(element.shadowRoot, selector);
            if (found) return found;
        }
    }
    return null;
}
return findElement(document, arguments[0]);
"""

count_elements_script = """
function countMatches(root, selector) {
    if (!root) return 0;
    let count = 0;
    try {
        count += root.querySelectorAll(selector).length;
    } catch (e) {}
    const elements = root.querySelectorAll ? root.querySelectorAll("*") : [];
    for (const element of elements) {
        if (element.shadowRoot) {
            count += countMatches(element.shadowRoot, selector);
        }
    }
    return count;
}
return countMatches(document, arguments[0]);
"""

count_footer_buttons_script = """
function countFooterButtons(root, containerSelector) {
    if (!root) return 0;
    let count = 0;
    try {
        const containers = root.querySelectorAll(containerSelector);
        for (const c of containers) {
            const directChildren = Array.from(c.children).filter(
                child => child.tagName && child.tagName.toLowerCase() === "izmo-emporio-div"
            );
            if (directChildren.length >= 1) {
                count += directChildren.length;
            } else {
                count += 1;
            }
        }
    } catch (e) {}
    const elements = root.querySelectorAll ? root.querySelectorAll("*") : [];
    for (const element of elements) {
        if (element.shadowRoot) {
            count += countFooterButtons(element.shadowRoot, containerSelector);
        }
    }
    return count;
}
return countFooterButtons(document, arguments[0]);
"""

banner_visible_script = """
function findVisibleBanner(root, selector) {
    if (!root) return null;
    try {
        const els = root.querySelectorAll(selector);
        for (const el of els) {
            const style = window.getComputedStyle(el);
            const rect = el.getBoundingClientRect();
            const parent = el.closest ? el.closest(".visualizerLoader") : null;
            const parentHidden = parent && parent.classList.contains("hide");
            if (
                style.display !== "none" &&
                style.visibility !== "hidden" &&
                style.opacity !== "0" &&
                rect.width > 0 && rect.height > 0 &&
                !parentHidden
            ) {
                return {
                    visible: true,
                    naturalWidth: el.naturalWidth || 0,
                    naturalHeight: el.naturalHeight || 0
                };
            }
        }
    } catch (e) {}
    const elements = root.querySelectorAll ? root.querySelectorAll("*") : [];
    for (const element of elements) {
        if (element.shadowRoot) {
            const found = findVisibleBanner(element.shadowRoot, selector);
            if (found) return found;
        }
    }
    return null;
}
return findVisibleBanner(document, arguments[0]);
"""

# UNIFIED counter for both the Trims popup and the Backgrounds popup —
# both reuse the identical structure: a VISIBLE (non-.hide)
# .more-items-popUp containing repeated .izmo-emporio-trim.item
# elements.
get_visible_popup_items_script = """
function findVisiblePopupItems(root) {
    if (!root) return {count: 0, items: []};
    let items = [];
    try {
        const popups = root.querySelectorAll(".more-items-popUp");
        for (const popup of popups) {
            if (popup.classList.contains("hide")) continue;
            const style = window.getComputedStyle(popup);
            const rect = popup.getBoundingClientRect();
            if (style.display === "none" || style.visibility === "hidden" || rect.width === 0) continue;

            const itemEls = popup.querySelectorAll(".izmo-emporio-trim.item");
            for (const el of itemEls) {
                items.push({
                    trim: el.getAttribute("trim") || "",
                    id: el.id || "",
                    text: (el.innerText || el.textContent || "").trim()
                });
            }
        }
    } catch (e) {}
    let count = items.length;
    const elements = root.querySelectorAll ? root.querySelectorAll("*") : [];
    for (const element of elements) {
        if (element.shadowRoot) {
            const sub = findVisiblePopupItems(element.shadowRoot);
            count += sub.count;
            items = items.concat(sub.items);
        }
    }
    return {count: count, items: items};
}
return findVisiblePopupItems(document);
"""

# DIAGNOSTIC ONLY: reports every .more-items-popUp in the DOM (hidden
# or not), used to explain a 0 count.
debug_all_popups_script = """
function findAllPopups(root) {
    if (!root) return [];
    let found = [];
    try {
        const popups = root.querySelectorAll(".more-items-popUp");
        for (const popup of popups) {
            const style = window.getComputedStyle(popup);
            const rect = popup.getBoundingClientRect();
            found.push({
                className: popup.className,
                hasHideClass: popup.classList.contains("hide"),
                display: style.display,
                width: Math.round(rect.width),
                height: Math.round(rect.height),
                itemCount: popup.querySelectorAll(".izmo-emporio-trim.item").length
            });
        }
    } catch (e) {}
    const elements = root.querySelectorAll ? root.querySelectorAll("*") : [];
    for (const element of elements) {
        if (element.shadowRoot) {
            found = found.concat(findAllPopups(element.shadowRoot));
        }
    }
    return found;
}
return findAllPopups(document);
"""

# Verifies whether ANY popup-like container exists yet (used right
# after a click attempt, before we know if it worked).
verify_any_popup_opened_script = """
function check(root) {
    if (!root) return false;
    try {
        const popups = root.querySelectorAll(
            ".more-items-popUp, .trim-popup-panel, .background-panel, .pannel-popup"
        );
        if (popups.length > 0) return true;
    } catch (e) {}
    const elements = root.querySelectorAll ? root.querySelectorAll("*") : [];
    for (const element of elements) {
        if (element.shadowRoot) {
            if (check(element.shadowRoot)) return true;
        }
    }
    return false;
}
return check(document);
"""

find_by_exact_text_script = """
function findByText(root, text) {
    if (!root) return null;
    try {
        const all = root.querySelectorAll("*");
        for (const el of all) {
            const t = (el.innerText || el.textContent || "").trim().toUpperCase();
            if (t !== text) continue;
            const rect = el.getBoundingClientRect();
            if (rect.width > 0 && rect.height > 0) return el;
        }
    } catch (e) {}
    const elements = root.querySelectorAll ? root.querySelectorAll("*") : [];
    for (const element of elements) {
        if (element.shadowRoot) {
            const found = findByText(element.shadowRoot, text);
            if (found) return found;
        }
    }
    return null;
}
return findByText(document, arguments[0]);
"""


# =========================================================
# UTILITIES
# =========================================================

def normalize_text(text):
    if not text:
        return ""
    return " ".join(text.split()).lower()


def get_models(driver):
    return driver.execute_script(find_models_script)


def get_edition_info_text(driver):
    return driver.execute_script(get_text_by_selector_script, SEL_EDITION_INFO)


def get_text(driver, selector):
    return driver.execute_script(get_text_by_selector_script, selector)


def get_element(driver, selector):
    try:
        return driver.execute_script(get_element_by_selector_script, selector)
    except Exception:
        return None


def count_matches(driver, selector):
    try:
        return driver.execute_script(count_elements_script, selector)
    except Exception:
        return 0


def count_footer_buttons(driver):
    try:
        return driver.execute_script(count_footer_buttons_script, SEL_FOOTER_BTN)
    except Exception:
        return 0


def wait_for_selected_model(driver, expected_model_name):
    expected = normalize_text(expected_model_name)

    def selected_model_loaded(driver):
        current = get_edition_info_text(driver)
        if not current:
            return False
        if expected in normalize_text(current):
            return current
        return False

    return WebDriverWait(driver, MODEL_LOAD_TIMEOUT).until(selected_model_loaded)


def watch_banner_during_load(driver, timeout=BANNER_POLL_TIMEOUT):
    start = time.time()
    end_time = start + timeout
    found = None
    while time.time() < end_time:
        try:
            result = driver.execute_script(banner_visible_script, SEL_BANNER_IMG)
            if result and result.get("visible"):
                found = result
                break
        except Exception:
            pass
        time.sleep(0.1)

    if not found:
        return "FAIL"

    nw, nh = found.get("naturalWidth", 0), found.get("naturalHeight", 0)
    if nw == EXPECTED_BANNER_WIDTH and nh == EXPECTED_BANNER_HEIGHT:
        return "PASS"
    return "FAIL"


def load_fresh_page(driver, label=""):
    print(f"Clearing cache and reloading page {label}...")
    try:
        driver.execute_cdp_cmd("Network.clearBrowserCache", {})
    except Exception as e:
        print("  (cache clear warning:", e, ")")

    driver.get(URL)

    WebDriverWait(driver, 30).until(
        lambda d: d.execute_script("return document.readyState") == "complete"
    )
    wait.until(lambda d: d.find_elements(By.CSS_SELECTOR, "izmo-emporio"))
    print("PAGE LOADED URL:", driver.current_url)

    time.sleep(INITIAL_LOAD_WAIT)


def check_default_color(driver, timeout=DEFAULT_COLOR_TIMEOUT):
    end_time = time.time() + timeout
    last_seen = None
    while time.time() < end_time:
        try:
            name = get_text(driver, SEL_DEFAULT_COLOR)
        except Exception:
            name = None
        if name and name.strip():
            return name.strip()
        last_seen = name
        time.sleep(0.3)
    return last_seen.strip() if last_seen else None


# =========================================================
# ROBUST POPUP OPENING (multi-strategy click + verify)
# =========================================================

def _try_click_strategies(driver, element):
    """Tries several ways of clicking the same element, returning True
    as soon as one doesn't raise an exception. Does NOT verify the
    popup opened — that's the caller's job."""
    driver.execute_script(
        "arguments[0].scrollIntoView({behavior:'instant', block:'center', inline:'center'});",
        element,
    )
    time.sleep(0.2)

    strategies = [
        ("ActionChains click", lambda el: ActionChains(driver).move_to_element(el).pause(0.15).click().perform()),
        ("JS dispatchEvent click", lambda el: driver.execute_script(
            "arguments[0].dispatchEvent(new MouseEvent('click', {bubbles:true, cancelable:true, view:window}));", el
        )),
        ("native .click()", lambda el: el.click()),
        ("mousedown+mouseup+click sequence", lambda el: driver.execute_script(
            """
            const el = arguments[0];
            const r = el.getBoundingClientRect();
            const x = r.left + r.width / 2, y = r.top + r.height / 2;
            for (const type of ['mouseover', 'mousedown', 'mouseup', 'click']) {
                el.dispatchEvent(new MouseEvent(type, {bubbles:true, cancelable:true, clientX:x, clientY:y, view:window}));
            }
            """, el
        )),
    ]

    for label, action in strategies:
        try:
            action(element)
            return label
        except Exception:
            continue
    return None


def open_popup_robust(driver, candidates, timeout=POPUP_OPEN_TIMEOUT):
    """Tries each candidate selector, and for each, tries multiple
    click strategies, checking after every attempt whether a popup
    verifiably opened. Returns True on success."""
    for label, selector in candidates:
        element = get_element(driver, selector)
        if element is None:
            print(f"  [{label}]: element not found")
            continue

        click_label = _try_click_strategies(driver, element)
        if click_label is None:
            print(f"  [{label}]: all click strategies raised exceptions")
            continue

        end_time = time.time() + 3  # short check after each click attempt
        opened = False
        while time.time() < end_time:
            try:
                if driver.execute_script(verify_any_popup_opened_script):
                    opened = True
                    break
            except Exception:
                pass
            time.sleep(0.2)

        if opened:
            print(f"  OPENED via [{label}] using strategy '{click_label}'")
            return True
        else:
            print(f"  [{label}] clicked ('{click_label}') but no popup appeared")

    # Last resort: exact text match on "SELECT TRIM"
    try:
        element = driver.execute_script(find_by_exact_text_script, "SELECT TRIM")
        if element is not None:
            click_label = _try_click_strategies(driver, element)
            if click_label:
                end_time = time.time() + 3
                while time.time() < end_time:
                    if driver.execute_script(verify_any_popup_opened_script):
                        print(f"  OPENED via [exact text 'SELECT TRIM'] using strategy '{click_label}'")
                        return True
                    time.sleep(0.2)
    except Exception:
        pass

    return False


def wait_for_popup_item_count(driver, timeout=15, debug_label=""):
    end_time = time.time() + timeout
    last = (0, [])
    while time.time() < end_time:
        try:
            result = driver.execute_script(get_visible_popup_items_script)
            count = result.get("count", 0)
            items = result.get("items", [])
            if count > 0:
                return count, items
            last = (count, items)
        except Exception:
            pass
        time.sleep(0.3)

    try:
        popups = driver.execute_script(debug_all_popups_script)
        if popups:
            print(f"  [{debug_label}] DIAGNOSTIC: {len(popups)} .more-items-popUp element(s) in DOM:")
            for p in popups:
                print(f"    class='{p['className']}' hide={p['hasHideClass']} "
                      f"display={p['display']} size={p['width']}x{p['height']} items={p['itemCount']}")
        else:
            print(f"  [{debug_label}] DIAGNOSTIC: no .more-items-popUp exists in the DOM at all")
    except Exception as e:
        print(f"  [{debug_label}] DIAGNOSTIC failed:", e)

    return last


def check_number_of_trims(driver):
    print("Opening trims popup...")
    opened = open_popup_robust(driver, TRIM_BUTTON_CANDIDATES)
    if not opened:
        print("TRIM POPUP: failed to open via any strategy")
        return 0

    count, items = wait_for_popup_item_count(driver, debug_label="TRIMS")
    trim_names = [it.get("trim", "") for it in items]
    print(f"NUMBER OF TRIMS ({count}):", trim_names)
    return count


def check_number_of_backgrounds(driver):
    print("Opening backgrounds popup...")
    opened = open_popup_robust(driver, BACKGROUND_BUTTON_CANDIDATES)
    if not opened:
        print("BACKGROUND POPUP: failed to open via any strategy")
        return 0

    count, items = wait_for_popup_item_count(driver, debug_label="BACKGROUNDS")
    bg_ids = [it.get("id", "") for it in items]
    print(f"NUMBER OF BGs ({count}):", bg_ids)
    return count


def check_accessory_tab_and_count(driver):
    """
    Validate the Accessories tab and calculate the accessory count.

    Counting rule:
        - No accessory cards -> 0
        - Card has displayed number N -> N + 1
        - Card has no displayed number -> 1

    Example:
        18 -> 19
        9  -> 10

    The accessory cards are identified from the actual DOM structure:
        izmo-emporio-div.izmo-emporio-accessory.item.action[accessoryid]

    Hidden accessory cards are included because the popup/carousel can
    contain valid cards marked with the `hide` class.

    The number is NOT taken from the whole card text as the primary source.
    This is important because product names can contain numbers such as "2pc".
    """

    print("Checking Accessories tab...")

    accessory_tab_status = (
        "PASS"
        if count_matches(driver, SEL_ACCESSORY_TAB) > 0
        else "FAIL"
    )

    print("Accessory Tab       :", accessory_tab_status)

    if accessory_tab_status == "FAIL":
        print("ACCESSORIES: tab not found")
        return accessory_tab_status, 0

    accessory_tab = get_element(driver, SEL_ACCESSORY_TAB)

    if accessory_tab is None:
        print("ACCESSORIES: tab element could not be retrieved")
        return accessory_tab_status, 0

    click_label = _try_click_strategies(driver, accessory_tab)

    if click_label is None:
        print("ACCESSORIES: click failed")
        return accessory_tab_status, 0

    print(
        f"ACCESSORIES TAB CLICKED using strategy '{click_label}'"
    )

    # ---------------------------------------------------------
    # JavaScript:
    # Find ALL accessory cards belonging to the current model.
    #
    # IMPORTANT: hidden cards are intentionally included. The Accessories
    # popup/carousel may mark cards with the `hide` class even though those
    # cards are valid accessories for the current model.
    # ---------------------------------------------------------
    accessory_cards_script = r"""
    function elementIsShown(el) {
        if (!el) return false;

        const style = window.getComputedStyle(el);

        if (style.display === "none" ||
            style.visibility === "hidden" ||
            style.opacity === "0") {
            return false;
        }

        return true;
    }

    function hasRenderedChild(el) {
        if (!el) return false;

        // Check the element itself.
        if (elementIsShown(el)) {
            try {
                const r = el.getBoundingClientRect();
                if (r.width > 0 && r.height > 0) {
                    return true;
                }
            } catch (e) {}
        }

        // A custom element can have no box while its child content
        // is rendered. Check descendants instead.
        const descendants = el.querySelectorAll
            ? el.querySelectorAll("*")
            : [];

        for (const child of descendants) {
            if (!elementIsShown(child)) continue;

            try {
                const r = child.getBoundingClientRect();
                if (r.width > 0 && r.height > 0) {
                    return true;
                }
            } catch (e) {}
        }

        return false;
    }

    function findCards(root) {
        if (!root) return [];

        let cards = [];

        try {
            const found = root.querySelectorAll(
                'izmo-emporio-div.izmo-emporio-accessory.item.action[accessoryid]'
            );

            for (const card of found) {
                // IMPORTANT: do NOT filter by the "hide" class or by
                // rendered size here. The Accessories carousel can keep
                // valid accessory cards hidden while still keeping them in
                // the current model's popup. We need ALL cards for the
                // current model, not only the currently visible page.
                cards.push(card);
            }
        } catch (e) {}

        // Search every shadow root recursively.
        const elements = root.querySelectorAll
            ? root.querySelectorAll("*")
            : [];

        for (const element of elements) {
            if (element.shadowRoot) {
                cards = cards.concat(
                    findCards(element.shadowRoot)
                );
            }
        }

        return cards;
    }

    function cleanPseudoContent(value) {
        if (value === null || value === undefined) return null;

        let v = String(value).trim();

        if (!v || v === "none" || v === "normal") {
            return null;
        }

        // CSS content can be returned as "18".
        v = v.replace(/^['\"]|['\"]$/g, "").trim();

        return /^\d+$/.test(v) ? v : null;
    }

    function getPseudoNumber(el, pseudo) {
        try {
            return cleanPseudoContent(
                window.getComputedStyle(el, pseudo).content
            );
        } catch (e) {
            return null;
        }
    }

    function exactNumericText(el) {
        const text = (
            el.innerText ||
            el.textContent ||
            ""
        ).trim();

        if (/^\d+$/.test(text)) {
            return text;
        }

        return null;
    }

    function getCardInfo(card) {
        const info = {
            accessoryid: card.getAttribute("accessoryid") || "",
            trim: card.getAttribute("trim") || "",
            numberCandidates: []
        };

        function addCandidate(value, source, element) {
            if (value === null || value === undefined) return;

            const n = String(value).trim();
            if (!/^\d+$/.test(n)) return;

            let x = 0;
            let y = 0;
            let width = 0;
            let height = 0;

            try {
                const r = element.getBoundingClientRect();
                x = r.left + r.width / 2;
                y = r.top + r.height / 2;
                width = r.width;
                height = r.height;
            } catch (e) {}

            info.numberCandidates.push({
                value: parseInt(n, 10),
                source: source,
                x: x,
                y: y,
                width: width,
                height: height
            });
        }

        // Check the card and every descendant.
        const nodes = [card];
        const descendants = card.querySelectorAll
            ? card.querySelectorAll("*")
            : [];

        for (const el of descendants) {
            if (elementIsShown(el)) {
                nodes.push(el);
            }
        }

        for (const el of nodes) {

            // 1. CSS generated badge/content.
            for (const pseudo of ["::before", "::after"]) {
                const n = getPseudoNumber(el, pseudo);
                if (n !== null) {
                    addCandidate(n, "CSS " + pseudo, el);
                }
            }

            // 2. A visible element containing ONLY digits.
            if (elementIsShown(el)) {
                const n = exactNumericText(el);
                if (n !== null) {
                    addCandidate(n, "numeric element", el);
                }
            }

            // 3. Explicit index/count/number-like attributes.
            try {
                for (const attr of Array.from(el.attributes || [])) {
                    const name = attr.name.toLowerCase();
                    const value = attr.value.trim();

                    if (
                        name.includes("index") ||
                        name.includes("count") ||
                        name.includes("number") ||
                        name.includes("position") ||
                        name.includes("order") ||
                        name.includes("seq")
                    ) {
                        if (/^\d+$/.test(value)) {
                            addCandidate(
                                value,
                                "index/count attribute",
                                el
                            );
                        }
                    }
                }
            } catch (e) {}
        }

        return info;
    }

    return findCards(document).map(getCardInfo);
    """

    # Wait briefly for the Accessories panel to render.
    end_time = time.time() + 10
    cards = []

    while time.time() < end_time:
        try:
            cards = driver.execute_script(
                accessory_cards_script
            ) or []

            if cards:
                break
        except Exception:
            pass

        time.sleep(0.25)

    # ---------------------------------------------------------
    # Deduplicate by accessoryid.
    # ---------------------------------------------------------
    unique_cards = []
    seen_ids = set()

    for card in cards:
        accessory_id = (
            card.get("accessoryid") or ""
        ).strip()

        if not accessory_id:
            continue

        if accessory_id in seen_ids:
            continue

        seen_ids.add(accessory_id)
        unique_cards.append(card)

    cards = unique_cards

    print(
        "TOTAL ACCESSORY CARDS FOUND:",
        len(cards)
    )

    # No cards = 0.
    if not cards:
        print("NUMBER OF ACCESSORIES: 0")
        return accessory_tab_status, 0

    total_accessories = 0

    for index, card in enumerate(cards, start=1):

        trim_name = card.get("trim") or ""
        candidates = card.get("numberCandidates") or []

        # -----------------------------------------------------
        # Select the displayed badge number.
        #
        # The badge in the UI is normally positioned toward the
        # lower/right portion of the card. Prefer candidates there.
        # This prevents numbers in unrelated card content from
        # being selected.
        # -----------------------------------------------------
        number = None
        source = None

        if candidates:

            # Prefer CSS generated numbers first because the UI badge
            # is commonly rendered with ::before / ::after.
            css_candidates = [
                c for c in candidates
                if c.get("source", "").startswith("CSS")
            ]

            if css_candidates:
                number_candidate = css_candidates[-1]
            else:
                # Prefer candidates located toward the lower-right of
                # the rendered card content.
                number_candidate = candidates[-1]

            number = number_candidate.get("value")
            source = number_candidate.get("source")

        if number is not None:
            contribution = number + 1

            print(
                f"  ACCESSORY {index}: {trim_name} | "
                f"number={number} ({source}) -> "
                f"contributes {contribution}"
            )
        else:
            contribution = 1

            print(
                f"  ACCESSORY {index}: {trim_name} | "
                f"no number -> contributes 1"
            )

        total_accessories += contribution

    print(
        "NUMBER OF ACCESSORIES:",
        total_accessories
    )

    return accessory_tab_status, total_accessories


# =========================================================
# EXCEL OUTPUT
# =========================================================

# NOTE: "Banner" added here, right after "Number of Trims" / before
# "Default Color" - reorder if you'd rather it sit elsewhere, just
# keep EXCEL_COLUMNS and _row_values() in the same order.
EXCEL_COLUMNS = [
    "Sl.no", "Year", "Make", "Model", "Base Trim", "Number of Trims",
    "Banner", "Default Color", "IzmoEmporio Text", "Int Hotspot Button",
    "Ext Hotspot Button", "BG Change Button", "Number of BGs",
    "Accessory Tab", "Number of Accessories",
    "Footer Buttons", "Make Logo", "Drag Text Logo",
]


def _row_values(r):
    return [
        r["number"], r["year"], r["make"], r["model"], r["base_trim"],
        r["number_of_trims"], r["banner"], r["default_color"], r["izmo_emporio"],
        r["int_hotspot"], r["ext_hotspot"], r["bg_change"],
        r["number_of_backgrounds"], r["accessory_tab"],
        r["number_of_accessories"], r["footer_buttons"], r["make_logo"],
        r["drag_text_logo"],
    ]


def write_results_to_excel(results, path=OUTPUT_XLSX_PATH):
    wb = Workbook()
    ws = wb.active
    ws.title = "Validation Results"

    header_font = Font(name="Arial", bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    body_font = Font(name="Arial")
    center = Alignment(horizontal="center", vertical="center")

    for col_idx, header in enumerate(EXCEL_COLUMNS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center

    for row_idx, r in enumerate(results, start=2):
        row_values = _row_values(r)
        for col_idx, value in enumerate(row_values, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.font = body_font
            cell.alignment = center

    for col_idx, header in enumerate(EXCEL_COLUMNS, start=1):
        max_len = max(
            [len(str(header))] + [len(str(_row_values(r)[col_idx - 1])) for r in results]
            if results else [len(str(header))]
        )
        ws.column_dimensions[get_column_letter(col_idx)].width = max_len + 4

    wb.save(path)
    print("")
    print(f"Excel results written to: {path}")


# =========================================================
# a) VALIDATE THE URL
# =========================================================

print("")
print("============================================================")
print("STEP A: VALIDATING URL")
print("============================================================")

url_valid = False

try:
    resp = requests.head(URL, allow_redirects=True, timeout=15)
    if resp.status_code >= 400:
        resp = requests.get(URL, allow_redirects=True, timeout=20)
    print("URL           :", URL)
    print("STATUS CODE   :", resp.status_code)
    url_valid = resp.status_code < 400
except requests.RequestException as e:
    print("URL REQUEST FAILED:", e)
    url_valid = False

print("URL VALID     :", "PASS" if url_valid else "FAIL")

if not url_valid:
    print("")
    print("URL failed validation. Aborting before opening the browser.")
    driver.quit()
    raise SystemExit


# =========================================================
# MAIN FLOW
# =========================================================

results = []

try:
    print("")
    print("============================================================")
    print("DISCOVERY: loading page once to count available models")
    print("============================================================")

    load_fresh_page(driver, label="(discovery)")

    models = get_models(driver)
    total_models = len(models)

    print("")
    print("TOTAL MODELS FOUND:", total_models)

    if total_models == 0:
        print("NO MODELS FOUND")
        raise SystemExit

    for model_index in range(total_models):

        print("")
        print("############################################################")
        print(f"VALIDATING MODEL {model_index + 1} / {total_models}")
        print("############################################################")

        load_fresh_page(driver, label=f"(model {model_index + 1})")

        models = get_models(driver)

        if model_index >= len(models):
            print("MODEL ELEMENT NOT FOUND")
            results.append({
                "number": model_index + 1, "year": "", "make": "", "model": "",
                "base_trim": "", "number_of_trims": 0, "banner": "FAIL",
                "default_color": "",
                "izmo_emporio": "FAIL", "int_hotspot": "FAIL",
                "ext_hotspot": "FAIL", "bg_change": "FAIL", "number_of_backgrounds": 0,
                "accessory_tab": "FAIL", "number_of_accessories": 0,
                "footer_buttons": "FAIL", "make_logo": "FAIL",
                "drag_text_logo": "FAIL",
            })
            continue

        model_element = models[model_index]

        listing_model_name = driver.execute_script(
            "return (arguments[0].innerText || arguments[0].textContent || '').trim();",
            model_element,
        )
        listing_model_name = " ".join(listing_model_name.split())
        print("MODEL NAME:", listing_model_name)

        driver.execute_script(
            "arguments[0].scrollIntoView({behavior:'instant', block:'center', inline:'center'});",
            model_element,
        )
        time.sleep(0.3)
        try:
            ActionChains(driver).move_to_element(model_element).pause(0.1).click().perform()
        except Exception:
            model_element.click()
        print("MODEL CLICKED")

        print("Watching for banner during load...")
        banner_status = watch_banner_during_load(driver)
        print("BANNER:", banner_status)

        print("Waiting for selected vehicle to load...")

        try:
            wait_for_selected_model(driver, listing_model_name)
            print("VEHICLE LOADED: PASS")
        except TimeoutException:
            print("VEHICLE LOADED: FAIL")
            results.append({
                "number": model_index + 1, "year": "", "make": "", "model": "",
                "base_trim": "", "number_of_trims": 0, "banner": banner_status,
                "default_color": "",
                "izmo_emporio": "FAIL", "int_hotspot": "FAIL",
                "ext_hotspot": "FAIL", "bg_change": "FAIL", "number_of_backgrounds": 0,
                "accessory_tab": "FAIL", "number_of_accessories": 0,
                "footer_buttons": "FAIL", "make_logo": "FAIL",
                "drag_text_logo": "FAIL",
            })
            print(f"MODEL {model_index + 1}: FAILED TO LOAD")
            continue

        print("Waiting for 3D car rendering...")
        time.sleep(CAR_RENDER_WAIT)

        try:
            def all_elements_present(d):
                izmo_ok = count_matches(d, SEL_IZMO_SVG) > 0
                logo_ok = count_matches(d, SEL_MAKE_LOGO) > 0
                footer_ok = count_footer_buttons(d) >= REQUIRED_FOOTER_BUTTONS
                return izmo_ok and logo_ok and footer_ok

            WebDriverWait(driver, FULL_CAR_TIMEOUT).until(all_elements_present)
        except Exception as e:
            print("Some UI elements did not settle in time:", e)

        izmo_status = "PASS" if count_matches(driver, SEL_IZMO_SVG) > 0 else "FAIL"
        footer_count = count_footer_buttons(driver)
        footer_status = "PASS" if footer_count >= REQUIRED_FOOTER_BUTTONS else "FAIL"
        logo_status = "PASS" if count_matches(driver, SEL_MAKE_LOGO) > 0 else "FAIL"
        ext_hotspot_status = "PASS" if count_matches(driver, SEL_EXT_HOTSPOT) > 0 else "FAIL"
        int_hotspot_status = "PASS" if count_matches(driver, SEL_INT_HOTSPOT) > 0 else "FAIL"
        bg_change_status = "PASS" if count_matches(driver, SEL_BG_CHANGE) > 0 else "FAIL"

        # Drag Text Logo:
        # PASS only when the handICon class is present.
        drag_text_logo_status = (
            "PASS"
            if count_matches(driver, SEL_DRAG_TEXT_LOGO) > 0
            else "FAIL"
        )

        print("IzmoEmporio Text    :", izmo_status)
        print("Footer Buttons      :", footer_status, f"({footer_count} found)")
        print("Make Logo           :", logo_status)
        print("Drag Text Logo      :", drag_text_logo_status)
        print("Ext Hotspot Button  :", ext_hotspot_status)
        print("Int Hotspot Button  :", int_hotspot_status)
        print("BG Change Button    :", bg_change_status)

        year, make, vehicle_model = "", "", ""
        model_text = get_edition_info_text(driver)
        if model_text:
            parts = model_text.split()
            if len(parts) >= 3:
                year = parts[0]
                make = parts[1]
                vehicle_model = " ".join(parts[2:])
        print("Year / Make / Model :", year, make, vehicle_model)

        base_trim = get_text(driver, SEL_BASE_TRIM) or ""
        print("Base Trim           :", base_trim or "FAIL (not found)")

        default_color = check_default_color(driver) or ""
        print("Default Color       :", default_color or "FAIL (not found)")

        number_of_trims = check_number_of_trims(driver)
        number_of_backgrounds = check_number_of_backgrounds(driver)

        accessory_tab_status, number_of_accessories = check_accessory_tab_and_count(driver)

        results.append({
            "number": model_index + 1,
            "year": year,
            "make": make,
            "model": vehicle_model,
            "base_trim": base_trim,
            "number_of_trims": number_of_trims,
            "banner": banner_status,
            "default_color": default_color,
            "izmo_emporio": izmo_status,
            "int_hotspot": int_hotspot_status,
            "ext_hotspot": ext_hotspot_status,
            "bg_change": bg_change_status,
            "number_of_backgrounds": number_of_backgrounds,
            "accessory_tab": accessory_tab_status,
            "number_of_accessories": number_of_accessories,
            "footer_buttons": footer_status,
            "make_logo": logo_status,
            "drag_text_logo": drag_text_logo_status,
        })

        print("")
        print(f"MODEL {model_index + 1} VALIDATION COMPLETED")

    write_results_to_excel(results)

finally:
    driver.quit()
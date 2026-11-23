# Generated from the user's uploaded Emporio validation script.
"""
izmo Emporio validation script - UNIVERSAL (works for any manufacturer's
izmo-emporio site, e.g. Chrysler, Volvo, etc.)

Usage:
    python emporio_validator.py [URL]

If no URL is given, defaults to the Chrysler Pacifica/Voyager site.

Writes results to an Excel file with FOUR sheets:

  Sheet 1 "Validation Results" (unchanged from before):
    Sl.no, Year, Make, Model, Base Trim, Number of Trims, Banner,
    Default Color, IzmoEmporio Text, Int Hotspot Button,
    Ext Hotspot Button, BG Change Button, Number of BGs,
    Accessory Tab, Number of Accessories, Footer Buttons, Make Logo,
    Drag Text Logo

  Sheet 2 "Main" (NEW):
    Year, Make, Model, Other Trims, Number of Accessories,
    Checked date with time
    Base-trim rows are highlighted.

CHANGE FROM PREVIOUS VERSION:
    - check_number_of_trims() already fetched every trim's name while
      counting them, but only ever printed the list and returned the
      count - the names were thrown away. Now returned and stored.
    - check_accessory_tab_and_count() already read each accessory
      card's "trim" attribute while counting, but likewise discarded
      it. Now also tries to read an ".item-name" label (same class
      pattern used for trim names) as a better display name, and both
      are returned alongside the count.
    - Added get_all_color_names(), a NEW best-effort check. There was
      no existing color-list popup logic to reuse (only the single
      currently-selected color name via SEL_DEFAULT_COLOR existed).
      This guesses at a color-item pattern that mirrors the site's own
      naming convention (izmo-emporio-trim.item, izmo-emporio-
      accessory.item -> izmo-emporio-color.item). If this returns
      nothing useful on your site, send the exact DevTools selector
      for one color swatch/popup item (the same way you did for the
      Accessories tab) and it can be made exact.
"""

import sys
import time
import re
from datetime import datetime
from zoneinfo import ZoneInfo
from urllib.parse import urlparse

import requests

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter


# =========================================================
# CONFIGURATION
# =========================================================

DEFAULT_URL = "https://ram3d.izdemosite.com/emporio-en-us.htm"
URL = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_URL

INITIAL_LOAD_WAIT = 15
MODEL_LOAD_TIMEOUT = 60
BANNER_POLL_TIMEOUT = 15
FULL_CAR_TIMEOUT = 60
CAR_RENDER_WAIT = 3
DEFAULT_COLOR_TIMEOUT = 20
POPUP_OPEN_TIMEOUT = 15

REQUIRED_FOOTER_BUTTONS = 2

EXPECTED_BANNER_WIDTH = 1920
EXPECTED_BANNER_HEIGHT = 274

_site_slug = re.sub(r"[^a-zA-Z0-9]+", "_", urlparse(URL).netloc).strip("_") or "emporio"
OUTPUT_XLSX_PATH = f"emporio_validation_results_{_site_slug}.xlsx"

# Filled by check_accessory_tab_and_count() for the current model.
# A blank/ALL trim attribute means the accessory is model-level and therefore
# available to every trim. This is used only for Accessories-sheet trim coverage.
LAST_MODEL_LEVEL_ALL_ACCESSORY_NAMES = []

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

# Best-effort candidates for whatever opens the full color list/popup.
# The color-section itself (showing the CURRENT color name) definitely
# exists (SEL_DEFAULT_COLOR proves that) - this guesses at a "see all
# colors" trigger near it, following the same container pattern as
# trims/backgrounds/accessories. If colors turn out to always be
# visible without needing a click, get_all_color_names() below still
# works since it searches the whole page, not just inside a popup.
COLOR_BUTTON_CANDIDATES = [
    ("color-section container",
     "izmo-emporio-div.izmo-emporio-visualizer-container > izmo-emporio-div.color-section"),
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

# Accessory card detection - now also captures a display name via
# ".item-name" (same class the trim items use), alongside the
# existing badge-number detection logic.
accessory_cards_script = r"""
function clean(v) {
    return String(v || "")
        .replace(/\s+/g, " ")
        .trim();
}

function badName(v) {
    const x = clean(v);

    return (
        !x ||
        /^Accessory\s+\d+$/i.test(x) ||
        /^[0-9a-f]{8}-[0-9a-f-]{4,}$/i.test(x)
    );
}

function isTargetCard(el) {
    try {
        return (
            el.matches &&
            el.matches(
                ".izmo-emporio-accessory.item.action"
            )
        );
    } catch (e) {
        return false;
    }
}

function walkEveryOpenShadowRoot(root, callback) {
    if (!root) return;

    try {
        const all = root.querySelectorAll
            ? Array.from(
                root.querySelectorAll("*")
            )
            : [];

        for (const el of all) {
            try {
                callback(el);
            } catch (e) {}

            try {
                if (el.shadowRoot) {
                    walkEveryOpenShadowRoot(
                        el.shadowRoot,
                        callback
                    );
                }
            } catch (e) {}
        }
    } catch (e) {}
}

function findAllCards() {
    const cards = [];

    function add(card) {
        if (
            card &&
            !cards.includes(card)
        ) {
            cards.push(card);
        }
    }

    /*
     * IMPORTANT:
     * Search globally through ALL open Shadow DOMs.
     *
     * We do not stop at .izmo-emporio-accessory-section because
     * some Emporio builds put the last accessory cards in another
     * nested custom-element ShadowRoot underneath that section.
     */
    walkEveryOpenShadowRoot(
        document,
        function(el) {
            if (isTargetCard(el)) {
                add(el);
            }
        }
    );

    /*
     * Also query the light DOM directly.
     */
    try {
        for (const card of document.querySelectorAll(
            ".izmo-emporio-accessory.item.action"
        )) {
            add(card);
        }
    } catch (e) {}

    return cards;
}

function getName(card) {
    let nameEl = null;

    try {
        nameEl = card.querySelector(
            ".item-name.fontFamilyRegular.fontSize12"
        );
    } catch (e) {}

    if (!nameEl) {
        try {
            nameEl = card.querySelector(
                ".item-name"
            );
        } catch (e) {}
    }

    /*
     * IMPORTANT:
     * .item-name may itself be inside the card's open Shadow DOM.
     */
    if (!nameEl) {
        try {
            walkEveryOpenShadowRoot(
                card,
                function(el) {
                    if (
                        nameEl ||
                        !el.matches
                    ) {
                        return;
                    }

                    try {
                        if (
                            el.matches(
                                ".item-name.fontFamilyRegular.fontSize12"
                            )
                        ) {
                            nameEl = el;
                            return;
                        }

                        if (
                            el.matches(
                                ".item-name"
                            )
                        ) {
                            nameEl = el;
                        }
                    } catch (e) {}
                }
            );
        } catch (e) {}
    }

    if (!nameEl) {
        return "";
    }

    const value = clean(
        nameEl.innerText ||
        nameEl.textContent ||
        ""
    );

    return badName(value)
        ? ""
        : value;
}

function getNumber(card) {
    const candidates = [];

    function add(value, source, priority) {
        const x = clean(value);

        if (/^\d+$/.test(x)) {
            candidates.push({
                value: parseInt(x, 10),
                source: source,
                priority: priority
            });
        }
    }

    /*
     * Explicit number/count attributes anywhere in this card.
     */
    walkEveryOpenShadowRoot(
        card,
        function(el) {
            try {
                for (const attr of Array.from(
                    el.attributes || []
                )) {
                    const n =
                        String(
                            attr.name || ""
                        ).toLowerCase();

                    if (
                        n.includes("count") ||
                        n.includes("number") ||
                        n.includes("quantity") ||
                        n.includes("total")
                    ) {
                        add(
                            attr.value,
                            "attribute:" + attr.name,
                            100
                        );
                    }
                }
            } catch (e) {}
        }
    );

    /*
     * CSS generated badge.
     */
    walkEveryOpenShadowRoot(
        card,
        function(el) {
            for (const pseudo of [
                "::before",
                "::after"
            ]) {
                try {
                    let value =
                        getComputedStyle(
                            el,
                            pseudo
                        ).content;

                    value = String(
                        value || ""
                    )
                        .replace(
                            /^['"]|['"]$/g,
                            ""
                        )
                        .trim();

                    add(
                        value,
                        "CSS " + pseudo,
                        90
                    );
                } catch (e) {}
            }
        }
    );

    /*
     * Standalone numeric element.
     */
    walkEveryOpenShadowRoot(
        card,
        function(el) {
            try {
                const text = clean(
                    el.innerText ||
                    el.textContent ||
                    ""
                );

                if (!/^\d+$/.test(text)) {
                    return;
                }

                const r =
                    el.getBoundingClientRect();

                add(
                    text,
                    "numeric element",
                    (
                        r.width <= 70 &&
                        r.height <= 70
                    )
                        ? 80
                        : 60
                );
            } catch (e) {}
        }
    );

    if (!candidates.length) {
        return {
            value: null,
            source: ""
        };
    }

    candidates.sort(
        (a, b) => {
            if (
                b.priority !==
                a.priority
            ) {
                return (
                    b.priority -
                    a.priority
                );
            }

            return (
                b.value -
                a.value
            );
        }
    );

    return candidates[0];
}

const cards = findAllCards();
const result = [];

for (
    let index = 0;
    index < cards.length;
    index++
) {
    const card = cards[index];
    const number =
        getNumber(card);

    result.push({
        cardIndex:
            index,
        accessoryid:
            card.getAttribute(
                "accessoryid"
            ) || "",
        trim:
            card.getAttribute(
                "trim"
            ) || "",
        itemName:
            getName(card),
        number:
            number.value,
        numberSource:
            number.source
    });
}

return result;
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




def _try_click_strategies(driver, element):
    """
    Try several click strategies for Emporio custom elements.
    Returns the strategy name that succeeded, otherwise None.
    """
    strategies = [
        (
            "WebElement.click",
            lambda: element.click()
        ),
        (
            "ActionChains.click",
            lambda: ActionChains(driver).move_to_element(element).pause(0.1).click().perform()
        ),
        (
            "JavaScript click",
            lambda: driver.execute_script(
                "arguments[0].click();", element
            )
        ),
        (
            "dispatch MouseEvent",
            lambda: driver.execute_script("""
                arguments[0].dispatchEvent(
                    new MouseEvent("click", {
                        bubbles: true,
                        cancelable: true,
                        view: window
                    })
                );
            """, element)
        ),
    ]

    for name, action in strategies:
        try:
            action()
            time.sleep(0.3)
            return name
        except Exception:
            continue

    return None


def open_popup_robust(driver, candidates):
    """
    Open a popup using the supplied Shadow-DOM-aware selector candidates.
    """
    for label, selector in candidates:
        try:
            element = get_element(driver, selector)
            if element is None:
                continue

            try:
                driver.execute_script("""
                    arguments[0].scrollIntoView({
                        behavior: "instant",
                        block: "center",
                        inline: "center"
                    });
                """, element)
            except Exception:
                pass

            strategy = _try_click_strategies(driver, element)

            if strategy is None:
                continue

            # Give the component a moment to create/render the popup.
            end = time.time() + POPUP_OPEN_TIMEOUT
            while time.time() < end:
                try:
                    if driver.execute_script(
                        verify_any_popup_opened_script
                    ):
                        print(
                            f"POPUP OPENED using '{label}' / '{strategy}'"
                        )
                        return True
                except Exception:
                    pass
                time.sleep(0.25)

        except Exception:
            continue

    return False


def wait_for_popup_item_count(driver, debug_label="POPUP", timeout=POPUP_OPEN_TIMEOUT):
    """
    Wait until a visible .more-items-popUp contains items.
    Returns (count, item dictionaries).
    """
    end_time = time.time() + timeout
    last_result = {"count": 0, "items": []}

    while time.time() < end_time:
        try:
            result = driver.execute_script(
                get_visible_popup_items_script
            ) or {"count": 0, "items": []}

            last_result = result

            if result.get("count", 0) > 0:
                return (
                    int(result.get("count", 0)),
                    result.get("items", []) or []
                )
        except Exception:
            pass

        time.sleep(0.25)

    print(
        f"{debug_label}: visible popup item count = "
        f"{last_result.get('count', 0)}"
    )

    # Diagnostic information is useful when a site changes its markup.
    try:
        print(
            f"{debug_label}: ALL POPUPS:",
            driver.execute_script(debug_all_popups_script)
        )
    except Exception:
        pass

    return (
        int(last_result.get("count", 0)),
        last_result.get("items", []) or []
    )


def check_number_of_trims(driver):
    """
    Open the trim selector, collect every visible trim item, and return:
        (number_of_trims, trim_names)

    The site's visible label is preferred, with the `trim` attribute as
    fallback. Duplicate names are removed case-insensitively while
    preserving page order.
    """
    print("Opening trims popup...")

    opened = open_popup_robust(
        driver,
        TRIM_BUTTON_CANDIDATES
    )

    if not opened:
        print("TRIMS POPUP: failed to open via any strategy")
        return 0, []

    count, items = wait_for_popup_item_count(
        driver,
        debug_label="TRIMS"
    )

    trim_names = []
    seen = set()

    for item in items:
        attr_name = " ".join(
            str(item.get("trim", "") or "").split()
        ).strip()

        text_name = " ".join(
            str(item.get("text", "") or "").split()
        ).strip()

        # Prefer visible text when it is meaningful; otherwise use trim attr.
        name = text_name or attr_name

        # Some builds put additional UI text into innerText. If it does not
        # look useful, use the trim attribute.
        if (
            not name
            or name.casefold() in {
                "trim",
                "trims",
                "select trim"
            }
        ):
            name = attr_name

        if not name:
            continue

        key = name.casefold()
        if key not in seen:
            seen.add(key)
            trim_names.append(name)

    # The actual usable list is more reliable than the raw DOM count if
    # duplicate/stale popup nodes exist.
    usable_count = len(trim_names)

    print(
        f"NUMBER OF TRIMS ({usable_count}):",
        trim_names
    )

    return usable_count, trim_names


def _normalize_trim_value(value):
    return " ".join(str(value or "").split()).strip()


def _find_trim_items_selenium(driver):
    """Return all trim popup items through open Shadow DOM using Selenium."""
    return _find_all_in_open_shadow_dom(
        driver,
        ".more-items-popUp .izmo-emporio-trim.item"
    ) if "_find_all_in_open_shadow_dom" in globals() else []


def _trim_item_matches(element, expected):
    expected_norm = _normalize_trim_value(expected).casefold()
    try:
        attr = _normalize_trim_value(
            element.get_attribute("trim")
        ).casefold()
    except Exception:
        attr = ""
    try:
        text = _normalize_trim_value(
            element.text
        ).casefold()
    except Exception:
        text = ""
    return bool(
        expected_norm and
        (attr == expected_norm or text == expected_norm)
    )


def _is_displayed_safe(element):
    try:
        return element.is_displayed()
    except Exception:
        return False


def _visible_trim_signature(driver):
    """Return the current visible trim names in the open popup."""
    values = []
    try:
        items = _find_all_in_open_shadow_dom(
            driver,
            ".more-items-popUp .izmo-emporio-trim.item"
        )
    except Exception:
        items = []

    for item in items:
        if not _is_displayed_safe(item):
            continue
        try:
            value = _normalize_trim_value(item.get_attribute("trim"))
        except Exception:
            value = ""
        if not value:
            try:
                value = _normalize_trim_value(item.text)
            except Exception:
                value = ""
        if value:
            values.append(value.casefold())

    return tuple(values)


def _wait_for_trim_header(driver, expected, timeout=12):
    """Wait until the vehicle header/base-trim text matches the requested trim."""
    expected_norm = _normalize_trim_value(expected).casefold()
    deadline = time.time() + timeout

    while time.time() < deadline:
        try:
            current = _normalize_trim_value(
                get_text(driver, SEL_BASE_TRIM) or ""
            )
            current_norm = current.casefold()
            if (
                current_norm == expected_norm
                or expected_norm in current_norm
                or current_norm in expected_norm
            ):
                return True
        except Exception:
            pass
        time.sleep(0.25)

    return False


def _click_trim_pagination_next(driver):
    """
    Advance the trim selector by one set and verify that the visible trim set
    really changed.

    Emporio implementations differ: the control may be inside the popup, a
    sibling of it, a custom izmo-emporio-div, a button/anchor, or only expose
    an SVG arrow. Therefore this searches both the popup and its nearby/global
    navigation controls and validates the result by comparing visible trim
    names before/after the click.
    """
    before = _visible_trim_signature(driver)
    if not before:
        return False

    try:
        popups = [p for p in _find_all_in_open_shadow_dom(driver, '.more-items-popUp')
                  if _is_displayed_safe(p)]
    except Exception:
        popups=[]

    candidates=[]
    seen=set()

    def add(el):
        if el is None:
            return
        try:
            key=el.id
        except Exception:
            key=id(el)
        if key not in seen:
            seen.add(key); candidates.append(el)

    selectors=[
        "[aria-label*='next' i]", "[title*='next' i]", "[data-testid*='next' i]",
        "[class*='next' i]", "[aria-label*='right' i]", "[title*='right' i]",
        "[data-testid*='right' i]", "[class*='right' i]",
        "[class*='chevron' i]", "[class*='arrow' i]",
        "button", "a", "[role='button']", "svg", "path"
    ]

    roots=list(popups)
    # Also inspect the visible vehicle/trim popup area globally. Some builds
    # place the carousel arrow outside .more-items-popUp.
    try:
        roots.append(driver)
    except Exception:
        pass

    for root in roots:
        for selector in selectors:
            try:
                elements=root.find_elements(By.CSS_SELECTOR,selector)
            except Exception:
                elements=[]
            for el in elements:
                if not _is_displayed_safe(el):
                    continue
                try:
                    cls=(el.get_attribute('class') or '').lower()
                    aria=(el.get_attribute('aria-label') or '').lower()
                    title=(el.get_attribute('title') or '').lower()
                    testid=(el.get_attribute('data-testid') or '').lower()
                    text=_normalize_trim_value(el.text).lower()
                    sig=' '.join([cls,aria,title,testid,text])
                except Exception:
                    cls=''; sig=''; text=''

                if 'izmo-emporio-trim item' in cls:
                    continue

                likely=(
                    any(x in sig for x in ('next','right','chevron-right','arrow-right'))
                    or text in {'>','>>','›','»','→'}
                )

                # For SVG/path with no attributes, its clickable ancestors can
                # still be the carousel control. Add ancestors, but they are
                # accepted only if clicking changes the visible trim set.
                if likely or selector in {'svg','path'}:
                    add(el)
                    try:
                        parent=el.find_element(By.XPATH,'..')
                        add(parent)
                        grand=parent.find_element(By.XPATH,'..')
                        add(grand)
                        great=grand.find_element(By.XPATH,'..')
                        add(great)
                    except Exception:
                        pass

    # De-duplicate candidate order while trying the most explicit controls first.
    def candidate_score(el):
        try:
            sig=' '.join([
                (el.get_attribute('class') or '').lower(),
                (el.get_attribute('aria-label') or '').lower(),
                (el.get_attribute('title') or '').lower(),
                (el.get_attribute('data-testid') or '').lower(),
                _normalize_trim_value(el.text).lower()
            ])
        except Exception:
            sig=''
        score=0
        for token,pts in [('next',100),('right',90),('chevron',80),('arrow',70),('>',60)]:
            if token in sig: score=max(score,pts)
        return -score

    candidates.sort(key=candidate_score)

    for el in candidates:
        for action in (
            lambda e=el: e.click(),
            lambda e=el: ActionChains(driver).move_to_element(e).pause(0.08).click().perform(),
        ):
            try:
                action()
            except Exception:
                continue
            end=time.time()+3
            while time.time()<end:
                after=_visible_trim_signature(driver)
                if after and after != before:
                    print('TRIM POPUP NEXT: visible set changed')
                    return True
                time.sleep(0.2)

    # Keyboard fallback: carousel implementations often respond to RIGHT even
    # when their arrow has no useful DOM attributes.
    keyboard_targets=popups[:]
    try:
        visible_items=[x for x in _find_all_in_open_shadow_dom(
            driver,'.more-items-popUp .izmo-emporio-trim.item'
        ) if _is_displayed_safe(x)]
        if visible_items:
            keyboard_targets.append(visible_items[-1])
    except Exception:
        pass

    for target in keyboard_targets:
        for key in (Keys.ARROWRIGHT, Keys.PAGEDOWN):
            try:
                target.send_keys(key)
            except Exception:
                try:
                    ActionChains(driver).move_to_element(target).click().send_keys(key).perform()
                except Exception:
                    continue
            end=time.time()+2.5
            while time.time()<end:
                after=_visible_trim_signature(driver)
                if after and after != before:
                    print(f'TRIM POPUP NEXT: keyboard {key} changed visible set')
                    return True
                time.sleep(0.2)

    return False


def _click_exact_trim_element_and_confirm(driver, element, expected):
    """Click a trim item and return True ONLY after the vehicle confirms it."""
    try:
        driver.execute_script(
            "arguments[0].scrollIntoView({block:'center',inline:'center'});",
            element
        )
    except Exception:
        pass

    clicked=False
    try:
        element.click(); clicked=True
    except Exception:
        try:
            ActionChains(driver).move_to_element(element).pause(0.1).click().perform()
            clicked=True
        except Exception:
            pass

    if not clicked:
        return False

    time.sleep(0.8)
    if _wait_for_trim_header(driver, expected, timeout=12):
        print(f"TRIM CONFIRMED IN HEADER: '{expected}'")
        return True

    # Last-resort DOM event, but confirmation is still mandatory.
    try:
        driver.execute_script(
            "arguments[0].dispatchEvent(new MouseEvent('click',{bubbles:true,cancelable:true,view:window}));",
            element
        )
        if _wait_for_trim_header(driver, expected, timeout=8):
            print(f"TRIM CONFIRMED AFTER EVENT: '{expected}'")
            return True
    except Exception:
        pass

    return False


def select_trim_by_name(driver, trim_name, timeout=60):
    """
    Select an exact trim, including trims after the first six-item set.

    The key rule is: NEVER return True merely because a hidden DOM element was
    clicked. The vehicle header must confirm the requested trim. If it does
    not, the code continues paging the trim popup and tries again.
    """
    expected=_normalize_trim_value(trim_name)
    if not expected:
        return False

    # If already selected, don't toggle the dropdown unnecessarily.
    if _wait_for_trim_header(driver, expected, timeout=1.5):
        print(f"TRIM ALREADY SELECTED: '{expected}'")
        return True

    try:
        close_color_popup_if_open(driver)
    except Exception:
        pass
    try:
        close_trim_popup_if_open(driver)
    except Exception:
        pass

    if not open_popup_robust(driver, TRIM_BUTTON_CANDIDATES):
        print(f"TRIM SELECT: could not open popup for '{expected}'")
        return False

    deadline=time.time()+timeout
    pages=0
    max_pages=20
    seen_page_signatures=set()

    while time.time()<deadline and pages<=max_pages:
        try:
            items=_find_all_in_open_shadow_dom(
                driver,'.more-items-popUp .izmo-emporio-trim.item'
            )
        except Exception:
            items=[]

        visible_matches=[]
        hidden_matches=[]
        for item in items:
            if not _trim_item_matches(item, expected):
                continue
            if _is_displayed_safe(item):
                visible_matches.append(item)
            else:
                hidden_matches.append(item)

        # First choice: visible exact match.
        for item in visible_matches:
            if _click_exact_trim_element_and_confirm(driver,item,expected):
                return True

        # Record the current page so a broken next control cannot loop forever.
        signature=_visible_trim_signature(driver)
        if signature:
            if signature in seen_page_signatures and pages>0:
                # We are cycling; stop trying the same page.
                pass
            else:
                seen_page_signatures.add(signature)

        # Advance to the next set before using a hidden-element fallback.
        if pages < max_pages and _click_trim_pagination_next(driver):
            pages += 1
            time.sleep(0.5)
            continue

        # If the matching item exists hidden in the DOM, try it, but REQUIRE
        # header confirmation. This supports builds where all pages are in the
        # DOM but only six are visually displayed.
        for item in hidden_matches:
            if _click_exact_trim_element_and_confirm(driver,item,expected):
                return True

        # Re-open the popup if a failed click closed it, then continue.
        try:
            close_trim_popup_if_open(driver)
            if time.time()<deadline:
                open_popup_robust(driver,TRIM_BUTTON_CANDIDATES)
        except Exception:
            pass
        time.sleep(0.3)

    print(f"TRIM SELECT FAILED AFTER PAGING: '{expected}'")
    return False


def _get_shadow_root(element):
    """Return Selenium 4's open ShadowRoot, or None."""
    try:
        return element.shadow_root
    except Exception:
        return None


def _find_all_in_open_shadow_dom(driver, selector):
    """
    Selenium-only recursive search through light DOM and all open Shadow DOMs.
    No JavaScript DOM querying is used here.
    """
    results = []
    seen_elements = set()
    visited_roots = set()

    def add_unique(element):
        if element is None:
            return
        try:
            key = element.id
        except Exception:
            key = id(element)
        if key not in seen_elements:
            seen_elements.add(key)
            results.append(element)

    def walk(root):
        if root is None:
            return
        try:
            root_key = id(root)
            if root_key in visited_roots:
                return
            visited_roots.add(root_key)
        except Exception:
            pass

        try:
            for element in root.find_elements(By.CSS_SELECTOR, selector):
                add_unique(element)
        except Exception:
            pass

        try:
            all_elements = root.find_elements(By.CSS_SELECTOR, "*")
        except Exception:
            all_elements = []

        for element in all_elements:
            shadow = _get_shadow_root(element)
            if shadow is not None:
                walk(shadow)

    walk(driver)
    return results


def _clean_color_name(value):
    return " ".join(str(value or "").split()).strip()


def _unique_color_names(names):
    unique = []
    seen = set()
    for raw_name in names:
        name = _clean_color_name(raw_name)
        if not name:
            continue
        key = name.casefold()
        if key not in seen:
            seen.add(key)
            unique.append(name)
    return unique


def _safe_displayed(element):
    try:
        return element.is_displayed()
    except Exception:
        return False


def _get_current_color_sections(driver):
    """
    Re-discover the current color section on every call.

    Do NOT reject a section merely because the custom-element wrapper reports
    is_displayed() == False. Emporio sometimes keeps the wrapper's geometry at
    zero while its child swatches are usable. Prefer the section with the most
    live color evidence.
    """
    sections=_find_all_in_open_shadow_dom(driver,'.color-section')
    scored=[]

    for section in sections:
        try:
            color_elements=section.find_elements(By.CSS_SELECTOR,'[data-color-name]')
        except Exception:
            color_elements=[]
        try:
            more=section.find_elements(By.CSS_SELECTOR,'.moreColors.action > .moreColorsImg')
        except Exception:
            more=[]

        visible_color_count=0
        for el in color_elements:
            if _safe_displayed(el):
                visible_color_count+=1

        try:
            section_visible=_safe_displayed(section)
        except Exception:
            section_visible=False

        score=(visible_color_count*20) + min(len(color_elements),20) + (100 if more else 0)
        if section_visible:
            score+=5

        if color_elements or more:
            scored.append((score,section))

    scored.sort(key=lambda x:x[0],reverse=True)
    return [section for _,section in scored]


def _get_color_name_from_swatch(swatch):
    try:
        elements = swatch.find_elements(
            By.CSS_SELECTOR,
            "[data-color-name]"
        )
    except Exception:
        elements = []

    if not elements:
        return ""

    try:
        return _clean_color_name(
            elements[0].get_attribute("data-color-name")
        )
    except Exception:
        return ""


def _collect_popup_colors_selenium(driver):
    """
    Read all expanded More Colors swatches:
        .color-container -> .color-swatch -> [data-color-name]

    Visibility is intentionally NOT checked on popup swatches.
    """
    colors = []

    containers = _find_all_in_open_shadow_dom(
        driver,
        ".color-container"
    )

    for container in containers:
        try:
            swatches = container.find_elements(
                By.CSS_SELECTOR,
                ".color-swatch"
            )
        except Exception:
            swatches = []

        for swatch in swatches:
            name = _get_color_name_from_swatch(swatch)
            if name:
                colors.append(name)

        if not swatches:
            try:
                elements = container.find_elements(
                    By.CSS_SELECTOR,
                    "[data-color-name]"
                )
            except Exception:
                elements = []

            for element in elements:
                try:
                    name = _clean_color_name(
                        element.get_attribute("data-color-name")
                    )
                except Exception:
                    name = ""
                if name:
                    colors.append(name)

    return _unique_color_names(colors)


def _collect_visible_colors_from_current_section(driver, section):
    """Read visible data-color-name elements from one current color section."""
    colors = []
    try:
        elements = section.find_elements(
            By.CSS_SELECTOR,
            "[data-color-name]"
        )
    except Exception:
        elements = []

    for element in elements:
        try:
            if not _safe_displayed(element):
                continue
            name = _clean_color_name(
                element.get_attribute("data-color-name")
            )
        except Exception:
            name = ""
        if name:
            colors.append(name)

    return _unique_color_names(colors)


def _click_more_colors_in_current_section(driver, section):
    """Click the More Colors parent in the current color section."""
    try:
        images = section.find_elements(
            By.CSS_SELECTOR,
            ".moreColors.action > .moreColorsImg"
        )
    except Exception:
        images = []

    for image in images:
        try:
            parent = image.find_element(By.XPATH, "..")
        except Exception:
            continue

        try:
            if not parent.is_displayed():
                continue
        except Exception:
            pass

        try:
            driver.execute_script(
                "arguments[0].scrollIntoView({block:'center',inline:'center'});",
                parent
            )
        except Exception:
            pass

        try:
            parent.click()
            print("  MORE COLORS: Selenium click succeeded")
            return True
        except Exception:
            try:
                ActionChains(driver).move_to_element(parent).pause(0.1).click().perform()
                print("  MORE COLORS: ActionChains click succeeded")
                return True
            except Exception:
                pass

    return False


def get_all_color_names(driver, timeout=35, previous_trim=None):
    """
    Collect colors for the currently selected trim.

    More Colors path:
      current .color-section -> .moreColors.action -> click ->
      .color-container -> .color-swatch -> [data-color-name]

    Without More Colors:
      current .color-section -> [data-color-name]

    Every lookup is re-discovered from the live DOM. No color WebElement from
    a previous trim is reused. This is important for Emporio's six-trim UI,
    where the component is replaced/refreshed when moving to the next set.
    """
    deadline=time.time()+timeout
    last_colors=[]

    while time.time()<deadline:
        sections=_get_current_color_sections(driver)
        if not sections:
            time.sleep(0.5)
            continue

        section=sections[0]

        try:
            more_images=section.find_elements(
                By.CSS_SELECTOR,'.moreColors.action > .moreColorsImg'
            )
        except Exception:
            more_images=[]

        if more_images:
            print('  COLOR: MORE COLORS FOUND IN CURRENT SECTION')
            try:
                close_color_popup_if_open(driver)
            except Exception:
                pass

            if _click_more_colors_in_current_section(driver,section):
                popup_deadline=min(deadline,time.time()+20)
                best=[]
                stable=0
                previous=None
                while time.time()<popup_deadline:
                    colors=_collect_popup_colors_selenium(driver)
                    if colors:
                        best=colors
                        key=tuple(x.casefold() for x in colors)
                        if key==previous:
                            stable+=1
                        else:
                            previous=key
                            stable=1
                        if stable>=2:
                            print(f'  MORE COLORS: FINAL COUNT = {len(best)}')
                            return best
                    time.sleep(0.4)
                if best:
                    print(f'  MORE COLORS: FINAL COUNT = {len(best)}')
                    return best

        # Direct current-section fallback.
        colors=_collect_visible_colors_from_current_section(driver,section)
        if colors:
            print(f'  VISIBLE COLORS: COUNT = {len(colors)}')
            return colors
        last_colors=colors
        time.sleep(0.5)

    print(f'  COLOR: timeout -> returning {len(last_colors)} colors')
    return last_colors


def close_trim_popup_if_open(driver):
    """Close the trim popup if it is still visible."""
    try:
        is_open = driver.execute_script(r"""
        function visible(el) {
            if (!el) return false;
            try {
                const s = getComputedStyle(el);
                const r = el.getBoundingClientRect();
                return (
                    !el.classList.contains("hide") &&
                    s.display !== "none" &&
                    s.visibility !== "hidden" &&
                    r.width > 0 &&
                    r.height > 0
                );
            } catch (e) {
                return false;
            }
        }

        function find(root) {
            if (!root) return false;

            try {
                const popups = root.querySelectorAll
                    ? root.querySelectorAll(".more-items-popUp")
                    : [];

                for (const popup of popups) {
                    if (visible(popup)) return true;
                }
            } catch (e) {}

            try {
                const elements = root.querySelectorAll
                    ? root.querySelectorAll("*")
                    : [];

                for (const el of elements) {
                    if (el.shadowRoot && find(el.shadowRoot)) {
                        return true;
                    }
                }
            } catch (e) {}

            return false;
        }

        return find(document);
        """)

        if is_open:
            element = get_element(
                driver,
                TRIM_BUTTON_CANDIDATES[0][1]
            )
            if element is not None:
                _try_click_strategies(
                    driver,
                    element
                )
                time.sleep(0.3)
    except Exception:
        pass


def close_color_popup_if_open(driver):
    """
    Close the More Colors popup before moving to another trim.
    Searches light DOM and open Shadow DOM.
    """
    script = r"""
    function visible(el) {
        if (!el) return false;

        try {
            const s = getComputedStyle(el);
            const r = el.getBoundingClientRect();

            return (
                s.display !== "none" &&
                s.visibility !== "hidden" &&
                s.opacity !== "0" &&
                r.width > 0 &&
                r.height > 0
            );
        } catch (e) {
            return false;
        }
    }

    function clean(v) {
        return String(v || "")
            .replace(/\s+/g, " ")
            .trim();
    }

    function findAndClick(root) {
        if (!root) return false;

        try {
            const elements = root.querySelectorAll
                ? root.querySelectorAll("*")
                : [];

            /*
             * First look for an exact visible CANCEL control.
             */
            for (const el of elements) {

                if (!visible(el)) continue;

                const text = clean(
                    el.innerText ||
                    el.textContent ||
                    ""
                );

                if (
                    text.toUpperCase() === "CANCEL"
                ) {
                    try {
                        el.click();
                        return true;
                    } catch (e) {
                        try {
                            el.dispatchEvent(
                                new MouseEvent("click", {
                                    bubbles: true,
                                    cancelable: true,
                                    view: window
                                })
                            );
                            return true;
                        } catch (e2) {}
                    }
                }
            }
        } catch (e) {}

        /*
         * Then search nested Shadow DOM.
         */
        try {
            const elements = root.querySelectorAll
                ? root.querySelectorAll("*")
                : [];

            for (const el of elements) {
                if (
                    el.shadowRoot &&
                    findAndClick(el.shadowRoot)
                ) {
                    return true;
                }
            }
        } catch (e) {}

        return false;
    }

    return findAndClick(document);
    """

    try:
        result = driver.execute_script(script)

        if result:
            time.sleep(0.5)
            print("MORE COLORS: CLOSED")
            return True

    except Exception:
        pass

    return False



def _unique_case_insensitive_strings(values):
    result=[]
    seen=set()
    for raw in values or []:
        value=" ".join(str(raw or '').split()).strip()
        if not value:
            continue
        key=value.casefold()
        if key not in seen:
            seen.add(key)
            result.append(value)
    return result


def _merge_unique_preserve_order(first, second):
    return _unique_case_insensitive_strings(
        list(first or []) + list(second or [])
    )


def collect_accessories_and_colors_for_each_trim(
    driver,
    trim_names,
    initial_accessory_names=None,
    initial_all_trim_accessory_names=None
):
    """
    Collect Colors and Accessories per trim.

    ONE TRIM:
        - Do NOT click the trim.
        - The page is already on the only trim.
        - STILL collect all Colors.
        - STILL include ALL Accessories.
        - Reuse the complete accessory list collected by the initial
          Validation Results accessory check. This avoids toggling an
          already-open Accessories panel a second time.

    MULTIPLE TRIMS:
        - Select each trim.
        - Wait for that trim.
        - Collect Colors.
        - Open Accessories and collect all names.
    """
    colors_by_trim = {}
    accessories_by_trim = {}
    accessory_counts_by_trim = {}

    ordered_trims = []
    seen = set()

    for raw_trim in trim_names or []:
        trim = " ".join(
            str(raw_trim or "").split()
        ).strip()

        key = trim.casefold()

        if trim and key not in seen:
            seen.add(key)
            ordered_trims.append(trim)

    trim_count = len(ordered_trims)

    print("")
    print("=" * 70)
    print("STARTING EXACT PER-TRIM COLLECTION")
    print("=" * 70)
    print(
        "TOTAL TRIMS:",
        trim_count
    )
    print(
        "TRIMS:",
        ordered_trims
    )

    if not ordered_trims:
        return (
            colors_by_trim,
            accessories_by_trim,
            accessory_counts_by_trim
        )

    initial_all_trim_accessory_names = _unique_case_insensitive_strings(
        initial_all_trim_accessory_names or []
    )

    initial_accessory_names = [
        " ".join(
            str(x or "").split()
        ).strip()
        for x in (
            initial_accessory_names or []
        )
        if " ".join(
            str(x or "").split()
        ).strip()
    ]

    for index, trim in enumerate(
        ordered_trims,
        start=1
    ):
        print("")
        print("#" * 70)
        print(
            f"TRIM {index}/{trim_count}: {trim}"
        )
        print("#" * 70)

        colors_by_trim[trim] = []
        accessories_by_trim[trim] = []
        accessory_counts_by_trim[trim] = 0

        # ------------------------------------------------------
        # Trim selection
        # ------------------------------------------------------
        try:
            close_color_popup_if_open(driver)
        except Exception:
            pass

        if trim_count == 1:
            # IMPORTANT:
            # Do not click the only trim.
            print(
                f"[{trim}] ONE TRIM -> "
                "TRIM CLICK SKIPPED"
            )

            print(
                f"[{trim}] WAITING 5 SECONDS "
                "FOR CURRENT TRIM CONTENT..."
            )
            time.sleep(5)

        else:
            print(
                f"[{trim}] SELECTING TRIM..."
            )

            if not select_trim_by_name(
                driver,
                trim
            ):
                print(
                    f"[{trim}] FAILED TO SELECT TRIM"
                )
                continue

            print(
                f"[{trim}] TRIM SELECTED"
            )

            # Do not start color/accessory extraction until the vehicle
            # component confirms this exact trim. This prevents the first
            # six-trim batch from being reused for trim 7+.
            if not _wait_for_trim_header(driver, trim, timeout=20):
                print(f"[{trim}] WARNING: header confirmation not observed; forcing refresh wait")
            else:
                print(f"[{trim}] HEADER CONFIRMED")

            print(
                f"[{trim}] WAITING 5 SECONDS "
                "FOR TRIM CONTENT..."
            )
            time.sleep(5)

        # ======================================================
        # COLORS - ALWAYS RUN, INCLUDING SINGLE-TRIM MODELS
        # ======================================================
        print(
            f"[{trim}] FETCHING COLORS..."
        )

        try:
            close_color_popup_if_open(driver)
        except Exception:
            pass

        trim_colors = []

        for attempt in range(1, 4):
            try:
                trim_colors = (
                    get_all_color_names(
                        driver,
                        timeout=35,
                        previous_trim=trim
                    )
                    or []
                )
            except Exception as e:
                print(
                    f"[{trim}] COLOR ERROR "
                    f"ATTEMPT {attempt}: {e}"
                )
                trim_colors = []

            if trim_colors:
                break

            print(
                f"[{trim}] COLOR COUNT = 0 "
                f"ON ATTEMPT {attempt}/3"
            )

            try:
                close_color_popup_if_open(driver)
            except Exception:
                pass

            time.sleep(2)

        colors_by_trim[trim] = trim_colors

        print(
            f"[{trim}] COLOR COUNT = "
            f"{len(trim_colors)}"
        )
        print(
            f"[{trim}] COLOR NAMES = "
            f"{', '.join(trim_colors)}"
        )

        try:
            close_color_popup_if_open(driver)
        except Exception:
            pass

        time.sleep(1)

        # ======================================================
        # ACCESSORIES
        # ======================================================
        if trim_count == 1:
            # --------------------------------------------------
            # SINGLE TRIM:
            # The initial Validation Results check already opened
            # Accessories on this exact selected trim and collected
            # the full accessory names.
            #
            # Reuse them instead of clicking the tab a second time.
            # --------------------------------------------------
            print(
                f"[{trim}] SINGLE TRIM -> "
                "USING COMPLETE INITIAL ACCESSORY LIST"
            )

            accessories_by_trim[trim] = list(
                initial_accessory_names
            )

            accessory_counts_by_trim[trim] = (
                len(initial_accessory_names)
            )

            print(
                f"[{trim}] ACCESSORY NAME ROWS = "
                f"{len(accessories_by_trim[trim])}"
            )

            print(
                f"[{trim}] ACCESSORY NAMES = "
                f"{accessories_by_trim[trim]}"
            )

        else:
            # --------------------------------------------------
            # MULTI TRIM:
            # Explicitly collect Accessories for this trim.
            # --------------------------------------------------
            print(
                f"[{trim}] FETCHING ACCESSORIES..."
            )

            try:
                (
                    accessory_status,
                    accessory_count,
                    trim_accessories
                ) = check_accessory_tab_and_count(
                    driver
                )

                accessories_by_trim[trim] = (
                    trim_accessories or []
                )

                accessory_counts_by_trim[trim] = (
                    accessory_count or 0
                )

            except Exception as e:
                print(
                    f"[{trim}] ACCESSORY ERROR: {e}"
                )

                accessories_by_trim[trim] = []
                accessory_counts_by_trim[trim] = 0

            # Model-level accessories (trim attribute blank/ALL) are by
            # definition available to every trim. Add them to every trim
            # without changing the Validation Results model-level count.
            accessories_by_trim[trim] = _merge_unique_preserve_order(
                initial_all_trim_accessory_names,
                accessories_by_trim.get(trim, [])
            )

            print(
                f"[{trim}] ACCESSORY COUNT = "
                f"{accessory_counts_by_trim[trim]}"
            )

            print(
                f"[{trim}] ACCESSORY NAME ROWS = "
                f"{len(accessories_by_trim[trim])}"
            )

            print(
                f"[{trim}] ACCESSORY NAMES = "
                f"{accessories_by_trim[trim]}"
            )

            # Close the tab before the next trim.
            try:
                accessory_tab = get_element(
                    driver,
                    SEL_ACCESSORY_TAB
                )

                if accessory_tab is not None:
                    _try_click_strategies(
                        driver,
                        accessory_tab
                    )
                    time.sleep(0.8)

            except Exception:
                pass

    print("")
    print("=" * 70)
    print("EXACT PER-TRIM COLLECTION FINISHED")
    print("=" * 70)

    for trim in ordered_trims:
        print(
            f"{trim}: "
            f"{len(colors_by_trim.get(trim, []))} colors, "
            f"{len(accessories_by_trim.get(trim, []))} "
            "accessory names"
        )

    return (
        colors_by_trim,
        accessories_by_trim,
        accessory_counts_by_trim
    )



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



def _extract_visible_accessory_names_from_open_popup(driver):
    """
    After a numbered accessory card is clicked, Emporio can display the
    remaining accessory names in a popup/expanded panel.

    Search the complete visible DOM, including OPEN shadow roots, and
    return meaningful accessory-name strings.
    """
    script = r"""
    function visible(el) {
        if (!el) return false;

        try {
            const s = getComputedStyle(el);

            if (
                s.display === "none" ||
                s.visibility === "hidden" ||
                s.opacity === "0"
            ) {
                return false;
            }

            const r = el.getBoundingClientRect();

            return r.width > 0 && r.height > 0;
        } catch (e) {
            return false;
        }
    }

    function clean(v) {
        return String(v || "")
            .replace(/\s+/g, " ")
            .trim();
    }

    function addUnique(list, value) {
        value = clean(value);

        if (!value) return;

        if (
            /^\d+$/.test(value) ||
            value.toLowerCase() === "accessory" ||
            value.toLowerCase() === "accessories"
        ) {
            return;
        }

        if (!list.includes(value)) {
            list.push(value);
        }
    }

    function walk(root, callback) {
        if (!root) return;

        let children = [];

        try {
            children = root.querySelectorAll
                ? Array.from(root.querySelectorAll("*"))
                : [];
        } catch (e) {}

        for (const el of children) {

            callback(el);

            try {
                if (el.shadowRoot) {
                    walk(el.shadowRoot, callback);
                }
            } catch (e) {}
        }

        try {
            if (root.shadowRoot) {
                walk(root.shadowRoot, callback);
            }
        } catch (e) {}
    }

    const names = [];

    /*
     * Prefer the popup/expanded accessory area.
     * These selectors are based on the Emporio structure already
     * present in the validation script.
     */
    const popupSelectors = [
        ".more-items-popUp",
        ".more-items-popup",
        ".panel-popup",
        ".panel-popup_sidebar",
        ".panel-popup-sidebar",
        ".accessory-popup",
        ".accessories-popup"
    ];

    const popupRoots = [];

    function collectRoots(root) {
        if (!root) return;

        try {
            for (const selector of popupSelectors) {
                const found = root.querySelectorAll
                    ? root.querySelectorAll(selector)
                    : [];

                for (const el of found) {
                    if (visible(el)) {
                        popupRoots.push(el);
                    }
                }
            }
        } catch (e) {}

        try {
            const elements = root.querySelectorAll
                ? root.querySelectorAll("*")
                : [];

            for (const el of elements) {
                if (el.shadowRoot) {
                    collectRoots(el.shadowRoot);
                }
            }
        } catch (e) {}
    }

    collectRoots(document);

    /*
     * If a popup was found, inspect it first.
     * Otherwise inspect the whole visible page as a fallback.
     */
    const rootsToInspect =
        popupRoots.length
            ? popupRoots
            : [document];

    for (const root of rootsToInspect) {

        walk(root, function(el) {

            if (!visible(el)) return;

            let cls = "";

            try {
                cls = typeof el.className === "string"
                    ? el.className.toLowerCase()
                    : "";
            } catch (e) {}

            const tag = (
                el.tagName || ""
            ).toLowerCase();

            const text = clean(
                el.innerText || el.textContent || ""
            );

            if (!text || text.length > 120) {
                return;
            }

            /*
             * Strong selectors for accessory names.
             */
            const strongName =
                cls.includes("item-name") ||
                cls.includes("accessory-name") ||
                cls.includes("accessoryname") ||
                cls.includes("more-item") ||
                cls.includes("accessory-item");

            if (strongName) {
                addUnique(names, text);
                return;
            }

            /*
             * Emporio custom elements often hold the displayed
             * accessory label directly in an izmo-emporio-div.
             */
            if (
                tag === "izmo-emporio-div" &&
                text &&
                text.length <= 100
            ) {
                /*
                 * Avoid adding parent containers whose text is a
                 * combination of several child labels.
                 */
                const childName = el.querySelector
                    ? el.querySelector(
                        ".item-name, .accessory-name, .accessoryname"
                    )
                    : null;

                if (!childName) {
                    addUnique(names, text);
                }
            }

            /*
             * Name-bearing attributes.
             */
            try {
                for (const attr of [
                    "data-name",
                    "data-accessory-name",
                    "accessoryname",
                    "title",
                    "aria-label"
                ]) {
                    const value = clean(
                        el.getAttribute(attr)
                    );

                    if (
                        value &&
                        !/^\d+$/.test(value) &&
                        value.length <= 100
                    ) {
                        addUnique(names, value);
                    }
                }
            } catch (e) {}
        });
    }

    return names;
    """

    try:
        return driver.execute_script(script) or []
    except Exception:
        return []


def _click_numbered_accessory_card_and_get_names(
    driver,
    card_index,
    accessory_id="",
    main_name=""
):
    """
    Click ONE numbered accessory card and fetch its child accessory names.

    IMPORTANT:
      - The parent numbered card is identified by accessoryid whenever
        possible, not only by DOM index.
      - Child names are read from the exact Emporio class:
            .item-name.fontFamilyRegular.fontSize12
      - One actual child accessory may render its name in more than one
        nested element. Those nested presentation elements are collapsed
        using the child's accessoryid (preferred), image source, element id,
        or a positional fallback.
      - Same name + DIFFERENT accessoryid = different accessories.
      - Same name + SAME accessoryid = one logical accessory.
      - Child order is preserved.
    """

    click_script = r"""
    function clean(v) {
        return String(v || "").replace(/\s+/g, " ").trim();
    }

    function walk(root, callback) {
        if (!root) return;
        try {
            const all = root.querySelectorAll
                ? Array.from(root.querySelectorAll("*"))
                : [];

            for (const el of all) {
                callback(el);
                try {
                    if (el.shadowRoot) {
                        walk(el.shadowRoot, callback);
                    }
                } catch (e) {}
            }
        } catch (e) {}
    }

    function getCards() {
        const cards = [];

        function add(el) {
            if (el && !cards.includes(el)) {
                cards.push(el);
            }
        }

        walk(document, function(el) {
            try {
                if (
                    el.matches &&
                    el.matches(".izmo-emporio-accessory.item.action")
                ) {
                    add(el);
                }
            } catch (e) {}
        });

        try {
            for (const el of document.querySelectorAll(
                ".izmo-emporio-accessory.item.action"
            )) {
                add(el);
            }
        } catch (e) {}

        return cards;
    }

    const cards = getCards();
    const requestedId = clean(arguments[1]);
    const requestedName = clean(arguments[2]);
    const requestedIndex = Number(arguments[0]);

    let card = null;

    /*
     * FIRST CHOICE: exact accessoryid.
     */
    if (requestedId) {
        for (const candidate of cards) {
            try {
                if (
                    clean(candidate.getAttribute("accessoryid")) ===
                    requestedId
                ) {
                    card = candidate;
                    break;
                }
            } catch (e) {}
        }
    }

    /*
     * SECOND CHOICE: exact main-card name.
     */
    if (!card && requestedName) {
        for (const candidate of cards) {
            try {
                const nameEl =
                    candidate.querySelector(
                        ".item-name.fontFamilyRegular.fontSize12"
                    ) ||
                    candidate.querySelector(".item-name");

                const name = clean(
                    nameEl
                        ? (nameEl.innerText || nameEl.textContent || "")
                        : ""
                );

                if (
                    name &&
                    name.toUpperCase() === requestedName.toUpperCase()
                ) {
                    card = candidate;
                    break;
                }
            } catch (e) {}
        }
    }

    /*
     * FINAL FALLBACK: original global index.
     */
    if (!card && cards[requestedIndex]) {
        card = cards[requestedIndex];
    }

    if (!card) {
        return {
            clicked: false,
            id: "",
            trim: ""
        };
    }

    try {
        card.scrollIntoView({
            behavior: "instant",
            block: "center",
            inline: "center"
        });
    } catch (e) {}

    let clicked = false;

    try {
        card.click();
        clicked = true;
    } catch (e) {}

    if (!clicked) {
        try {
            card.dispatchEvent(
                new MouseEvent("click", {
                    bubbles: true,
                    cancelable: true,
                    view: window
                })
            );
            clicked = true;
        } catch (e) {}
    }

    return {
        clicked: clicked,
        id: clean(card.getAttribute("accessoryid")),
        trim: clean(card.getAttribute("trim"))
    };
    """

    try:
        info = driver.execute_script(
            click_script,
            card_index,
            accessory_id or "",
            main_name or ""
        ) or {}
    except Exception as e:
        print("  NUMBERED CARD CLICK ERROR:", e)
        return []

    if not info.get("clicked"):
        return []

    # Allow the popup to render before reading it.
    time.sleep(5)

    extract_script = r"""
    function clean(v) {
        return String(v || "").replace(/\s+/g, " ").trim();
    }

    function validName(v) {
        const x = clean(v);
        return (
            x &&
            !/^Accessory\s+\d+$/i.test(x) &&
            !/^[0-9a-f]{8}-[0-9a-f-]{4,}$/i.test(x)
        );
    }

    function visible(el) {
        if (!el) return false;

        try {
            if (el.classList && el.classList.contains("hide")) {
                return false;
            }

            const s = getComputedStyle(el);

            return (
                s.display !== "none" &&
                s.visibility !== "hidden" &&
                s.opacity !== "0"
            );
        } catch (e) {
            return true;
        }
    }

    function getAttr(el, names) {
        if (!el) return "";

        for (const name of names) {
            try {
                const value = clean(el.getAttribute(name));
                if (value) return value;
            } catch (e) {}
        }

        return "";
    }

    /*
     * Find the stable identity belonging to the ACTUAL accessory item.
     * The .item-name element itself often has no accessoryid, so climb
     * its ancestors until accessoryid/image/id is found.
     */
    function getNearestAccessoryInfo(el) {
        let current = el;

        for (let level = 0; level < 15 && current; level++) {
            try {
                const accessoryId = getAttr(current, [
                    "accessoryid",
                    "accessoryId",
                    "data-accessoryid",
                    "data-accessory-id"
                ]);

                if (accessoryId) {
                    return {
                        accessoryid: accessoryId,
                        elementId: clean(current.id),
                        image: ""
                    };
                }

                const image = getAttr(current, [
                    "src",
                    "data-src",
                    "data-image",
                    "data-image-src",
                    "image",
                    "data-image-url"
                ]);

                if (image) {
                    return {
                        accessoryid: "",
                        elementId: clean(current.id),
                        image: image
                    };
                }

                if (current.id) {
                    const id = clean(current.id);
                    if (
                        id &&
                        !/^izmo[a-z0-9-]+$/i.test(id)
                    ) {
                        return {
                            accessoryid: "",
                            elementId: id,
                            image: ""
                        };
                    }
                }
            } catch (e) {}

            current = current.parentElement;
        }

        return {
            accessoryid: "",
            elementId: "",
            image: ""
        };
    }

    function buildKey(name, info, index) {
        const n = clean(name).toUpperCase();

        if (info.accessoryid) {
            return "ID|" + info.accessoryid.toUpperCase() + "|" + n;
        }

        if (info.image) {
            return "IMG|" + info.image + "|" + n;
        }

        if (info.elementId) {
            return "EID|" + info.elementId + "|" + n;
        }

        /*
         * No stable identifier exists. Position is used only as a final
         * fallback so two real same-name items are not collapsed merely
         * because their names match.
         */
        return "POS|" + index + "|" + n;
    }

    const targetId = clean(arguments[0]);

    function collectNamesFromRoot(root, output) {
        if (!root) return;

        try {
            const elements = root.querySelectorAll
                ? Array.from(
                    root.querySelectorAll(
                        ".item-name.fontFamilyRegular.fontSize12"
                    )
                )
                : [];

            for (const el of elements) {
                if (!visible(el)) continue;

                const name = clean(
                    el.innerText || el.textContent || ""
                );

                if (!validName(name)) continue;

                const info = getNearestAccessoryInfo(el);

                // IMPORTANT: The accessory popup also contains trim labels
                // such as S, X-Dynamic SE, Outbound, X and V8.  Those are
                // NOT accessories.  A real popup accessory has its own
                // accessoryid on an ancestor.  Only accept names whose
                // nearest stable identity is an accessoryid, and exclude
                // the parent accessory itself.
                if (!info.accessoryid) continue;
                if (info.accessoryid.toUpperCase() === targetId.toUpperCase()) continue;

                output.push({
                    name: name,
                    accessoryid: info.accessoryid || "",
                    elementId: info.elementId || "",
                    image: info.image || "",
                    key: buildKey(name, info, output.length)
                });
            }
        } catch (e) {}

        /*
         * Some builds may only expose ".item-name". Do this second so the
         * exact class above remains the primary source.
         */
        try {
            const elements = root.querySelectorAll
                ? Array.from(root.querySelectorAll(".item-name"))
                : [];

            for (const el of elements) {
                if (!visible(el)) continue;

                const name = clean(
                    el.innerText || el.textContent || ""
                );

                if (!validName(name)) continue;

                const info = getNearestAccessoryInfo(el);

                // IMPORTANT: The accessory popup also contains trim labels
                // such as S, X-Dynamic SE, Outbound, X and V8.  Those are
                // NOT accessories.  A real popup accessory has its own
                // accessoryid on an ancestor.  Only accept names whose
                // nearest stable identity is an accessoryid, and exclude
                // the parent accessory itself.
                if (!info.accessoryid) continue;
                if (info.accessoryid.toUpperCase() === targetId.toUpperCase()) continue;

                output.push({
                    name: name,
                    accessoryid: info.accessoryid || "",
                    elementId: info.elementId || "",
                    image: info.image || "",
                    key: buildKey(name, info, output.length)
                });
            }
        } catch (e) {}

        /*
         * Enter nested open Shadow DOMs.
         */
        try {
            const all = root.querySelectorAll
                ? Array.from(root.querySelectorAll("*"))
                : [];

            for (const el of all) {
                if (el.shadowRoot) {
                    collectNamesFromRoot(el.shadowRoot, output);
                }
            }
        } catch (e) {}
    }

    function findTargetCard(root, targetId) {
        if (!root) return null;

        try {
            const cards = root.querySelectorAll
                ? Array.from(
                    root.querySelectorAll(
                        ".izmo-emporio-accessory.item.action[accessoryid]"
                    )
                )
                : [];

            for (const card of cards) {
                if (
                    clean(card.getAttribute("accessoryid")) ===
                    targetId
                ) {
                    return card;
                }
            }
        } catch (e) {}

        try {
            const all = root.querySelectorAll
                ? Array.from(root.querySelectorAll("*"))
                : [];

            for (const el of all) {
                if (el.shadowRoot) {
                    const found = findTargetCard(
                        el.shadowRoot,
                        targetId
                    );
                    if (found) return found;
                }
            }
        } catch (e) {}

        return null;
    }

    function findVisiblePopups(root, output) {
        if (!root) return;

        try {
            const popups = root.querySelectorAll
                ? Array.from(
                    root.querySelectorAll(".more-items-popUp")
                )
                : [];

            for (const popup of popups) {
                if (!visible(popup)) continue;
                collectNamesFromRoot(popup, output);
            }
        } catch (e) {}

        try {
            const all = root.querySelectorAll
                ? Array.from(root.querySelectorAll("*"))
                : [];

            for (const el of all) {
                if (el.shadowRoot) {
                    findVisiblePopups(el.shadowRoot, output);
                }
            }
        } catch (e) {}
    }

    const output = [];

    /*
     * 1. Prefer the clicked parent's own popup/container.
     */
    const targetCard = findTargetCard(document, targetId);

    if (targetCard) {
        /*
         * The numbered popup is rendered as a .more-items-popUp.  Search
         * the visible popup(s), but only accept child names carrying their
         * own accessoryid. This prevents trim labels inside the popup from
         * being returned as accessories.
         */
        findVisiblePopups(document, output);
    } else {
        findVisiblePopups(document, output);
    }

    /*
     * ID-aware de-duplication.
     *
     * SAME name + SAME accessoryid -> duplicate DOM presentation -> remove.
     * SAME name + DIFFERENT accessoryid -> separate logical accessories.
     */
    const result = [];
    const seen = new Set();

    for (const item of output) {
        const key = item.key || (
            "NAME|" + clean(item.name).toUpperCase()
        );

        if (seen.has(key)) {
            continue;
        }

        seen.add(key);
        result.push(item);
    }

    return result;
    """

    deadline = time.time() + 25
    found = []
    seen_keys = set()

    while time.time() < deadline:
        try:
            current = (
                driver.execute_script(
                    extract_script,
                    info.get("id", "")
                )
                or []
            )

            for raw in current:
                if not isinstance(raw, dict):
                    continue

                name = " ".join(
                    str(raw.get("name", "") or "").split()
                ).strip()

                if (
                    not name
                    or re.fullmatch(
                        r"Accessory\s+\d+",
                        name,
                        flags=re.IGNORECASE
                    )
                    or re.fullmatch(
                        r"[0-9a-fA-F]{8}-[0-9a-fA-F-]{4,}",
                        name
                    )
                ):
                    continue

                accessory_id_value = " ".join(
                    str(raw.get("accessoryid", "") or "").split()
                ).strip()

                element_id = " ".join(
                    str(raw.get("elementId", "") or "").split()
                ).strip()

                image = " ".join(
                    str(raw.get("image", "") or "").split()
                ).strip()

                if accessory_id_value:
                    key = (
                        "ID",
                        accessory_id_value.casefold(),
                        name.casefold()
                    )
                elif image:
                    key = (
                        "IMG",
                        image.casefold(),
                        name.casefold()
                    )
                elif element_id:
                    key = (
                        "EID",
                        element_id.casefold(),
                        name.casefold()
                    )
                else:
                    # Without any stable identifier, retain occurrences
                    # by DOM position/key instead of name-only de-duping.
                    key = (
                        "FALLBACK",
                        len(found),
                        name.casefold()
                    )

                if key not in seen_keys:
                    seen_keys.add(key)
                    found.append({
                        "name": name,
                        "accessoryid": accessory_id_value,
                        "elementId": element_id,
                        "image": image
                    })

            if found:
                # Give late-rendered popup children one extra opportunity.
                time.sleep(1.2)

                final = (
                    driver.execute_script(
                        extract_script,
                        info.get("id", "")
                    )
                    or []
                )

                for raw in final:
                    if not isinstance(raw, dict):
                        continue

                    name = " ".join(
                        str(raw.get("name", "") or "").split()
                    ).strip()

                    if (
                        not name
                        or re.fullmatch(
                            r"Accessory\s+\d+",
                            name,
                            flags=re.IGNORECASE
                        )
                        or re.fullmatch(
                            r"[0-9a-fA-F]{8}-[0-9a-fA-F-]{4,}",
                            name
                        )
                    ):
                        continue

                    accessory_id_value = " ".join(
                        str(raw.get("accessoryid", "") or "").split()
                    ).strip()

                    element_id = " ".join(
                        str(raw.get("elementId", "") or "").split()
                    ).strip()

                    image = " ".join(
                        str(raw.get("image", "") or "").split()
                    ).strip()

                    if accessory_id_value:
                        key = (
                            "ID",
                            accessory_id_value.casefold(),
                            name.casefold()
                        )
                    elif image:
                        key = (
                            "IMG",
                            image.casefold(),
                            name.casefold()
                        )
                    elif element_id:
                        key = (
                            "EID",
                            element_id.casefold(),
                            name.casefold()
                        )
                    else:
                        key = (
                            "FALLBACK",
                            len(found),
                            name.casefold()
                        )

                    if key not in seen_keys:
                        seen_keys.add(key)
                        found.append({
                            "name": name,
                            "accessoryid": accessory_id_value,
                            "elementId": element_id,
                            "image": image
                        })

                return found

        except Exception as e:
            print("  NUMBERED CARD EXTRACTION ERROR:", e)

        time.sleep(0.8)

    return found

def check_accessory_tab_and_count(driver):
    global LAST_MODEL_LEVEL_ALL_ACCESSORY_NAMES
    LAST_MODEL_LEVEL_ALL_ACCESSORY_NAMES = []
    """
    Fetch every accessory card from the page's complete open Shadow-DOM tree.

    COUNT LOGIC IS PRESERVED:
        normal card      -> 1
        numbered card N  -> N + 1

    NAME LOGIC:
        normal card      -> main card name
        numbered card   -> main card name + child popup names immediately
                           after that parent, before moving to the next card

    Duplicate handling:
        - same name + same accessoryid -> one logical accessory
        - same name + different accessoryid -> both are retained
        - if no stable ID exists, image / element ID / occurrence fallback
          is used instead of name-only de-duplication.
    """
    print("Checking Accessories tab...")

    status = (
        "PASS"
        if count_matches(driver, SEL_ACCESSORY_TAB) > 0
        else "FAIL"
    )

    print("Accessory Tab       :", status)

    if status == "FAIL":
        return status, 0, []

    tab = get_element(driver, SEL_ACCESSORY_TAB)

    if tab is None:
        return status, 0, []

    click_label = _try_click_strategies(driver, tab)

    if click_label is None:
        return status, 0, []

    print(
        "ACCESSORIES TAB CLICKED using strategy "
        f"'{click_label}'"
    )

    time.sleep(5)

    # Wait for the global accessory card list to stabilize.
    deadline = time.time() + 60
    cards = []
    previous_signature = None
    stable_polls = 0

    while time.time() < deadline:
        try:
            current = (
                driver.execute_script(accessory_cards_script)
                or []
            )

            if len(current) >= len(cards):
                cards = current

            signature = [
                (
                    str(x.get("accessoryid", "")),
                    str(x.get("trim", "")),
                    str(x.get("itemName", "")),
                    str(x.get("number", ""))
                )
                for x in current
            ]

            if signature == previous_signature:
                stable_polls += 1
            else:
                stable_polls = 0
                previous_signature = signature

            print(
                "  GLOBAL ACCESSORY CARD POLL:",
                len(current)
            )

            if len(current) > 0 and stable_polls >= 6:
                break

        except Exception as e:
            print("  ACCESSORY POLL ERROR:", e)

        time.sleep(1)

    print(
        "FINAL GLOBAL ACCESSORY CARDS FOUND:",
        len(cards)
    )

    if not cards:
        return status, 0, []

    accessory_names = []
    accessory_name_keys = set()
    total_accessories = 0
    model_level_all_names = []
    model_level_all_keys = set()

    def add_name_record(record, source=""):
        """
        Add one logical accessory occurrence using ID-aware identity.
        Returns True if the record was newly added.
        """
        if isinstance(record, dict):
            name = " ".join(
                str(record.get("name", "") or "").split()
            ).strip()

            accessory_id_value = " ".join(
                str(record.get("accessoryid", "") or "").split()
            ).strip()

            element_id = " ".join(
                str(record.get("elementId", "") or "").split()
            ).strip()

            image = " ".join(
                str(record.get("image", "") or "").split()
            ).strip()
        else:
            name = " ".join(
                str(record or "").split()
            ).strip()
            accessory_id_value = ""
            element_id = ""
            image = ""

        if (
            not name
            or re.fullmatch(
                r"Accessory\s+\d+",
                name,
                flags=re.IGNORECASE
            )
            or re.fullmatch(
                r"[0-9a-fA-F]{8}-[0-9a-fA-F-]{4,}",
                name
            )
        ):
            return False

        if accessory_id_value:
            key = (
                "ID",
                accessory_id_value.casefold(),
                name.casefold()
            )
        elif image:
            key = (
                "IMG",
                image.casefold(),
                name.casefold()
            )
        elif element_id:
            key = (
                "EID",
                element_id.casefold(),
                name.casefold()
            )
        else:
            # If there is genuinely no stable identifier, keep the
            # occurrence. Name alone is NOT treated as globally unique.
            key = (
                "FALLBACK",
                len(accessory_names),
                name.casefold()
            )

        if key in accessory_name_keys:
            print(
                "  DUPLICATE ACCESSORY REMOVED:",
                name,
                f"| ID={accessory_id_value or '<none>'}"
            )
            return False

        accessory_name_keys.add(key)
        accessory_names.append(name)

        print(
            f"  ACCESSORY ADDED [{len(accessory_names)}]:",
            name,
            f"| ID={accessory_id_value or '<none>'}"
        )

        return True

    for index, card in enumerate(cards, start=1):
        main_name = " ".join(
            str(card.get("itemName", "") or "").split()
        ).strip()

        number = card.get("number")
        parent_accessory_id = " ".join(
            str(card.get("accessoryid", "") or "").split()
        ).strip()

        names_for_card = []

        card_trim = " ".join(
            str(card.get("trim", "") or "").split()
        ).strip()
        card_is_model_level = (
            not card_trim
            or card_trim.casefold() in {"all", "all trims", "model"}
        )

        print("")
        print(f"ACCESSORY CARD {index}")
        print(
            "  MAIN NAME:",
            main_name or "<NOT FOUND>"
        )
        print(
            "  ACCESSORY ID:",
            parent_accessory_id or "<NOT FOUND>"
        )
        print(
            "  NUMBER:",
            number,
            card.get("numberSource", "")
        )

        # ----------------------------------------------------------
        # Parent/main accessory is ALWAYS processed first.
        # ----------------------------------------------------------
        if (
            main_name
            and not re.fullmatch(
                r"Accessory\s+\d+",
                main_name,
                flags=re.IGNORECASE
            )
            and not re.fullmatch(
                r"[0-9a-fA-F]{8}-[0-9a-f-]{4,}",
                main_name
            )
        ):
            names_for_card.append({
                "name": main_name,
                "accessoryid": parent_accessory_id,
                "elementId": "",
                "image": ""
            })
            if card_is_model_level and main_name.casefold() not in model_level_all_keys:
                model_level_all_keys.add(main_name.casefold())
                model_level_all_names.append(main_name)

        # ----------------------------------------------------------
        # Numbered card:
        #
        # A [number N]
        #   -> click A by accessoryid
        #   -> fetch A1 ... AN
        #   -> append children immediately
        #   -> move to next main card
        # ----------------------------------------------------------
        if number is not None:
            print(
                "  NUMBERED CARD -> CLICKING BY ACCESSORY ID..."
            )

            inner_records = (
                _click_numbered_accessory_card_and_get_names(
                    driver,
                    index - 1,
                    parent_accessory_id,
                    main_name
                )
            )

            print(
                "  INNER ACCESSORY RECORDS:",
                inner_records
            )

            for record in inner_records:
                if not isinstance(record, dict):
                    continue

                child_name = " ".join(
                    str(record.get("name", "") or "").split()
                ).strip()

                if not child_name:
                    continue

                # Prevent the parent itself from being re-added if the
                # popup inspection also sees the parent .item-name.
                child_id = " ".join(
                    str(record.get("accessoryid", "") or "").split()
                ).strip()

                if (
                    child_id
                    and parent_accessory_id
                    and child_id.casefold()
                    == parent_accessory_id.casefold()
                    and child_name.casefold()
                    == main_name.casefold()
                ):
                    continue

                names_for_card.append(record)
                if card_is_model_level and child_name.casefold() not in model_level_all_keys:
                    model_level_all_keys.add(child_name.casefold())
                    model_level_all_names.append(child_name)

        # ----------------------------------------------------------
        # Add names in EXACT processing order:
        # parent -> its children -> next parent.
        # ----------------------------------------------------------
        for record in names_for_card:
            add_name_record(record)

        # ----------------------------------------------------------
        # PRESERVE THE ORIGINAL COUNT RULE.
        #
        # Normal card  = 1
        # Numbered N   = N + 1
        #
        # Do NOT replace this with len(extracted_names).
        # ----------------------------------------------------------
        if number is not None:
            try:
                contribution = int(number) + 1
            except Exception:
                contribution = 1
        else:
            contribution = 1

        total_accessories += contribution

        print(
            "  FINAL CARD NAMES:",
            [
                r.get("name", "")
                for r in names_for_card
                if isinstance(r, dict)
            ]
        )
        print(
            "  COUNT CONTRIBUTION:",
            contribution
        )

    print("")
    print(
        "ALL ACCESSORY NAMES:",
        accessory_names
    )
    print(
        "TOTAL ACCESSORY NAME ROWS:",
        len(accessory_names)
    )
    print(
        "NUMBER OF ACCESSORIES:",
        total_accessories
    )

    LAST_MODEL_LEVEL_ALL_ACCESSORY_NAMES = list(model_level_all_names)
    print("MODEL-LEVEL ALL-TRIM ACCESSORIES:", LAST_MODEL_LEVEL_ALL_ACCESSORY_NAMES)

    return (
        status,
        total_accessories,
        accessory_names
    )

# =========================================================
# EXCEL OUTPUT
# =========================================================

EXCEL_COLUMNS = [
    "Sl.no", "Year", "Make", "Model", "Base Trim", "Number of Trims",
    "Banner", "Default Color", "IzmoEmporio Text", "Int Hotspot Button",
    "Ext Hotspot Button", "BG Change Button", "Number of BGs",
    "Accessory Tab", "Number of Accessories",
    "Footer Buttons", "Make Logo", "Drag Text Logo",
]

MAIN_COLUMNS = [
    "Year", "Make", "Model", "Other Trims",
    "Number of Accessories", "Checked date with time"
]
ACCESSORY_COLUMNS = ["Sl.No", "Year", "Make", "Model", "Accessories", "Trims"]
COLOR_COLUMNS = ["Year", "Make", "Model", "Other Trims", "Colors Count", "Color Names"]


def _row_values(r):
    return [
        r["number"], r["year"], r["make"], r["model"], r["base_trim"],
        r["number_of_trims"], r["banner"], r["default_color"], r["izmo_emporio"],
        r["int_hotspot"], r["ext_hotspot"], r["bg_change"],
        r["number_of_backgrounds"], r["accessory_tab"],
        r["number_of_accessories"], r["footer_buttons"], r["make_logo"],
        r["drag_text_logo"],
    ]


def _detail_row_values(r):
    return [
        r["number"], r["year"], r["make"], r["model"],
        ", ".join(r.get("trim_names", [])),
        ", ".join(r.get("color_names", [])),
        ", ".join(r.get("accessory_names", [])),
    ]


def _main_rows(results):
    rows = []
    for r in results:
        trims = r.get("trim_names", []) or []
        checked_at = r.get("checked_date_time", "")

        if trims:
            for trim in trims:
                rows.append([
                    r.get("year", ""),
                    r.get("make", ""),
                    r.get("model", ""),
                    trim,
                    r.get("number_of_accessories", 0),
                    checked_at,
                ])
        else:
            rows.append([
                r.get("year", ""),
                r.get("make", ""),
                r.get("model", ""),
                "",
                r.get("number_of_accessories", 0),
                checked_at,
            ])
    return rows

def _accessory_rows(results):
    """
    Build the Accessories sheet.

    REQUIREMENT:
      The number of rows for a model MUST match that model's
      "Number of Accessories" value.

    IMPORTANT:
      - One row represents one accessory occurrence.
      - Duplicate accessory names are NOT removed.
      - Therefore, if the count is 4 and the site contains:
            License Plate Frame
            License Plate Frame
            Molded Splash Guards, Front
            Molded Splash Guards, Rear
        all 4 rows are written.
      - Trim availability is calculated per occurrence where possible.
    """
    rows = []
    serial = 1

    for r in results:

        trim_names = [
            " ".join(
                str(t or "").split()
            ).strip()
            for t in (
                r.get("trim_names", [])
                or []
            )
            if " ".join(
                str(t or "").split()
            ).strip()
        ]

        ordered_trims = []

        for trim in trim_names:
            if trim.casefold() not in {
                x.casefold()
                for x in ordered_trims
            }:
                ordered_trims.append(trim)

        accessories_by_trim = (
            r.get(
                "accessories_by_trim",
                {}
            )
            or {}
        )

        if not ordered_trims:
            ordered_trims = [
                str(k).strip()
                for k in accessories_by_trim.keys()
                if str(k).strip()
            ]

        # ----------------------------------------------------------
        # Expected model-level count from Validation Results.
        # ----------------------------------------------------------
        try:
            expected_count = int(
                r.get(
                    "number_of_accessories",
                    0
                )
                or 0
            )
        except Exception:
            expected_count = 0

        # ----------------------------------------------------------
        # Preserve EVERY accessory occurrence.
        #
        # Priority:
        #   1) initial/default accessory_names
        #   2) first/default trim
        #   3) other trims, only if needed to reach expected count
        # ----------------------------------------------------------
        candidates = []

        def valid_name(raw):
            name = " ".join(
                str(raw or "").split()
            ).strip()

            if not name:
                return None

            if re.fullmatch(
                r"Accessory\s+\d+",
                name,
                flags=re.IGNORECASE
            ):
                return None

            if re.fullmatch(
                r"[0-9a-fA-F]{8}-[0-9a-fA-F-]{4,}",
                name
            ):
                return None

            return name

        def append_occurrence(raw):
            name = valid_name(raw)
            if name:
                candidates.append(name)

        # Default/current trim result.
        for name in (
            r.get("accessory_names", [])
            or []
        ):
            append_occurrence(name)

        # Default/first trim result.
        if ordered_trims:
            first_trim = ordered_trims[0]

            for name in (
                accessories_by_trim.get(
                    first_trim,
                    []
                )
                or []
            ):
                append_occurrence(name)

        # Supplement with every trim if the initial/default scan is short.
        if (
            expected_count > 0 and
            len(candidates) < expected_count
        ):
            for trim in ordered_trims:
                for name in (
                    accessories_by_trim.get(
                        trim,
                        []
                    )
                    or []
                ):
                    append_occurrence(name)

                    if (
                        len(candidates)
                        >= expected_count
                    ):
                        break

                if (
                    len(candidates)
                    >= expected_count
                ):
                    break

        # ----------------------------------------------------------
        # DO NOT deduplicate.
        #
        # Exactly expected_count occurrences are emitted when the
        # model-level count is known.
        # ----------------------------------------------------------
        if expected_count > 0:
            selected = candidates[
                :expected_count
            ]
        else:
            selected = candidates

        print("")
        print(
            "ACCESSORIES SHEET:",
            r.get("year", ""),
            r.get("make", ""),
            r.get("model", "")
        )
        print(
            "  Expected count:",
            expected_count
        )
        print(
            "  Collected occurrences:",
            len(candidates)
        )
        print(
            "  Rows to write:",
            len(selected)
        )

        if (
            expected_count > 0 and
            len(selected) < expected_count
        ):
            print(
                "  WARNING: Cannot meet expected count. "
                f"Expected {expected_count}, "
                f"but only {len(selected)} valid accessory "
                "occurrences were collected."
            )

        # ----------------------------------------------------------
        # Determine trim coverage for each occurrence.
        #
        # Duplicate names are handled by occurrence order within each
        # trim. This means:
        #
        #   Trim A:
        #       License Plate Frame
        #       License Plate Frame
        #
        # is treated as occurrence 1 and occurrence 2 rather than
        # collapsing both into one name.
        # ----------------------------------------------------------
        occurrence_maps = {}

        model_all_names = set(
            x.casefold()
            for x in (r.get("model_level_all_accessory_names", []) or [])
            if str(x).strip()
        )

        for trim in ordered_trims:

            trim_occurrences = []

            for raw_name in (
                accessories_by_trim.get(
                    trim,
                    []
                )
                or []
            ):
                name = valid_name(
                    raw_name
                )

                if name:
                    trim_occurrences.append(
                        name
                    )

            occurrence_maps[
                trim
            ] = trim_occurrences

        # Count how many times each name has been used in the output.
        output_occurrence_counter = {}

        for display_name in selected:

            name_key = display_name.casefold()

            occurrence_index = (
                output_occurrence_counter.get(
                    name_key,
                    0
                )
            )

            output_occurrence_counter[
                name_key
            ] = occurrence_index + 1

            found_trims = []

            for trim in ordered_trims:

                trim_occurrences = (
                    occurrence_maps.get(
                        trim,
                        []
                    )
                )

                matching_index = 0

                for occurrence_name in (
                    trim_occurrences
                ):
                    if (
                        occurrence_name.casefold()
                        == name_key
                    ):
                        if (
                            matching_index
                            == occurrence_index
                        ):
                            found_trims.append(
                                trim
                            )
                            break

                        matching_index += 1

            # ------------------------------------------------------
            # Fallback:
            # if occurrence-level matching cannot be established,
            # still report all trims where the name exists.
            # ------------------------------------------------------
            if (
                not found_trims and
                ordered_trims
            ):
                for trim in ordered_trims:
                    exists = any(
                        name.casefold()
                        == name_key
                        for name in occurrence_maps.get(
                            trim,
                            []
                        )
                    )

                    if exists:
                        found_trims.append(
                            trim
                        )

            unique_found_trims = []

            for trim in ordered_trims:
                if (
                    trim in found_trims and
                    trim not in unique_found_trims
                ):
                    unique_found_trims.append(
                        trim
                    )

            # Model-level accessory: its source card had blank/ALL trim,
            # therefore it is available to every trim.
            if name_key in model_all_names and ordered_trims:
                trim_output = "ALL"
            elif (
                ordered_trims and
                len(unique_found_trims)
                == len(ordered_trims)
            ):
                trim_output = "ALL"
            else:
                trim_output = ", ".join(
                    unique_found_trims
                )

            rows.append([
                serial,
                r.get("year", ""),
                r.get("make", ""),
                r.get("model", ""),
                display_name,
                trim_output
            ])

            serial += 1

    return rows

def _color_rows(results):
    """
    One row per trim.
    All colors for that exact trim are comma-separated.
    """
    rows = []

    for r in results:

        trim_names = [
            " ".join(
                str(t or "").split()
            ).strip()
            for t in (
                r.get("trim_names", [])
                or []
            )
        ]

        trim_names = [
            t for t in trim_names if t
        ]

        colors_by_trim = (
            r.get("colors_by_trim", {})
            or {}
        )

        if trim_names:

            for trim in trim_names:

                raw_colors = (
                    colors_by_trim.get(
                        trim,
                        []
                    )
                    or []
                )

                colors = []
                seen = set()

                for raw_color in raw_colors:

                    color = " ".join(
                        str(raw_color or "").split()
                    ).strip()

                    key = color.casefold()

                    if color and key not in seen:
                        seen.add(key)
                        colors.append(color)

                rows.append([
                    r.get("year", ""),
                    r.get("make", ""),
                    r.get("model", ""),
                    trim,
                    len(colors),
                    ", ".join(colors)
                ])

        else:

            raw_colors = (
                r.get("color_names", [])
                or []
            )

            colors = []
            seen = set()

            for raw_color in raw_colors:

                color = " ".join(
                    str(raw_color or "").split()
                ).strip()

                key = color.casefold()

                if color and key not in seen:
                    seen.add(key)
                    colors.append(color)

            rows.append([
                r.get("year", ""),
                r.get("make", ""),
                r.get("model", ""),
                "",
                len(colors),
                ", ".join(colors)
            ])

    return rows

def _write_dynamic_sheet(ws, columns, rows):
    header_font=Font(name="Arial",bold=True,color="FFFFFF")
    header_fill=PatternFill(start_color="1F4E78",end_color="1F4E78",fill_type="solid")
    body_font=Font(name="Arial"); center=Alignment(horizontal="center",vertical="center"); wrap=Alignment(horizontal="left",vertical="top",wrap_text=True)
    for i,h in enumerate(columns,1):
        c=ws.cell(1,i,h); c.font=header_font; c.fill=header_fill; c.alignment=center
    for ri,row in enumerate(rows,2):
        for ci,v in enumerate(row,1):
            c=ws.cell(ri,ci,v); c.font=body_font; c.alignment=wrap if isinstance(v,str) and len(v)>40 else center
    for ci,h in enumerate(columns,1):
        m=len(str(h))
        for row in rows:
            if ci-1 < len(row): m=max(m,len(str(row[ci-1])[:60]))
        ws.column_dimensions[get_column_letter(ci)].width=min(m+4,60)

def _write_sheet(ws, columns, results, row_value_fn):
    header_font = Font(name="Arial", bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    body_font = Font(name="Arial")
    center = Alignment(horizontal="center", vertical="center")
    wrap_left = Alignment(horizontal="left", vertical="top", wrap_text=True)

    for col_idx, header in enumerate(columns, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center

    for row_idx, r in enumerate(results, start=2):
        row_values = row_value_fn(r)
        for col_idx, value in enumerate(row_values, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.font = body_font
            # Long comma-separated lists read better left-aligned + wrapped
            cell.alignment = wrap_left if isinstance(value, str) and len(value) > 40 else center

    for col_idx, header in enumerate(columns, start=1):
        max_len = max(
            [len(str(header))] + [len(str(row_value_fn(r)[col_idx - 1])[:60]) for r in results]
            if results else [len(str(header))]
        )
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 4, 60)


def _highlight_base_trim_rows(ws, results):
    """Highlight the Main-sheet row corresponding to each model's base trim."""
    base_keys = set()

    for r in results:
        year = " ".join(str(r.get("year", "") or "").split()).casefold()
        make = " ".join(str(r.get("make", "") or "").split()).casefold()
        model = " ".join(str(r.get("model", "") or "").split()).casefold()
        base_trim = " ".join(str(r.get("base_trim", "") or "").split()).casefold()

        if year and make and model and base_trim:
            base_keys.add((year, make, model, base_trim))

    # Highlight the complete Main row.
    base_fill = PatternFill(
        start_color="FFF2CC",
        end_color="FFF2CC",
        fill_type="solid"
    )
    base_font = Font(name="Arial", bold=True)

    for row_idx in range(2, ws.max_row + 1):
        key = (
            " ".join(str(ws.cell(row_idx, 1).value or "").split()).casefold(),
            " ".join(str(ws.cell(row_idx, 2).value or "").split()).casefold(),
            " ".join(str(ws.cell(row_idx, 3).value or "").split()).casefold(),
            " ".join(str(ws.cell(row_idx, 4).value or "").split()).casefold(),
        )

        if key in base_keys:
            for col_idx in range(1, ws.max_column + 1):
                cell = ws.cell(row_idx, col_idx)
                cell.fill = base_fill
                cell.font = base_font


def write_results_to_excel(results, path=OUTPUT_XLSX_PATH):
    wb = Workbook()

    # =====================================================
    # Sheet 1: Validation Results
    # =====================================================
    ws1 = wb.active
    ws1.title = "Validation Results"
    _write_sheet(
        ws1,
        EXCEL_COLUMNS,
        results,
        _row_values
    )

    # =====================================================
    # Sheet 2: Main
    # =====================================================
    ws2 = wb.create_sheet("Main")
    main_rows = _main_rows(results)
    _write_dynamic_sheet(
        ws2,
        MAIN_COLUMNS,
        main_rows
    )
    _highlight_base_trim_rows(ws2, results)

    # =====================================================
    # Sheet 3: Accessories
    # =====================================================
    ws3 = wb.create_sheet("Accessories")

    accessory_rows = _accessory_rows(
        results
    )

    _write_dynamic_sheet(
        ws3,
        ACCESSORY_COLUMNS,
        accessory_rows
    )

    # =====================================================
    # Sheet 4: Colors
    # =====================================================
    ws4 = wb.create_sheet("Colors")
    color_rows = _color_rows(
        results
    )

    _write_dynamic_sheet(
        ws4,
        COLOR_COLUMNS,
        color_rows
    )

    wb.save(path)

    print("")
    print(
        f"Excel results written to: {path}"
    )
    print(
        "Accessories sheet total rows:",
        len(accessory_rows)
    )
    print(
        "Colors sheet rows:",
        len(color_rows)
    )

    # =====================================================
    # MODEL-LEVEL ACCESSORY COUNT VALIDATION
    # =====================================================
    print("")
    print("=" * 70)
    print("ACCESSORY COUNT VS ACCESSORIES SHEET CHECK")
    print("=" * 70)

    all_pass = True

    for result in results:

        year = result.get(
            "year",
            ""
        )
        make = result.get(
            "make",
            ""
        )
        model = result.get(
            "model",
            ""
        )

        try:
            expected = int(
                result.get(
                    "number_of_accessories",
                    0
                )
                or 0
            )
        except Exception:
            expected = 0

        actual = sum(
            1
            for row in accessory_rows
            if (
                row[1] == year and
                row[2] == make and
                row[3] == model
            )
        )

        passed = (
            expected == actual
        )

        if not passed:
            all_pass = False

        print(
            f"{year} {make} {model}: "
            f"Expected={expected}, "
            f"Accessories Sheet Rows={actual} -> "
            f"{'PASS' if passed else 'FAIL'}"
        )

    print(
        "OVERALL ACCESSORY ROW CHECK:",
        "PASS" if all_pass else "FAIL"
    )


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
                "checked_date_time": datetime.now(ZoneInfo("Asia/Kolkata")).strftime("%Y-%m-%d %H:%M:%S"),
                "base_trim": "", "number_of_trims": 0, "banner": "FAIL",
                "default_color": "",
                "izmo_emporio": "FAIL", "int_hotspot": "FAIL",
                "ext_hotspot": "FAIL", "bg_change": "FAIL", "number_of_backgrounds": 0,
                "accessory_tab": "FAIL", "number_of_accessories": 0,
                "footer_buttons": "FAIL", "make_logo": "FAIL",
                "drag_text_logo": "FAIL",
                "trim_names": [], "color_names": [], "accessory_names": [],
                "colors_by_trim": {}, "accessories_by_trim": {}, "accessory_counts_by_trim": {},
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
                "checked_date_time": datetime.now(ZoneInfo("Asia/Kolkata")).strftime("%Y-%m-%d %H:%M:%S"),
                "base_trim": "", "number_of_trims": 0, "banner": banner_status,
                "default_color": "",
                "izmo_emporio": "FAIL", "int_hotspot": "FAIL",
                "ext_hotspot": "FAIL", "bg_change": "FAIL", "number_of_backgrounds": 0,
                "accessory_tab": "FAIL", "number_of_accessories": 0,
                "footer_buttons": "FAIL", "make_logo": "FAIL",
                "drag_text_logo": "FAIL",
                "trim_names": [], "color_names": [], "accessory_names": [],
                "colors_by_trim": {}, "accessories_by_trim": {}, "accessory_counts_by_trim": {},
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

        # Keep the existing Validation Results behavior:
        # default/base color and base-trim accessory count are collected
        # from the vehicle's initial selected trim.
        print("Fetching all color names for base trim...")
        color_names = get_all_color_names(driver)
        print(
            f"BASE-TRIM COLORS ({len(color_names)}):",
            color_names
        )

        # Close the base-trim More Colors popup before opening the trim selector.
        close_color_popup_if_open(driver)

        # Get all trim names.
        number_of_trims, trim_names = check_number_of_trims(driver)

        # Close the trim popup before the existing validation checks.
        close_trim_popup_if_open(driver)

        number_of_backgrounds = check_number_of_backgrounds(driver)

        # Existing Validation Results accessory count is preserved.
        accessory_tab_status, number_of_accessories, accessory_names = (
            check_accessory_tab_and_count(driver)
        )

        # NEW:
        # Click every trim and collect exact accessories + colors.
        (
            colors_by_trim,
            accessories_by_trim,
            accessory_counts_by_trim
        ) = collect_accessories_and_colors_for_each_trim(
            driver,
            trim_names,
            accessory_names,
            LAST_MODEL_LEVEL_ALL_ACCESSORY_NAMES
        )

        checked_date_time = datetime.now(ZoneInfo("Asia/Kolkata")).strftime("%Y-%m-%d %H:%M:%S")

        results.append({
            "number": model_index + 1,
            "year": year,
            "checked_date_time": checked_date_time,
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
            "trim_names": trim_names,
            "color_names": color_names,
            "accessory_names": accessory_names,
            "model_level_all_accessory_names": list(
                LAST_MODEL_LEVEL_ALL_ACCESSORY_NAMES
            ),

            # Per-trim data used by Accessories and Colors sheets.
            "colors_by_trim": colors_by_trim,
            "accessories_by_trim": accessories_by_trim,
            "accessory_counts_by_trim": accessory_counts_by_trim,
        })

        print("")
        print(f"MODEL {model_index + 1} VALIDATION COMPLETED")

    # Diagnostic: prove that per-trim dictionaries actually contain
    # data before Excel generation.
    print("")
    print("============================================================")
    print("PER-TRIM DATA SUMMARY BEFORE EXCEL")
    print("============================================================")

    for result in results:
        print(
            f"{result.get('year')} "
            f"{result.get('make')} "
            f"{result.get('model')}"
        )
        print(
            "  COLORS BY TRIM:",
            result.get("colors_by_trim", {})
        )
        print(
            "  ACCESSORIES BY TRIM:",
            result.get("accessories_by_trim", {})
        )

    write_results_to_excel(results)

finally:
    driver.quit()

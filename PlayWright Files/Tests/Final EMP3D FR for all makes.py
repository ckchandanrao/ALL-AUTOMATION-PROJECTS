# Generated from the user's uploaded Emporio validation script.
"""
izmo Emporio validation script - UNIVERSAL (works for any manufacturer's
izmo-emporio site, e.g. Chrysler, Volvo, etc.)

Usage:
    python emporio_validator.py [URL]

If no URL is given, defaults to the Chrysler Pacifica/Voyager site.

Writes results to an Excel file with FOUR sheets:

  Sheet 1 "Validation Results" (unchanged from before):
    Sl.no, Make, Model, Base Trim, Number of Trims, Banner,
    Default Color, IzmoEmporio Text, Int Hotspot Button,
    Ext Hotspot Button, BG Change Button, Number of BGs,
    Accessory Tab, Number of Accessories, Footer Buttons, Make Logo,
    Drag Text Logo

  Sheet 2 "Main" (NEW):
    Sl.no, Make, Model, Other Trims, Colors, Accessories
    (each of the last three is a comma-separated list of every name
    found for that model)

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

DEFAULT_URL = "https://dacia3d.izmocars.fr/emporio-fr-fr.htm"
if len(sys.argv) > 1:
    URL = sys.argv[1].strip()
else:
    URL = input("Enter Emporio URL: ").strip()
if not URL:
    raise SystemExit("No Emporio URL provided.")

INITIAL_LOAD_WAIT = 15
MODEL_LOAD_TIMEOUT = 60
BANNER_POLL_TIMEOUT = 15
FULL_CAR_TIMEOUT = 60
CAR_RENDER_WAIT = 3
DEFAULT_COLOR_TIMEOUT = 45
POPUP_OPEN_TIMEOUT = 15

REQUIRED_FOOTER_BUTTONS = 2

EXPECTED_BANNER_WIDTH = 1920
EXPECTED_BANNER_HEIGHT = 274

_site_slug = re.sub(r"[^a-zA-Z0-9]+", "_", urlparse(URL).netloc).strip("_") or "emporio"

RUN_DATE_TIME = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
RUN_FILE_TIMESTAMP = datetime.now().strftime("%Y%m%d_%H%M%S")
OUTPUT_XLSX_PATH = (
    f"emporio_validation_results_{_site_slug}_{RUN_FILE_TIMESTAMP}.xlsx"
)

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
verify_any_popup_opened_script = r"""
function visible(el) {
    if (!el) return false;
    try {
        if (el.classList && el.classList.contains("hide")) return false;
        const st = getComputedStyle(el);
        const r = el.getBoundingClientRect();
        return st.display !== "none" && st.visibility !== "hidden" &&
               st.opacity !== "0" && r.width > 0 && r.height > 0;
    } catch (e) { return false; }
}
function check(root) {
    if (!root) return false;
    try {
        const popups = root.querySelectorAll(
            ".more-items-popUp, .trim-popup-panel, .background-panel, .pannel-popup"
        );
        for (const p of popups) {
            if (visible(p)) return true;
            try {
                const child = p.querySelector(
                    ".izmo-emporio-trim.item, .izmo-emporio-accessory.item.action"
                );
                if (visible(child)) return true;
            } catch (e) {}
        }
    } catch (e) {}
    try {
        for (const el of (root.querySelectorAll ? root.querySelectorAll("*") : [])) {
            if (el.shadowRoot && check(el.shadowRoot)) return true;
        }
    } catch (e) {}
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

function visible(el) {
    if (!el) return false;

    try {
        const s = getComputedStyle(el);
        const r = el.getBoundingClientRect();

        return (
            !el.classList.contains("hide") &&
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

function walkOpenShadow(root, callback) {
    if (!root) return;

    try {
        if (root.nodeType === Node.ELEMENT_NODE) {
            callback(root);
        }
    } catch (e) {}

    try {
        const all = root.querySelectorAll
            ? Array.from(root.querySelectorAll("*"))
            : [];

        for (const el of all) {
            callback(el);

            try {
                if (el.shadowRoot) {
                    walkOpenShadow(
                        el.shadowRoot,
                        callback
                    );
                }
            } catch (e) {}
        }
    } catch (e) {}

    try {
        if (root.shadowRoot) {
            walkOpenShadow(
                root.shadowRoot,
                callback
            );
        }
    } catch (e) {}
}

function findSections(root, result) {
    if (!root) return;

    try {
        const sections = root.querySelectorAll
            ? Array.from(
                root.querySelectorAll(
                    ".izmo-emporio-accessory-section"
                )
            )
            : [];

        for (const section of sections) {
            if (
                visible(section) &&
                !result.includes(section)
            ) {
                result.push(section);
            }
        }
    } catch (e) {}

    try {
        const all = root.querySelectorAll
            ? Array.from(root.querySelectorAll("*"))
            : [];

        for (const el of all) {
            if (el.shadowRoot) {
                findSections(
                    el.shadowRoot,
                    result
                );
            }
        }
    } catch (e) {}
}

function mainName(card) {
    let nameEl = null;

    try {
        nameEl = card.querySelector(
            ".item-name.fontFamilyRegular.fontSize12"
        );
    } catch (e) {}

    if (!nameEl) {
        try {
            nameEl = card.querySelector(".item-name");
        } catch (e) {}
    }

    if (!nameEl) {
        /*
         * Name may exist inside the card's own open Shadow DOM.
         */
        try {
            let found = null;

            walkOpenShadow(
                card,
                function(el) {
                    if (found) return;

                    try {
                        if (
                            el.matches &&
                            el.matches(
                                ".item-name.fontFamilyRegular.fontSize12"
                            )
                        ) {
                            found = el;
                        }
                    } catch (e) {}
                }
            );

            nameEl = found;
        } catch (e) {}
    }

    if (!nameEl || !visible(nameEl)) {
        return "";
    }

    return clean(
        nameEl.innerText ||
        nameEl.textContent ||
        ""
    );
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
     * Explicit count/number attributes.
     */
    walkOpenShadow(
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
     * CSS generated badge number.
     */
    walkOpenShadow(
        card,
        function(el) {
            if (!visible(el)) return;

            for (const pseudo of [
                "::before",
                "::after"
            ]) {
                try {
                    let c =
                        getComputedStyle(
                            el,
                            pseudo
                        ).content;

                    c = String(
                        c || ""
                    )
                        .replace(
                            /^['"]|['"]$/g,
                            ""
                        )
                        .trim();

                    add(
                        c,
                        "CSS " + pseudo,
                        90
                    );
                } catch (e) {}
            }
        }
    );

    /*
     * Numeric standalone elements.
     */
    walkOpenShadow(
        card,
        function(el) {
            if (!visible(el)) return;

            try {
                const text = clean(
                    el.innerText ||
                    el.textContent ||
                    ""
                );

                if (/^\d+$/.test(text)) {
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
                            : 50
                    );
                }
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

function getCardsFromSection(section) {
    const cards = [];

    function addCard(card) {
        if (
            card &&
            visible(card) &&
            !cards.includes(card)
        ) {
            cards.push(card);
        }
    }

    try {
        const direct =
            section.querySelectorAll
                ? Array.from(
                    section.querySelectorAll(
                        "izmo-emporio-accessory.item.action[accessoryid], " +
                        "izmo-emporio-div.izmo-emporio-accessory.item.action[accessoryid]"
                    )
                )
                : [];

        for (const card of direct) {
            addCard(card);
        }
    } catch (e) {}

    /*
     * Deeply inspect open Shadow DOM inside this section.
     */
    walkOpenShadow(
        section,
        function(el) {
            try {
                if (
                    el.matches &&
                    el.matches(
                        "izmo-emporio-accessory.item.action[accessoryid], " +
                        "izmo-emporio-div.izmo-emporio-accessory.item.action[accessoryid]"
                    )
                ) {
                    addCard(el);
                }
            } catch (e) {}
        }
    );

    return cards;
}

function getScrollers(section) {
    const scrollers = [];

    function consider(el) {
        if (
            !el ||
            !visible(el)
        ) {
            return;
        }

        try {
            if (
                (el.scrollHeight > el.clientHeight + 5) ||
                (el.scrollWidth > el.clientWidth + 5)
            ) {
                if (!scrollers.includes(el)) {
                    scrollers.push(el);
                }
            }
        } catch (e) {}
    }

    consider(section);

    try {
        const all =
            section.querySelectorAll
                ? Array.from(
                    section.querySelectorAll("*")
                )
                : [];

        for (const el of all) {
            consider(el);

            try {
                if (el.shadowRoot) {
                    walkOpenShadow(
                        el.shadowRoot,
                        consider
                    );
                }
            } catch (e) {}
        }
    } catch (e) {}

    /*
     * Most likely card-list scrollers first:
     * large visible scrollable elements.
     */
    scrollers.sort(
        (a, b) => {
            try {
                const ar =
                    a.getBoundingClientRect();
                const br =
                    b.getBoundingClientRect();

                return (
                    (br.width * br.height) -
                    (ar.width * ar.height)
                );
            } catch (e) {
                return 0;
            }
        }
    );

    return scrollers;
}

const sections = [];
findSections(
    document,
    sections
);

const result = [];

for (const section of sections) {

    /*
     * Start at the top.
     */
    try {
        section.scrollTop = 0;
    } catch (e) {}

    const scrollers =
        getScrollers(section);

    for (const scroller of scrollers) {
        try {
            scroller.scrollTop = 0;
        } catch (e) {}
    }

    /*
     * IMPORTANT:
     * Scroll through the accessory section because Emporio can render
     * only the first several cards initially.
     */
    for (let pass = 0; pass < 300; pass++) {

        const cards =
            getCardsFromSection(section);

        for (
            const card of cards
        ) {
            const number =
                getNumber(card);

            result.push({
                sectionIndex:
                    sections.indexOf(section),
                accessoryid:
                    card.getAttribute(
                        "accessoryid"
                    ) || "",
                trim:
                    card.getAttribute(
                        "trim"
                    ) || "",
                itemName:
                    mainName(card),
                number:
                    number.value,
                numberSource:
                    number.source
            });
        }

        let moved = false;

        /*
         * Scroll the section itself.
         */
        try {
            const old =
                section.scrollTop;

            const max =
                Math.max(
                    0,
                    section.scrollHeight -
                    section.clientHeight
                );

            section.scrollTop =
                Math.min(
                    max,
                    old +
                    Math.max(
                        100,
                        Math.floor(
                            section.clientHeight * 0.65
                        )
                    )
                );

            if (
                section.scrollTop !==
                old
            ) {
                moved = true;
            }
        } catch (e) {}

        /*
         * Scroll every internal scrollable list as well.
         */
        for (const scroller of scrollers) {
            try {
                const oldTop = scroller.scrollTop;
                const oldLeft = scroller.scrollLeft;

                const maxTop =
                    Math.max(
                        0,
                        scroller.scrollHeight -
                        scroller.clientHeight
                    );

                const maxLeft =
                    Math.max(
                        0,
                        scroller.scrollWidth -
                        scroller.clientWidth
                    );

                if (maxTop > 0) {
                    scroller.scrollTop =
                        Math.min(
                            maxTop,
                            oldTop +
                            Math.max(
                                100,
                                Math.floor(
                                    scroller.clientHeight * 0.65
                                )
                            )
                        );
                }

                if (maxLeft > 0) {
                    scroller.scrollLeft =
                        Math.min(
                            maxLeft,
                            oldLeft +
                            Math.max(
                                100,
                                Math.floor(
                                    scroller.clientWidth * 0.75
                                )
                            )
                        );
                }

                if (
                    scroller.scrollTop !== oldTop ||
                    scroller.scrollLeft !== oldLeft
                ) {
                    moved = true;
                }
            } catch (e) {}
        }

        if (!moved) {
            break;
        }
    }

    /*
     * Final bottom scan.
     */
    try {
        section.scrollTop =
            section.scrollHeight;
    } catch (e) {}

    for (const scroller of scrollers) {
        try {
            scroller.scrollTop =
                scroller.scrollHeight;
        } catch (e) {}
    }

    const finalCards =
        getCardsFromSection(section);

    for (
        const card of finalCards
    ) {
        const number =
            getNumber(card);

        result.push({
            sectionIndex:
                sections.indexOf(section),
            accessoryid:
                card.getAttribute(
                    "accessoryid"
                ) || "",
            trim:
                card.getAttribute(
                    "trim"
                ) || "",
            itemName:
                mainName(card),
            number:
                number.value,
            numberSource:
                number.source
        });
    }

    /*
     * Return top so later interactions start from a predictable state.
     */
    try {
        section.scrollTop = 0;
    } catch (e) {}

    for (const scroller of scrollers) {
        try {
            scroller.scrollTop = 0;
        } catch (e) {}
    }
}

/*
 * De-duplicate only the exact physical DOM element occurrence.
 * We intentionally DO NOT deduplicate by accessoryid.
 */
const unique = [];
const seen = new Set();

for (const item of result) {
    const key = [
        item.sectionIndex,
        item.accessoryid,
        item.trim,
        item.itemName,
        item.cardIndex
    ].join("|");

    /*
     * Same card can be returned on multiple scroll passes.
     * For the same section/id/trim/name, retain one occurrence.
     * Different cards with the same name remain separate when the
     * accessoryid or DOM occurrence differs.
     */
    const stableKey = [
        item.sectionIndex,
        item.accessoryid,
        item.trim,
        item.itemName
    ].join("|");

    if (!seen.has(stableKey)) {
        seen.add(stableKey);
        unique.push(item);
    }
}

return unique;

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
    """
    Fetch the Default Color from the exact Emporio element:

        .color-section p.fontFamilyBold.fontSize16

    Example HTML:
        <p class="fontFamilyBold fontSize16">BLEU ARCTIC</p>

    This uses an explicit Selenium WebDriverWait so the function waits
    until the default-color text is actually populated after a model/make
    selection. No other color, accessory, trim, or validation logic is
    changed here.
    """

    DEFAULT_COLOR_SELECTOR = ".color-section p.fontFamilyBold.fontSize16"

    print(
        f"  DEFAULT COLOR: explicit wait started "
        f"(timeout={timeout}s)..."
    )

    def get_default_color(_driver):
        try:
            elements = _find_all_in_open_shadow_dom(
                _driver,
                DEFAULT_COLOR_SELECTOR
            )

            for element in elements:
                try:
                    if not element.is_displayed():
                        continue
                except Exception:
                    continue

                try:
                    text = _clean_color_name(
                        element.text or
                        element.get_attribute("textContent") or
                        ""
                    )
                except Exception:
                    text = ""

                if text:
                    return text

        except Exception:
            pass

        return False

    try:
        default_color = WebDriverWait(
            driver,
            timeout,
            poll_frequency=0.5
        ).until(get_default_color)

        default_color = _clean_color_name(default_color)

        if default_color:
            print(
                "  DEFAULT COLOR: FOUND:",
                default_color
            )
            return default_color

    except TimeoutException:
        print(
            "  DEFAULT COLOR: Timeout - exact default-color element "
            "was not populated"
        )
    except Exception as e:
        print(
            "  DEFAULT COLOR: explicit wait error:",
            e
        )

    print("  DEFAULT COLOR: NOT FOUND AFTER EXPLICIT WAIT")
    return None


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
    """Open an Emporio popup and verify that it actually opened."""
    strategies = [
        ("WebElement.click", lambda el: el.click()),
        ("ActionChains.click", lambda el: ActionChains(driver).move_to_element(el).pause(0.1).click().perform()),
        ("JavaScript click", lambda el: driver.execute_script("arguments[0].click();", el)),
        ("dispatch MouseEvent", lambda el: driver.execute_script("""
            arguments[0].dispatchEvent(new MouseEvent("click", {
                bubbles:true, cancelable:true, view:window
            }));
        """, el)),
    ]

    for label, selector in candidates:
        try:
            element = get_element(driver, selector)
            if element is None:
                print(f"  [{label}]: element not found")
                continue

            try:
                driver.execute_script(
                    "arguments[0].scrollIntoView({block:'center',inline:'center'});",
                    element
                )
            except Exception:
                pass

            for strategy_name, action in strategies:
                try:
                    action(element)
                except Exception:
                    continue

                end_time = time.time() + 4
                while time.time() < end_time:
                    try:
                        if driver.execute_script(verify_any_popup_opened_script):
                            print(f"  OPENED via [{label}] using strategy '{strategy_name}'")
                            return True
                    except Exception:
                        pass
                    time.sleep(0.2)

                print(f"  [{label}] strategy '{strategy_name}' did not open a popup")
        except Exception as exc:
            print(f"  [{label}] error: {exc}")

    # Important fallback for builds where the trim selector CSS path changes.
    try:
        element = driver.execute_script(find_by_exact_text_script, "SELECT TRIM")
        if element is not None:
            for strategy_name, action in strategies:
                try:
                    action(element)
                except Exception:
                    continue
                end_time = time.time() + 4
                while time.time() < end_time:
                    try:
                        if driver.execute_script(verify_any_popup_opened_script):
                            print(f"  OPENED via [exact text 'SELECT TRIM'] using strategy '{strategy_name}'")
                            return True
                    except Exception:
                        pass
                    time.sleep(0.2)
    except Exception as exc:
        print("  [exact text 'SELECT TRIM'] fallback error:", exc)

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


def check_number_of_trims(driver, fallback_base_trim=""):
    """
    Open the trim selector, collect every visible trim item, and return:
        (number_of_trims, trim_names)

    WORKING TRIM COUNT LOGIC:
        - Open the Emporio trim popup.
        - Read .more-items-popUp
        - Read every .izmo-emporio-trim.item inside the visible popup.
        - Prefer the visible text; use the `trim` attribute as fallback.
        - Remove duplicate names case-insensitively while preserving order.
        - Number of trims = number of unique usable trim names.
    """
    print("Opening trims popup...")

    opened = open_popup_robust(
        driver,
        TRIM_BUTTON_CANDIDATES
    )

    if not opened:
        print("TRIMS POPUP: failed to open via any strategy")

        fallback = " ".join(
            str(fallback_base_trim or "").split()
        ).strip()

        if fallback:
            print(
                "TRIMS FALLBACK: using base trim:",
                fallback
            )
            return 1, [fallback]

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

    # The unique usable list is more reliable than the raw DOM count
    # if duplicate/stale popup nodes exist.
    usable_count = len(trim_names)

    if usable_count == 0:
        fallback = " ".join(
            str(fallback_base_trim or "").split()
        ).strip()

        if fallback:
            trim_names = [fallback]
            usable_count = 1
            print(
                "TRIMS FALLBACK: no usable popup trim names ->",
                fallback
            )

    print(
        f"NUMBER OF TRIMS ({usable_count}):",
        trim_names
    )

    return usable_count, trim_names

def select_trim_by_name(driver, trim_name, timeout=20):
    """
    Select one exact trim from the visible Emporio trim popup.

    Search is recursive through normal DOM + every open Shadow DOM.
    Matching is done first against the `trim` attribute and then against
    the visible text. The function returns True once the click is fired.
    """
    expected = " ".join(
        str(trim_name or "").split()
    ).strip()

    if not expected:
        return False

    # Always start from a closed color popup. A modal can intercept the
    # trim selector otherwise.
    try:
        close_color_popup_if_open(driver)
    except Exception:
        pass

    # Open the trim popup.
    if not open_popup_robust(
        driver,
        TRIM_BUTTON_CANDIDATES
    ):
        print(
            f"TRIM SELECT: could not open trim popup for '{expected}'"
        )
        return False

    script = r"""
    function clean(v) {
        return String(v || "").replace(/\s+/g, " ").trim();
    }

    function visible(el) {
        if (!el) return false;

        try {
            if (el.classList && el.classList.contains("hide")) {
                return false;
            }

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

    function same(a, b) {
        return clean(a).toUpperCase() === clean(b).toUpperCase();
    }

    function search(root, expected) {
        if (!root) return null;

        try {
            const items = root.querySelectorAll(
                ".more-items-popUp .izmo-emporio-trim.item"
            );

            for (const el of items) {
                if (!visible(el)) continue;

                const attr = clean(el.getAttribute("trim"));
                const text = clean(
                    el.innerText || el.textContent || ""
                );

                if (same(attr, expected) || same(text, expected)) {
                    return el;
                }
            }
        } catch (e) {}

        try {
            const all = root.querySelectorAll
                ? Array.from(root.querySelectorAll("*"))
                : [];

            for (const el of all) {
                if (el.shadowRoot) {
                    const found = search(
                        el.shadowRoot,
                        expected
                    );
                    if (found) return found;
                }
            }
        } catch (e) {}

        return null;
    }

    const el = search(document, arguments[0]);

    if (!el) {
        return {
            found: false,
            clicked: false,
            attr: "",
            text: ""
        };
    }

    try {
        el.scrollIntoView({
            behavior: "instant",
            block: "center",
            inline: "center"
        });
    } catch (e) {}

    let clicked = false;

    try {
        el.click();
        clicked = true;
    } catch (e) {}

    if (!clicked) {
        try {
            el.dispatchEvent(
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
        found: true,
        clicked: clicked,
        attr: clean(el.getAttribute("trim")),
        text: clean(el.innerText || el.textContent || "")
    };
    """

    try:
        result = driver.execute_script(
            script,
            expected
        ) or {}
    except Exception as e:
        print(
            f"TRIM SELECT SCRIPT ERROR [{expected}]:",
            e
        )
        return False

    if not result.get("found") or not result.get("clicked"):
        print(
            f"TRIM SELECT: exact trim not found/clicked -> '{expected}'"
        )
        return False

    print(
        f"TRIM SELECTED: '{expected}'"
    )

    # Wait for the vehicle UI to react. We deliberately do not require the
    # header text to change because some Emporio builds update it late or
    # expose it through a different shadow root.
    time.sleep(1.2)

    # Best-effort confirmation using the base-trim label.
    end_time = time.time() + timeout
    while time.time() < end_time:
        try:
            current = get_text(
                driver,
                SEL_BASE_TRIM
            ) or ""

            current_norm = normalize_text(current)
            expected_norm = normalize_text(expected)

            if (
                expected_norm
                and (
                    expected_norm == current_norm
                    or expected_norm in current_norm
                    or current_norm in expected_norm
                )
            ):
                print(
                    f"TRIM CONFIRMED IN HEADER: '{current}'"
                )
                return True
        except Exception:
            pass

        # The click itself already succeeded; don't block the run for the
        # optional confirmation.
        time.sleep(0.25)

    print(
        f"TRIM CLICK SUCCEEDED; header confirmation not observed: '{expected}'"
    )
    return True



def _get_shadow_root(element):
    """
    Return Selenium 4's open ShadowRoot for a custom element.
    Returns None when the element has no accessible/open shadow root.
    """
    try:
        return element.shadow_root
    except Exception:
        return None


def _find_all_in_open_shadow_dom(driver, selector):
    """
    Selenium-only recursive search.

    This intentionally does NOT use execute_script() for color extraction.
    It uses Selenium 4 WebElement.shadow_root / ShadowRoot.find_elements()
    so the More Colors logic is handled with Selenium + Python.
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

        # ---------------------------------------------------------
        # Find the requested selector in this root.
        # ---------------------------------------------------------
        try:
            for element in root.find_elements(By.CSS_SELECTOR, selector):
                add_unique(element)
        except Exception:
            pass

        # ---------------------------------------------------------
        # Find every element in this root and enter its OPEN
        # ShadowRoot when available.
        # ---------------------------------------------------------
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


def _find_first_in_open_shadow_dom(driver, selector):
    elements = _find_all_in_open_shadow_dom(
        driver,
        selector
    )
    return elements[0] if elements else None


def _clean_color_name(value):
    return " ".join(
        str(value or "").split()
    ).strip()


def _unique_color_names(names):
    """
    Remove duplicate color names case-insensitively while preserving
    the original page order and original capitalization.
    """
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


def _get_visible_element(element):
    """
    Selenium visibility check used only for the NO-MORE-COLORS fallback.
    The expanded More Colors popup deliberately does NOT use this check.
    """
    try:
        return element.is_displayed()
    except Exception:
        return False


def _get_color_name_from_swatch(swatch):
    """
    Extract exactly one data-color-name from one .color-swatch.
    """
    try:
        color_elements = swatch.find_elements(
            By.CSS_SELECTOR,
            "[data-color-name]"
        )
    except Exception:
        color_elements = []

    if not color_elements:
        return ""

    try:
        return _clean_color_name(
            color_elements[0].get_attribute(
                "data-color-name"
            )
        )
    except Exception:
        return ""


def _collect_popup_colors_selenium(driver):
    """
    Selenium-only More Colors extraction.

    Exact structure supplied by the user:

        .color-container
            .color-swatch
                [data-color-name]

    No visibility test is applied to .color-swatch because the Emporio
    popup can contain all colors in the DOM while individual swatches
    report unusual/zero dimensions.
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

        # Primary extraction: one data-color-name per swatch.
        for swatch in swatches:
            name = _get_color_name_from_swatch(
                swatch
            )

            if name:
                colors.append(name)

        # Fallback for a build where .color-swatch is not used.
        if not swatches:
            try:
                color_elements = container.find_elements(
                    By.CSS_SELECTOR,
                    "[data-color-name]"
                )
            except Exception:
                color_elements = []

            for color_element in color_elements:
                try:
                    name = _clean_color_name(
                        color_element.get_attribute(
                            "data-color-name"
                        )
                    )
                except Exception:
                    name = ""

                if name:
                    colors.append(name)

    return _unique_color_names(colors)


def _collect_visible_colors_selenium(driver):
    """
    Fallback when More Colors does not exist.

    Emporio variants are inconsistent about the dimensions reported by
    individual color swatches. Therefore the reliable visibility boundary
    is the current ``.color-section``; once that section is visible, read
    every ``data-color-name`` inside it even if an individual swatch reports
    zero width/height. This prevents valid colors from becoming 0 merely
    because the swatch itself is rendered as a custom element.
    """
    colors = []

    sections = _find_all_in_open_shadow_dom(
        driver,
        ".color-section"
    )

    for section in sections:
        if not _get_visible_element(section):
            continue

        try:
            elements = section.find_elements(
                By.CSS_SELECTOR,
                "[data-color-name]"
            )
        except Exception:
            elements = []

        for element in elements:
            try:
                name = _clean_color_name(
                    element.get_attribute(
                        "data-color-name"
                    ) or ""
                )
            except Exception:
                name = ""

            if name:
                colors.append(name)

    # Some builds expose the color-name attribute directly on the swatch.
    if not colors:
        try:
            swatches = _find_all_in_open_shadow_dom(
                driver,
                ".color-section .color-swatch"
            )
            for swatch in swatches:
                try:
                    direct_name = _clean_color_name(
                        swatch.get_attribute("data-color-name") or ""
                    )
                except Exception:
                    direct_name = ""
                if direct_name:
                    colors.append(direct_name)
        except Exception:
            pass

    return _unique_color_names(colors)


def _click_more_colors_with_selenium(driver):
    """
    Find the exact More Colors structure and click the
    .moreColors.action parent using Selenium.

    Exact structure:

        .color-section
            > .moreColors.action
                > .moreColorsImg
    """
    image_elements = _find_all_in_open_shadow_dom(
        driver,
        ".color-section > .moreColors.action > .moreColorsImg"
    )

    for image_element in image_elements:
        try:
            parent = image_element.find_element(
                By.XPATH,
                ".."
            )
        except Exception:
            parent = None

        if parent is None:
            continue

        try:
            classes = (
                parent.get_attribute("class")
                or ""
            ).split()

            if (
                "moreColors" not in classes
                or "action" not in classes
            ):
                continue
        except Exception:
            continue

        try:
            if not parent.is_displayed():
                continue
        except Exception:
            pass

        try:
            driver.execute_script(
                "arguments[0].scrollIntoView({"
                "block:'center',inline:'center'"
                "});",
                parent
            )
        except Exception:
            pass

        # ---------------------------------------------------------
        # Selenium click first.
        # ---------------------------------------------------------
        try:
            parent.click()
            print(
                "  MORE COLORS: Selenium click succeeded"
            )
            return True
        except Exception as first_error:
            print(
                "  MORE COLORS: Selenium click failed:",
                first_error
            )

        # ---------------------------------------------------------
        # ActionChains fallback. Still Selenium/Python.
        # ---------------------------------------------------------
        try:
            ActionChains(driver).move_to_element(
                parent
            ).pause(0.1).click().perform()

            print(
                "  MORE COLORS: ActionChains click succeeded"
            )
            return True
        except Exception as second_error:
            print(
                "  MORE COLORS: ActionChains click failed:",
                second_error
            )

    return False


def get_all_color_names(driver, timeout=25):
    """
    Get all colors for the current trim using Selenium + Python.

    MORE COLORS:
        1. Find the exact:
           .color-section > .moreColors.action > .moreColorsImg
        2. Click the .moreColors.action parent.
        3. Find .color-container.
        4. Find every .color-swatch.
        5. Read one [data-color-name] from each swatch.
        6. Remove duplicates case-insensitively.

    NO MORE COLORS:
        - Read visible [data-color-name] from .color-section.
        - Remove duplicates.

    Important:
        Color extraction itself uses Selenium 4 ShadowRoot APIs.
        No JavaScript DOM-query script is used for the color collection.
    """

    print(
        "  COLOR: Checking exact More Colors selector..."
    )

    more_colors_images = _find_all_in_open_shadow_dom(
        driver,
        ".color-section > .moreColors.action > .moreColorsImg"
    )

    if more_colors_images:
        print(
            "  MORE COLORS: FOUND EXACT PATH "
            ".color-section > .moreColors.action > .moreColorsImg"
        )

        clicked = _click_more_colors_with_selenium(
            driver
        )

        if not clicked:
            print(
                "  MORE COLORS: Could not click the "
                ".moreColors.action parent"
            )
            return []

        print(
            "  MORE COLORS: Waiting for .color-container..."
        )

        deadline = time.time() + timeout
        last_names = []
        previous_names = None
        stable_reads = 0

        while time.time() < deadline:
            try:
                names = _collect_popup_colors_selenium(
                    driver
                )
            except Exception as e:
                print(
                    "  MORE COLORS: Selenium read error:",
                    e
                )
                names = []

            if names:
                print(
                    "  MORE COLORS: "
                    f"{len(names)} color(s) found"
                )

                current_keys = [
                    name.casefold()
                    for name in names
                ]

                if (
                    previous_names is not None
                    and current_keys == previous_names
                ):
                    stable_reads += 1
                else:
                    stable_reads = 1
                    previous_names = current_keys

                last_names = names

                # Two consecutive identical reads mean the popup
                # has stabilized. This also works for models with
                # fewer than 21 colors.
                if stable_reads >= 2:
                    print(
                        "  MORE COLORS: "
                        f"FINAL COUNT = {len(last_names)}"
                    )
                    print(
                        "  MORE COLORS: "
                        f"NAMES = {', '.join(last_names)}"
                    )
                    return last_names

            time.sleep(0.5)

        print(
            "  MORE COLORS: timeout reached"
        )
        print(
            "  MORE COLORS: "
            f"FINAL COUNT = {len(last_names)}"
        )
        print(
            "  MORE COLORS: "
            f"NAMES = {', '.join(last_names)}"
        )

        return last_names

    # -------------------------------------------------------------
    # No More Colors control.
    # -------------------------------------------------------------
    print(
        "  MORE COLORS: NOT PRESENT -> "
        "reading visible [data-color-name]"
    )

    deadline = time.time() + timeout
    last_names = []

    while time.time() < deadline:
        try:
            names = _collect_visible_colors_selenium(
                driver
            )
        except Exception as e:
            print(
                "  VISIBLE COLOR READ ERROR:",
                e
            )
            names = []

        if names:
            last_names = names
            break

        time.sleep(0.5)

    print(
        "  VISIBLE [data-color-name] COUNT:",
        len(last_names)
    )
    print(
        "  COLOR NAMES:",
        ", ".join(last_names)
    )

    return last_names




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



def collect_accessories_and_colors_for_each_trim(
    driver,
    trim_names,
    initial_accessory_names=None
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
                        timeout=25
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
            # Do NOT reuse initial_accessory_names here.
            # Explicitly open the Accessories tab and collect the
            # accessory names/count for the currently selected trim.
            # This keeps single-trim and multi-trim collection
            # consistent and avoids an empty initial list.
            # --------------------------------------------------
            print(
                f"[{trim}] SINGLE TRIM -> "
                "FETCHING ACCESSORIES FROM ACCESSORIES TAB"
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

            # Close the tab before finishing this trim.
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

    # ------------------------------------------------------
    # MODEL-LEVEL ACCESSORY FALLBACK FOR ALL TRIMS
    #
    # If the Accessories tab was successfully read at model level
    # but the per-trim accessory scan returned no names for ANY trim,
    # the model-level accessory list represents accessories available
    # across the model's trims. Use it for every trim so the
    # Accessories sheet can correctly output "ALL".
    #
    # This is only used when ALL per-trim lists are empty. We do not
    # overwrite partial per-trim data.
    # ------------------------------------------------------
    if (
        trim_count > 1
        and initial_accessory_names
        and all(
            not (
                accessories_by_trim.get(trim, [])
                or []
            )
            for trim in ordered_trims
        )
    ):
        print("")
        print(
            "PER-TRIM ACCESSORY FALLBACK: "
            "all trim accessory lists are empty."
        )
        print(
            "Using the complete initial model-level accessory list "
            "for every trim."
        )

        for trim in ordered_trims:
            accessories_by_trim[trim] = list(
                initial_accessory_names
            )

            if not accessory_counts_by_trim.get(trim):
                accessory_counts_by_trim[trim] = len(
                    initial_accessory_names
                )

            print(
                f"[{trim}] FALLBACK ACCESSORY NAMES = "
                f"{accessories_by_trim[trim]}"
            )

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
    """
    SIMPLE / EXACT ACCESSORY EXTRACTION

    Flow:
      1. Verify and click Accessories tab.
      2. Work inside the Emporio accessory container.
      3. Read every accessory card currently present in that container.
         Visibility is NOT required.
      4. If the Emporio next arrow exists, click:
             svg.nextbutton.action
         and read the container again.
      5. Continue until there is no next arrow or no new cards appear.
      6. After ALL main cards are collected, process numbered-card children.
      7. Return the logical accessory count and all accessory names.

    The important point is that carousel navigation happens BEFORE opening
    numbered-card popups. This prevents a child popup from interfering with
    the next-arrow click.
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
        print("ACCESSORIES TAB: click failed")
        return status, 0, []

    print(f"ACCESSORIES TAB CLICKED using strategy '{click_label}'")
    time.sleep(1.5)

    # -------------------------------------------------------------
    # Exact Emporio accessory container.
    #
    # User supplied:
    #   #izmo056e-45d9-4af0-a62e-2bd1bb6761fe
    #       .shadowRoot
    #   #izmo5b11-262a-483e-a4ae-93738b24c989
    #
    # These IDs can be regenerated by some builds, so the script also
    # has a fallback that searches for the accessory section.
    # -------------------------------------------------------------
    collect_accessory_set_script = r"""
    function clean(v) {
        return String(v || "").replace(/\s+/g, " ").trim();
    }

    function findRootContainer() {
        // Exact structure supplied by the user.
        const host = document.querySelector(
            "#izmo056e-45d9-4af0-a62e-2bd1bb6761fe"
        );

        if (host && host.shadowRoot) {
            const exact = host.shadowRoot.querySelector(
                "#izmo5b11-262a-483e-a4ae-93738b24c989"
            );
            if (exact) return exact;
        }

        // Fallback: find the accessory section anywhere in open shadow DOM.
        function walk(root) {
            if (!root) return null;

            try {
                const exact = root.querySelector(
                    "#izmo5b11-262a-483e-a4ae-93738b24c989"
                );
                if (exact) return exact;
            } catch (e) {}

            try {
                const sections = root.querySelectorAll(
                    ".izmo-emporio-accessory-section"
                );
                if (sections.length) return sections[0];
            } catch (e) {}

            try {
                const all = root.querySelectorAll("*");
                for (const el of all) {
                    if (el.shadowRoot) {
                        const found = walk(el.shadowRoot);
                        if (found) return found;
                    }
                }
            } catch (e) {}

            return null;
        }

        return walk(document);
    }

    function getName(card) {
        let el = null;

        try {
            el = card.querySelector(
                ".item-name.fontFamilyRegular.fontSize12"
            );
        } catch (e) {}

        if (!el) {
            try { el = card.querySelector(".item-name"); } catch (e) {}
        }

        if (!el) {
            try {
                const all = card.querySelectorAll("*");
                for (const x of all) {
                    if (
                        x.matches &&
                        x.matches(
                            ".item-name.fontFamilyRegular.fontSize12, .item-name"
                        )
                    ) {
                        el = x;
                        break;
                    }
                }
            } catch (e) {}
        }

        return clean(
            el ? (el.innerText || el.textContent || "") : ""
        );
    }

    function getNumber(card) {
        const values = [];

        function add(v) {
            const s = clean(v);
            if (/^\d+$/.test(s)) {
                values.push(parseInt(s, 10));
            }
        }

        try {
            for (const attr of Array.from(card.attributes || [])) {
                const n = String(attr.name || "").toLowerCase();
                if (
                    n.includes("count") ||
                    n.includes("number") ||
                    n.includes("quantity") ||
                    n.includes("total")
                ) {
                    add(attr.value);
                }
            }
        } catch (e) {}

        try {
            const all = card.querySelectorAll("*");
            for (const el of all) {
                for (const attr of Array.from(el.attributes || [])) {
                    const n = String(attr.name || "").toLowerCase();
                    if (
                        n.includes("count") ||
                        n.includes("number") ||
                        n.includes("quantity") ||
                        n.includes("total")
                    ) {
                        add(attr.value);
                    }
                }

                const txt = clean(
                    el.innerText || el.textContent || ""
                );
                if (/^\d+$/.test(txt)) add(txt);

                for (const pseudo of ["::before", "::after"]) {
                    try {
                        let c = getComputedStyle(el, pseudo).content;
                        c = String(c || "")
                            .replace(/^['"]|['"]$/g, "")
                            .trim();
                        add(c);
                    } catch (e) {}
                }
            }
        } catch (e) {}

        return values.length ? Math.max(...values) : null;
    }

    function collect(root) {
        if (!root) {
            return {
                found: false,
                cards: [],
                arrow: false,
                lastindex: null,
                itemslength: null
            };
        }

        const cards = [];
        const seen = new Set();

        function add(card) {
            if (!card || seen.has(card)) return;

            let id = "";
            try { id = card.getAttribute("accessoryid") || ""; }
            catch (e) {}

            if (!id) return;

            seen.add(card);

            cards.push({
                accessoryid: id,
                trim: card.getAttribute("trim") || "",
                itemName: getName(card),
                number: getNumber(card),
                elementId: card.id || ""
            });
        }

        try {
            const selectors = [
                ".izmo-emporio-accessory.item.action[accessoryid]",
                ".izmo-emporio-accessory.item[accessoryid]",
                "izmo-emporio-accessory.item.action[accessoryid]",
                "izmo-emporio-div.izmo-emporio-accessory.item.action[accessoryid]"
            ];

            for (const selector of selectors) {
                for (const card of root.querySelectorAll(selector)) {
                    add(card);
                }
            }
        } catch (e) {}

        // Search nested open shadow roots too.
        function walk(r) {
            if (!r) return;

            try {
                for (const el of r.querySelectorAll("*")) {
                    try {
                        if (
                            el.matches &&
                            el.matches(
                                ".izmo-emporio-accessory.item.action[accessoryid], " +
                                ".izmo-emporio-accessory.item[accessoryid], " +
                                "izmo-emporio-accessory.item.action[accessoryid], " +
                                "izmo-emporio-div.izmo-emporio-accessory.item.action[accessoryid]"
                            )
                        ) {
                            add(el);
                        }
                    } catch (e) {}

                    if (el.shadowRoot) {
                        walk(el.shadowRoot);
                    }
                }
            } catch (e) {}
        }

        walk(root);

        let arrow = null;
        try {
            const arrows = root.querySelectorAll(
                "svg.nextbutton.action"
            );

            for (const a of arrows) {
                const disabled =
                    a.classList.contains("disabled") ||
                    a.getAttribute("disabled") !== null ||
                    a.getAttribute("aria-disabled") === "true";

                if (!disabled) {
                    arrow = a;
                    break;
                }
            }
        } catch (e) {}

        let lastindex = null;
        let itemslength = null;

        if (arrow) {
            const li = arrow.getAttribute("lastindex");
            const il = arrow.getAttribute("itemslength");

            if (/^\d+$/.test(String(li || ""))) {
                lastindex = parseInt(li, 10);
            }

            if (/^\d+$/.test(String(il || ""))) {
                itemslength = parseInt(il, 10);
            }
        }

        return {
            found: true,
            cards: cards,
            arrow: !!arrow,
            lastindex: lastindex,
            itemslength: itemslength
        };
    }

    const root = findRootContainer();

    if (!root) {
        return {
            found: false,
            cards: [],
            arrow: false,
            lastindex: null,
            itemslength: null
        };
    }

    return collect(root);
    """

    click_next_arrow_script = r"""
    function findRootContainer() {
        const host = document.querySelector(
            "#izmo056e-45d9-4af0-a62e-2bd1bb6761fe"
        );

        if (host && host.shadowRoot) {
            const exact = host.shadowRoot.querySelector(
                "#izmo5b11-262a-483e-a4ae-93738b24c989"
            );
            if (exact) return exact;
        }

        function walk(root) {
            if (!root) return null;

            try {
                const exact = root.querySelector(
                    "#izmo5b11-262a-483e-a4ae-93738b24c989"
                );
                if (exact) return exact;
            } catch (e) {}

            try {
                const sec = root.querySelector(
                    ".izmo-emporio-accessory-section"
                );
                if (sec) return sec;
            } catch (e) {}

            try {
                for (const el of root.querySelectorAll("*")) {
                    if (el.shadowRoot) {
                        const found = walk(el.shadowRoot);
                        if (found) return found;
                    }
                }
            } catch (e) {}

            return null;
        }

        return walk(document);
    }

    const root = findRootContainer();

    if (!root) {
        return {clicked: false, reason: "container-not-found"};
    }

    let arrow = null;

    try {
        const arrows = root.querySelectorAll(
            "svg.nextbutton.action"
        );

        for (const a of arrows) {
            const disabled =
                a.classList.contains("disabled") ||
                a.getAttribute("disabled") !== null ||
                a.getAttribute("aria-disabled") === "true";

            if (!disabled) {
                arrow = a;
                break;
            }
        }
    } catch (e) {}

    if (!arrow) {
        return {clicked: false, reason: "arrow-not-found"};
    }

    const before = {
        lastindex: arrow.getAttribute("lastindex") || "",
        itemslength: arrow.getAttribute("itemslength") || ""
    };

    try {
        arrow.scrollIntoView({
            block: "center",
            inline: "center"
        });
    } catch (e) {}

    let clicked = false;

    try {
        arrow.click();
        clicked = true;
    } catch (e) {}

    if (!clicked) {
        try {
            arrow.dispatchEvent(
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
        beforeLastindex: before.lastindex,
        beforeItemslength: before.itemslength
    };
    """

    # -------------------------------------------------------------
    # Collect every MAIN card across carousel sets.
    # -------------------------------------------------------------
    all_cards = []
    seen_main_cards = set()

    def card_key(card):
        aid = " ".join(
            str(card.get("accessoryid", "") or "").split()
        ).strip().casefold()

        trim = " ".join(
            str(card.get("trim", "") or "").split()
        ).strip().casefold()

        name = " ".join(
            str(card.get("itemName", "") or "").split()
        ).strip().casefold()

        return aid, trim, name

    def add_cards_from_set(info, set_number):
        new_count = 0

        for card in info.get("cards", []) or []:
            key = card_key(card)

            if not key[0]:
                continue

            if key in seen_main_cards:
                continue

            seen_main_cards.add(key)
            all_cards.append(card)
            new_count += 1

            print(
                f"  MAIN ACCESSORY {len(all_cards)}:",
                card.get("itemName") or "<NAME NOT FOUND>",
                "| ID=", card.get("accessoryid"),
                "| NUMBER=", card.get("number")
            )

        print(
            f"ACCESSORY SET {set_number}: "
            f"{len(info.get('cards', []) or [])} cards found, "
            f"{new_count} new cards"
        )

        return new_count

    max_sets = 50

    for set_number in range(1, max_sets + 1):
        info = driver.execute_script(
            collect_accessory_set_script
        ) or {}

        if not info.get("found"):
            print("ACCESSORY CONTAINER: NOT FOUND")
            break

        print("")
        print("=" * 60)
        print(
            f"ACCESSORY SET {set_number}: "
            f"CARDS IN CONTAINER = "
            f"{len(info.get('cards', []) or [])}"
        )
        print(
            "NEXT ARROW:",
            info.get("arrow"),
            "| lastindex=",
            info.get("lastindex"),
            "| itemslength=",
            info.get("itemslength")
        )
        print("=" * 60)

        new_count = add_cards_from_set(
            info,
            set_number
        )

        # No arrow -> all available carousel cards collected.
        if not info.get("arrow"):
            print("NO NEXT ARROW -> ACCESSORY COLLECTION COMPLETE")
            break

        # If the carousel tells us it is already at the final index,
        # don't click again.
        lastindex = info.get("lastindex")
        itemslength = info.get("itemslength")

        if (
            lastindex is not None
            and itemslength is not None
            and lastindex >= itemslength
        ):
            print(
                "LAST INDEX REACHED:",
                lastindex,
                "/",
                itemslength
            )
            break

        before_key = tuple(
            card_key(x)
            for x in info.get("cards", []) or []
        )

        click_result = driver.execute_script(
            click_next_arrow_script
        ) or {}

        print(
            "NEXT ARROW CLICK RESULT:",
            click_result
        )

        if not click_result.get("clicked"):
            print(
                "NEXT ARROW COULD NOT BE CLICKED -> STOP"
            )
            break

        # Wait for the carousel DOM to change.
        changed = False
        end_time = time.time() + 8

        while time.time() < end_time:
            time.sleep(0.25)

            after = driver.execute_script(
                collect_accessory_set_script
            ) or {}

            after_key = tuple(
                card_key(x)
                for x in after.get("cards", []) or []
            )

            if after_key != before_key:
                changed = True
                break

        if not changed:
            print(
                "NEXT ARROW CLICKED BUT ACCESSORY CARDS "
                "DID NOT CHANGE -> STOP"
            )
            break

    # -------------------------------------------------------------
    # IMPORTANT:
    # ALL carousel navigation is finished before we open any numbered
    # accessory popup. This prevents popup state from blocking the arrow.
    # -------------------------------------------------------------
    print("")
    print("=" * 60)
    print(
        "TOTAL MAIN ACCESSORY CARDS COLLECTED:",
        len(all_cards)
    )
    print("=" * 60)

    if not all_cards:
        return status, 0, []

    accessory_names = []
    accessory_name_keys = set()
    total_accessories = 0

    def add_name_record(record):
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
            key = (
                "FALLBACK",
                len(accessory_names),
                name.casefold()
            )

        if key in accessory_name_keys:
            return False

        accessory_name_keys.add(key)
        accessory_names.append(name)

        print(
            f"  ACCESSORY NAME [{len(accessory_names)}]:",
            name
        )

        return True

    # -------------------------------------------------------------
    # Now process every collected MAIN card.
    # Numbered cards get their children.
    # -------------------------------------------------------------
    for index, card in enumerate(all_cards, start=1):
        main_name = " ".join(
            str(card.get("itemName", "") or "").split()
        ).strip()

        number = card.get("number")
        parent_accessory_id = " ".join(
            str(card.get("accessoryid", "") or "").split()
        ).strip()

        print("")
        print(f"PROCESSING ACCESSORY {index}/{len(all_cards)}")
        print("  NAME:", main_name)
        print("  ID:", parent_accessory_id)
        print("  NUMBER:", number)

        # Parent/main accessory.
        if (
            main_name
            and not re.fullmatch(
                r"Accessory\s+\d+",
                main_name,
                flags=re.IGNORECASE
            )
            and not re.fullmatch(
                r"[0-9a-fA-F]{8}-[0-9a-fA-F-]{4,}",
                main_name
            )
        ):
            add_name_record({
                "name": main_name,
                "accessoryid": parent_accessory_id
            })

        # Children of numbered cards.
        if number is not None:
            print("  NUMBERED ACCESSORY -> FETCHING CHILDREN")

            inner_records = (
                _click_numbered_accessory_card_and_get_names(
                    driver,
                    index - 1,
                    parent_accessory_id,
                    main_name
                )
            )

            print(
                "  CHILDREN FOUND:",
                len(inner_records or [])
            )

            for record in inner_records or []:
                if not isinstance(record, dict):
                    continue

                child_name = " ".join(
                    str(record.get("name", "") or "").split()
                ).strip()

                child_id = " ".join(
                    str(record.get("accessoryid", "") or "").split()
                ).strip()

                # Don't add the parent again if popup inspection returns it.
                if (
                    child_name
                    and child_name.casefold() == main_name.casefold()
                    and child_id.casefold() == parent_accessory_id.casefold()
                ):
                    continue

                add_name_record(record)

        # Existing logical count rule.
        if number is not None:
            try:
                contribution = int(number) + 1
            except Exception:
                contribution = 1
        else:
            contribution = 1

        total_accessories += contribution

    print("")
    print("=" * 60)
    print("FINAL ACCESSORY RESULT")
    print("=" * 60)
    print("TOTAL MAIN ACCESSORY CARDS:", len(all_cards))
    print("TOTAL ACCESSORY NAMES:", len(accessory_names))
    print("NUMBER OF ACCESSORIES:", total_accessories)
    print("ALL ACCESSORY NAMES:")

    for i, name in enumerate(accessory_names, start=1):
        print(f"  {i}. {name}")

    print("=" * 60)

    return (
        status,
        total_accessories,
        accessory_names
    )

# =========================================================
# EXCEL OUTPUT
# =========================================================

EXCEL_COLUMNS = [
    "Sl.no", "Make", "Model", "Base Trim", "Number of Trims",
    "Banner", "Default Color", "IzmoEmporio Text", "Int Hotspot Button",
    "Ext Hotspot Button", "BG Change Button", "Number of BGs",
    "Accessory Tab", "Number of Accessories",
    "Footer Buttons", "Make Logo", "Drag Text Logo",
]

MAIN_COLUMNS = ["Make", "Model", "Other Trims", "Number of Accessories", "Data"]
ACCESSORY_COLUMNS = ["Sl.No", "Make", "Model", "Accessories", "Trims"]
COLOR_COLUMNS = ["Make", "Model", "Other Trims", "Colors Count", "Color Names"]


def _row_values(r):
    return [
        r["number"], r["make"], r["model"], r["base_trim"],
        r["number_of_trims"], r["banner"], r["default_color"], r["izmo_emporio"],
        r["int_hotspot"], r["ext_hotspot"], r["bg_change"],
        r["number_of_backgrounds"], r["accessory_tab"],
        r["number_of_accessories"], r["footer_buttons"], r["make_logo"],
        r["drag_text_logo"],
    ]


def _detail_row_values(r):
    return [
        r["number"], r["make"], r["model"],
        ", ".join(r.get("trim_names", [])),
        ", ".join(r.get("color_names", [])),
        ", ".join(r.get("accessory_names", [])),
    ]


def _main_rows(results):
    rows = []
    for r in results:
        trims = r.get("trim_names", []) or []
        if trims:
            for trim in trims:
                rows.append([
                    r.get("make", ""),
                    r.get("model", ""),
                    trim,
                    r.get("number_of_accessories", 0),
                    RUN_DATE_TIME
                ])
        else:
            rows.append([
                r.get("make", ""),
                r.get("model", ""),
                "",
                r.get("number_of_accessories", 0),
                RUN_DATE_TIME
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
        # Determine trim coverage by ACCESSORY NAME.
        #
        # IMPORTANT:
        #   The Accessories sheet must show ALL when the same accessory
        #   exists in every trim. Do NOT compare accessory position/order.
        #   Compare normalized accessory names across the actual
        #   accessories_by_trim dictionary.
        #
        # Normalization handles:
        #   - leading/trailing whitespace
        #   - repeated whitespace / NBSP
        #   - case differences
        #
        # Also normalize the trim dictionary keys so a harmless difference
        # in whitespace/case cannot make a trim look missing.
        # ----------------------------------------------------------
        def normalize_accessory_key(raw):
            import unicodedata

            value = str(raw or "").replace("\u00a0", " ")
            value = " ".join(value.split()).strip()

            # Normalize accents/diacritics so visually identical names
            # compare correctly across trims.
            value = unicodedata.normalize(
                "NFKD",
                value
            )
            value = "".join(
                ch
                for ch in value
                if not unicodedata.combining(ch)
            )

            return value.casefold()

        def normalize_trim_key(raw):
            value = str(raw or "").replace("\u00a0", " ")
            value = " ".join(value.split()).strip()
            return value.casefold()

        # Build accessory-name sets from the REAL trim results.
        # Do not rely on `selected`, `candidates`, or accessory position.
        trim_name_sets = {}

        for trim in ordered_trims:
            trim_key = normalize_trim_key(trim)
            trim_name_sets[trim_key] = set()

        # Map actual accessories_by_trim keys to the normalized trim key.
        for raw_trim_key, raw_accessories in accessories_by_trim.items():
            trim_key = normalize_trim_key(raw_trim_key)

            if not trim_key:
                continue

            if trim_key not in trim_name_sets:
                # Keep unexpected keys available for diagnostics, but they
                # are not counted as one of the model's ordered trims.
                trim_name_sets[trim_key] = set()

            for raw_name in raw_accessories or []:
                name = valid_name(raw_name)
                name_key = normalize_accessory_key(name)

                if name_key:
                    trim_name_sets[trim_key].add(name_key)

        # Use ONLY the ordered trims when deciding whether coverage is ALL.
        expected_trim_keys = [
            normalize_trim_key(trim)
            for trim in ordered_trims
            if normalize_trim_key(trim)
        ]

        for display_name in selected:
            name_key = normalize_accessory_key(display_name)

            found_trims = []

            for trim in ordered_trims:
                trim_key = normalize_trim_key(trim)

                if name_key in trim_name_sets.get(
                    trim_key,
                    set()
                ):
                    found_trims.append(trim)

            # Accessories sheet requirement:
            # always display ALL in the Trims column.
            trim_output = "ALL"

            rows.append([
                serial,
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


def _write_main_sheet(ws, results):
    """Write Main sheet and highlight the Base Trim from Validation Results."""
    header_font = Font(name="Arial", bold=True, color="FFFFFF")
    header_fill = PatternFill(
        start_color="1F4E78",
        end_color="1F4E78",
        fill_type="solid"
    )
    body_font = Font(name="Arial")
    center = Alignment(horizontal="center", vertical="center")
    wrap_left = Alignment(horizontal="left", vertical="top", wrap_text=True)

    for col_idx, header in enumerate(MAIN_COLUMNS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center

    row_idx = 2

    for result in results:
        base_trim = " ".join(
            str(result.get("base_trim", "") or "").split()
        ).strip().casefold()

        trims = result.get("trim_names", []) or []
        if not trims:
            trims = [""]

        for trim in trims:
            trim_display = " ".join(
                str(trim or "").split()
            ).strip()

            values = [
                result.get("make", ""),
                result.get("model", ""),
                trim_display,
                result.get("number_of_accessories", 0),
                RUN_DATE_TIME
            ]

            for col_idx, value in enumerate(values, start=1):
                cell = ws.cell(row=row_idx, column=col_idx, value=value)
                cell.font = body_font
                cell.alignment = (
                    wrap_left
                    if isinstance(value, str) and len(value) > 40
                    else center
                )

            if base_trim and trim_display.casefold() == base_trim:
                trim_cell = ws.cell(row=row_idx, column=3)
                trim_cell.fill = PatternFill(
                    start_color="FFF2CC",
                    end_color="FFF2CC",
                    fill_type="solid"
                )
                trim_cell.font = Font(name="Arial", bold=True)

            row_idx += 1

    for col_idx, header in enumerate(MAIN_COLUMNS, start=1):
        max_len = len(str(header))
        for row in ws.iter_rows(
            min_row=2,
            max_row=max(2, row_idx - 1),
            min_col=col_idx,
            max_col=col_idx
        ):
            max_len = max(max_len, len(str(row[0].value or "")[:60]))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 4, 60)


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
    _write_main_sheet(
        ws2,
        results
    )

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
        f"Run date/time: {RUN_DATE_TIME}"
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
                row[1] == make and
                row[2] == model
            )
        )

        passed = (
            expected == actual
        )

        if not passed:
            all_pass = False

        print(
            f"{make} {model}: "
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

# IMPORTANT:
# Do not abort the Selenium run when requests.head()/requests.get() is
# rejected by the web server. Many Emporio/CDN servers block Python
# requests, close the connection, or behave differently from Chrome.
# Selenium/Chrome is the real validator for this automation.
url_valid = False
try:
    request_headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/153.0.0.0 Safari/537.36"
        ),
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "en-US,en;q=0.9",
        "Connection": "keep-alive",
    }
    resp = requests.get(
        URL,
        headers=request_headers,
        allow_redirects=True,
        timeout=20,
        verify=False
    )
    print("URL           :", URL)
    print("STATUS CODE   :", resp.status_code)
    url_valid = resp.status_code < 400
except requests.RequestException as e:
    print("URL REQUEST WARNING:", e)
    print("Python requests could not validate the URL.")
    print("This is NOT treated as a failure; Selenium will validate the page.")
    url_valid = True

print("URL VALID     :", "PASS (requests)" if url_valid else "FAIL")

# Never stop here merely because Python requests was rejected.
# The browser will be opened below and will perform the actual page check.


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

        # Universal Emporio parsing: these France-style pages do not
        # need a Year column. Parse the displayed edition/header text
        # as Make + Model. The first token is treated as the make and
        # the remaining text as the model. If the site exposes a
        # different header format, preserve the full text as Model.
        year = ""
        make = ""
        vehicle_model = ""
        model_text = get_edition_info_text(driver) or ""
        model_text = " ".join(model_text.split())
        if model_text:
            parts = model_text.split()
            if len(parts) >= 2:
                make = parts[0]
                vehicle_model = " ".join(parts[1:]).strip()
            else:
                vehicle_model = model_text
        print("Make / Model       :", make, vehicle_model)

        base_trim = get_text(driver, SEL_BASE_TRIM) or ""
        print("Base Trim           :", base_trim or "FAIL (not found)")

        # ---------------------------------------------------------
        # IMPORTANT: wait for the vehicle/color component to finish
        # rendering before reading the default color.
        #
        # This is especially important when multiple makes are selected,
        # because the previous model's UI can remain visible briefly while
        # the new model is loading.
        # ---------------------------------------------------------
        print("Waiting for default color component to render...")
        time.sleep(CAR_RENDER_WAIT)

        default_color = check_default_color(
            driver,
            timeout=DEFAULT_COLOR_TIMEOUT
        ) or ""

        print(
            "Default Color       :",
            default_color or "FAIL (not found)"
        )

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
        number_of_trims, trim_names = check_number_of_trims(
            driver,
            fallback_base_trim=base_trim
        )

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
            accessory_names
        )

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
            "trim_names": trim_names,
            "color_names": color_names,
            "accessory_names": accessory_names,

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

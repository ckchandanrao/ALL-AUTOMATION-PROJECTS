from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time

# ==================================================
# TITLE CASE FUNCTION
# ==================================================

def tc(value):

    if isinstance(value, str):
        return value.title().strip()

    return value




# # ================= INPUT =================
# INPUT_MODEL = input("Enter Model: ").strip().lower()
# INPUT_YEAR = input("Enter Year: ").strip()

# ================= MULTIPLE INPUT =================

all_inputs = []

total = int(input("How many models you want to run: "))

for i in range(total):

    print(f"\nEnter Details {i+1}")

    model = input("Enter Model: ").strip().lower()
    year = input("Enter Year: ").strip()

    all_inputs.append({
        "model": model,
        "year": year
    })







driver = webdriver.Chrome()
driver.maximize_window()
wait = WebDriverWait(driver, 25)

# ---------------- LOGIN ----------------
driver.get("https://sandbox-admin.partsgorilla.com/login")

wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='text']"))).send_keys("izmosupport")
driver.find_element(By.XPATH, "//input[@type='password']").send_keys("adm!n@supp@rt")
driver.find_element(By.XPATH, "//input[@type='submit']").click()

# ---------------- NAVIGATION ----------------
wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="sideNavMenu"]/li[3]/a/span[2]'))).click()
#wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="modeldt_column_filter"]'))).click()



# ==================================================
# FINAL MERGED DATA
# ==================================================

all_main_data = []
all_trim_data = []
all_final_accessories = []



for entry in all_inputs:

    INPUT_MODEL = entry["model"]
    INPUT_YEAR = entry["year"]

    print(f"\n🚀 Running => {INPUT_MODEL} | {INPUT_YEAR}")


    # ==================================================
    # REOPEN MODEL PAGE + FILTER PANEL
    # ==================================================

    # driver.get("https://sandbox-admin.partsgorilla.com/model")

    # wait.until(EC.presence_of_element_located((
    #     By.XPATH,
    #     '//*[@id="modeldt"]'
    # )))

    # time.sleep(2)

    # ==================================================
# FULL RESET BEFORE EVERY MODEL
# ==================================================

    driver.get("https://sandbox-admin.partsgorilla.com/model")

    # wait full page load
    wait.until(EC.presence_of_element_located((
        By.XPATH,
        '//*[@id="modeldt"]'
    )))

    # wait loading overlay disappear
    try:
        wait.until(EC.invisibility_of_element_located((
            By.XPATH,
            '//*[@id="modeldt_processing"]'
        )))
    except:
        pass

    time.sleep(5)






        # ==================================================
    # OPEN FILTER PANEL AGAIN
    # ==================================================

    wait.until(EC.presence_of_element_located((
        By.XPATH,
        '//*[@id="modeldt_column_filter"]'
    )))

    time.sleep(3)

   

        # ==================================================
    # FORCE OPEN FILTER PANEL
    # ==================================================

    filter_btn = driver.find_element(
        By.XPATH,
        '//*[@id="modeldt_column_filter"]'
    )

    # click until filter items appear
    for _ in range(3):

        driver.execute_script(
            "arguments[0].click();",
            filter_btn
        )

        time.sleep(2)

        try:
            if driver.find_element(By.ID, "filter-items").is_displayed():
                break
        except:
            pass

    time.sleep(2)



    # 🔥 MERGE STRUCTURE (for accessories)
    accessory_map = {}
    trim_data = []
    model_color_set = set()

   

    wait.until(EC.presence_of_element_located((
        By.ID,
        "filter-items"
    )))

    time.sleep(2)



    # ==================================================
    # OPEN MODEL FILTER
    # ==================================================

    time.sleep(3)

    model_filter = driver.find_element(
        By.XPATH,
        '//*[@id="filter-items"]/div[3]/span/span/span'
    )

    driver.execute_script("""
    arguments[0].dispatchEvent(
        new MouseEvent('click', {
            bubbles: true,
            cancelable: true,
            view: window
        })
    );
    """, model_filter)

    time.sleep(2)





    # ==================================================
    # MODEL INPUT
    # ==================================================



    model_input = wait.until(EC.presence_of_element_located((
        By.XPATH,
        '//*[@id="modeldt_Model"]/div/div/input'
    )))

    # force visible
    driver.execute_script(
        "arguments[0].click();",
        model_input
    )

    time.sleep(1)

    model_input.clear()

    time.sleep(1)

    model_input.send_keys(INPUT_MODEL)

    time.sleep(1)

    # set value using JS
    driver.execute_script("""
    arguments[0].value = arguments[1];
    arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
    arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
    """, model_input, INPUT_MODEL)

    time.sleep(1)





    driver.find_element(
        By.XPATH,
        '//*[@id="modeldt_Model"]/div/div/button'
    ).click()

    time.sleep(2)


    # ==================================================
    # OPEN YEAR FILTER
    # ==================================================

    year_filter = driver.find_element(
        By.XPATH,
        '//*[@id="filter-items"]/div[5]/span/span/span'
    )

    driver.execute_script("""
    arguments[0].dispatchEvent(
        new MouseEvent('click', {
            bubbles: true,
            cancelable: true,
            view: window
        })
    );
    """, year_filter)

    time.sleep(2)


    # ==================================================
    # YEAR INPUT
    # ==================================================

    year_input = wait.until(EC.presence_of_element_located((
        By.XPATH,
        '//*[@id="modeldt_Year"]/div/div/input'
    )))

    

    driver.execute_script(
        "arguments[0].click();",
        year_input
    )

    time.sleep(1)

    year_input.clear()

    time.sleep(1)

    year_input.send_keys(INPUT_YEAR)

    time.sleep(1)

    driver.execute_script("""
    arguments[0].value = arguments[1];
    arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
    arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
    """, year_input, INPUT_YEAR)

    time.sleep(1)


    driver.find_element(
        By.XPATH,
        '//*[@id="modeldt_Year"]/div/div/button'
    ).click()

    # ==================================================
    # WAIT FOR TABLE
    # ==================================================

    wait.until(EC.invisibility_of_element_located((
        By.XPATH,
        '//*[@id="modeldt_processing"]'
    )))

    time.sleep(2)





    # ---------------- COUNTRY ----------------
    country = driver.find_element(By.XPATH, '//*[@id="modeldt"]/tbody/tr[1]/td[2]').text.strip()

    # ---------------- CLICK MODEL ----------------
    for row in driver.find_elements(By.XPATH, '//*[@id="modeldt"]/tbody/tr'):
        try:
            if row.find_element(By.XPATH, "./td[3]").text.lower() == INPUT_MODEL:
                row.find_element(By.XPATH, "./td[3]/a").click()
                break
        except:
            continue

    # ---------------- MODEL DETAILS ----------------
    make = wait.until(EC.visibility_of_element_located((By.XPATH,'//*[@id="nav-7-1"]/div/div[2]/div[2]'))).text
    model = wait.until(EC.visibility_of_element_located((By.XPATH,'//*[@id="nav-7-1"]/div/div[3]/div[2]'))).text
    year = wait.until(EC.visibility_of_element_located((By.XPATH,'//*[@id="nav-7-1"]/div/div[5]/div[2]'))).text


    # ==================================================
    # MODEL COLORS
    # ==================================================
    color_tab_model = wait.until(EC.element_to_be_clickable((
        By.XPATH,
        '/html/body/main/div[2]/div[2]/div/ui-view/div/div[2]/div[2]/header/div/div/ul/li[2]/a'
    )))
    driver.execute_script("arguments[0].click();", color_tab_model)

    time.sleep(2)

    while True:

        wait.until(EC.presence_of_element_located((
            By.XPATH,
            '//*[@id="modelColorDt"]/tbody'
        )))

        time.sleep(1)

        rows = driver.find_elements(
            By.XPATH,
            '//*[@id="modelColorDt"]/tbody/tr'
        )

        for r in rows:

            cols = r.find_elements(By.TAG_NAME, "td")

            if not cols:
                continue

            color = cols[0].text.strip()

            # Skip invalid rows silently
            if (
                not color
                or color.upper() == "NA"
                or "NO RECORDS FOUND" in color.upper()
            ):
                continue

            model_color_set.add(color)





        # ---------------- PAGINATION ----------------
        try:
            next_btn = driver.find_element(
                By.XPATH,
                '//*[@id="modelColorDt_next"]'
            )

            if "disabled" not in next_btn.get_attribute("class"):

                old = rows[0]

                driver.execute_script(
                    "arguments[0].click();",
                    next_btn
                )

                wait.until(EC.staleness_of(old))

            else:
                break

        except:
            break

    print("✅ Model colors done")









    # ==================================================
    # MODEL ACCESSORIES
    # ==================================================
    acc_tab_model = wait.until(EC.element_to_be_clickable((
        By.XPATH,
        '/html/body/main/div[2]/div[2]/div/ui-view/div/div[2]/div[2]/header/div/div/ul/li[4]/a'
    )))
    driver.execute_script("arguments[0].click();", acc_tab_model)

    time.sleep(2)

    model_accessories = []

    while True:
        # wait for table
        wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="modelAccessories"]/tbody')))
        time.sleep(1)

        rows = driver.find_elements(By.XPATH,'//*[@id="modelAccessories"]/tbody/tr')

        for r in rows:
            cols = r.find_elements(By.TAG_NAME,"td")

            # skip invalid rows
            if len(cols) < 5 or "No data" in r.text:
                continue

            part = cols[1].text.strip()

            accessory_map[part] = {
                "Category": cols[2].text.strip(),
                "Status": cols[4].text.strip(),
                "Trims": {"ALL"}
            }

        # -------- PAGINATION --------
        try:
            next_btn = driver.find_element(By.XPATH, '//*[@id="modelAccessories_next"]')

            if "disabled" not in next_btn.get_attribute("class"):
                old = rows[0]
                driver.execute_script("arguments[0].click();", next_btn)
                wait.until(EC.staleness_of(old))
            else:
                break
        except:
            break

    # if no accessories at all
    if not accessory_map:
        accessory_map["NA"] = {"Category":"NA","Status":"NA","Trims":{"ALL"}}

    print("✅ Model accessories done")

    # ==================================================
    # GO TO TRIMS
    # ==================================================
    trim_tab = wait.until(EC.element_to_be_clickable((
        By.XPATH,
        '/html/body/main/div[2]/div[2]/div/ui-view/div/div[2]/div[2]/header/div/div/ul/li[3]/a'
    )))
    driver.execute_script("arguments[0].click();", trim_tab)

    time.sleep(2)

    # ==================================================
    # TRIMS LOOP (COLORS + ACCESSORIES)
    # ==================================================
    while True:
        trim_rows = driver.find_elements(By.XPATH,'//*[@id="trimDt"]/tbody/tr')
        if not trim_rows:
            break

        i = 0
        while i < len(trim_rows):
            try:
                trim_rows = driver.find_elements(By.XPATH,'//*[@id="trimDt"]/tbody/tr')
                row = trim_rows[i]

                trim_name = row.find_element(By.XPATH,"./td[2]/a").text.strip()
                trim_status = row.find_element(By.XPATH,"./td[6]/span").text.strip()

                # 🚫 Skip inactive trims
                if trim_status.lower() != "active":
                    i += 1
                    continue

                row.find_element(By.XPATH,"./td[2]/a").click()

                # -------- TRIM COLORS --------
                color_tab_trim = wait.until(EC.element_to_be_clickable((
                    By.XPATH,
                    '/html/body/main/div[2]/div[2]/div/ui-view/div/div[2]/div[2]/div[1]/header/div/div/ul/li[2]/a'
                )))
                driver.execute_script("arguments[0].click();", color_tab_trim)

                time.sleep(1)

                trim_colors = []

                while True:
                    # wait for table
                    wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="trimColorDt"]/tbody')))
                    time.sleep(1)

                    rows = driver.find_elements(By.XPATH, '//*[@id="trimColorDt"]/tbody/tr')

                    for r in rows:
                        cols = r.find_elements(By.TAG_NAME, "td")
                        if cols:
                            color = cols[0].text.strip()
                            
                            # if color and color.upper() != "NA":
                            #     trim_colors.append(color)
                            color = cols[0].text.strip()

                        # Skip invalid rows silently
                        if (
                            not color
                            or color.upper() == "NA"
                            or "NO RECORDS FOUND" in color.upper()
                        ):
                            continue

                        trim_colors.append(color)






        # -------- PAGINATION --------
                    try:
                        next_btn = driver.find_element(By.XPATH, '//*[@id="trimColorDt_next"]')

                        if "disabled" not in next_btn.get_attribute("class"):
                            old = rows[0]
                            driver.execute_script("arguments[0].click();", next_btn)
                            wait.until(EC.staleness_of(old))
                        else:
                            break
                    except:
                        break

                # remove duplicates
                trim_colors = list(set(trim_colors))             



                #combined_colors = sorted(set(trim_colors).union(model_color_set))
                # ensure trim_colors is always a list
                trim_colors = trim_colors if trim_colors else []

                # always include model colors
                combined_colors = sorted(set(model_color_set).union(set(trim_colors)))

                # -------- TRIM ACCESSORIES --------
                acc_tab_trim = wait.until(EC.element_to_be_clickable((
                    By.XPATH,
                    '/html/body/main/div[2]/div[2]/div/ui-view/div/div[2]/div[2]/div[1]/header/div/div/ul/li[3]/a'
                )))
                driver.execute_script("arguments[0].click();", acc_tab_trim)

                time.sleep(1)

                rows = driver.find_elements(By.XPATH,'//*[@id="trimAccessoriesDt"]/tbody/tr')
                for r in rows:
                    cols = r.find_elements(By.TAG_NAME,"td")
                    if len(cols) < 5:
                        continue

                    part = cols[1].text.strip()

                    if part in accessory_map:
                        accessory_map[part]["Trims"].add(trim_name)
                    else:
                        accessory_map[part] = {
                            "Category": cols[2].text.strip(),
                            "Status": cols[4].text.strip(),
                            "Trims": {trim_name}
                        }

                # -------- STORE TRIM DATA --------
                trim_data.append({
                    "Country": country,
                    "Year": year,
                    "Make": make,
                    "Model": model,
                    "Trim": trim_name,
                    "Trim_Status": trim_status,
                    "Color_Count": len(combined_colors),
                    "Colors": ",".join(combined_colors)
                })

                # -------- BACK --------
                back_btn = wait.until(EC.element_to_be_clickable((
                    By.XPATH,
                    '/html/body/main/div[2]/div[2]/breadcrumbs/div/ul/li[2]/span/a'
                )))
                driver.execute_script("arguments[0].click();", back_btn)

                wait.until(EC.presence_of_element_located((By.XPATH,'//*[@id="trimDt"]/tbody/tr')))

                i += 1

            except Exception as e:
                print("Trim error:", e)
                i += 1

        # -------- PAGINATION --------
        try:
            next_btn = driver.find_element(By.XPATH,'//*[@id="trimDt_next"]')
            if "disabled" not in next_btn.get_attribute("class"):
                old = driver.find_elements(By.XPATH,'//*[@id="trimDt"]/tbody/tr')[0]
                driver.execute_script("arguments[0].click();", next_btn)
                wait.until(EC.staleness_of(old))
            else:
                break
        except:
            break


    # ==================================================
    # BUILD TRIM → ACCESSORY COUNT MAP
    # ==================================================
    trim_count_map = {}

    for part, val in accessory_map.items():

        if part=="NA":
            continue

        for t in val["Trims"]:
            if t not in trim_count_map:
                trim_count_map[t] = set()
            trim_count_map[t].add(part)

    # Convert sets to counts
    for t in trim_count_map:
        trim_count_map[t] = len(trim_count_map[t])

    # ==================================================
    # FINAL ACCESSORY OUTPUT (UPDATED)
    # ==================================================
    final_accessories = []

    for part, val in accessory_map.items():

        trims_list = val["Trims"]
        trim_string = ",".join(sorted(trims_list))

        # ✅ Correct count logic (based on trim group)
        if "ALL" in trims_list:
            acc_count = trim_count_map.get("ALL", 0)
        else:
            acc_count = max([trim_count_map.get(t, 0) for t in trims_list])

        final_accessories.append({
            "Country": country,
            "Year": year,
            "Make": make,
            "Model": model,
            "PartName": part,
            "Category": val["Category"],
            "Status": val["Status"],
            "Trim": trim_string,
            "Accessory_Count": acc_count   # ✅ NEW COLUMN
        })

    # ==================================================
    # BUILD TRIM → ACCESSORY COUNT MAP (FROM ACCESSORIES)
    # ==================================================
    trim_accessory_count_map = {}

    for row in final_accessories:
        trims = row["Trim"].split(",")

        for t in trims:
            t = t.strip()

            if t == "ALL":
                # Apply to all trims
                for tr in trim_data:
                    trim_accessory_count_map[tr["Trim"]] = row["Accessory_Count"]
            else:
                trim_accessory_count_map[t] = row["Accessory_Count"]

    # ==================================================
    # MAIN SHEET CREATION
    # ==================================================
    main_data = []

    for tr in trim_data:
        trim_name = tr["Trim"]

        main_data.append({
            "Country": tr["Country"],
            "Year": tr["Year"],
            "Make": tr["Make"],
            "Model": tr["Model"],
            "Trim": trim_name,
            "Accessories_Count": trim_accessory_count_map.get(trim_name, 0)
        })

    print("MAIN DATA COUNT :", len(main_data))
    print("TRIM DATA COUNT :", len(trim_data))
    print("ACCESSORY COUNT :", len(final_accessories))

    # ==================================================
    # MERGE ALL RESULTS
    # ==================================================

    all_main_data.extend(main_data)
    all_trim_data.extend(trim_data)
    all_final_accessories.extend(final_accessories)



    print("TOTAL MAIN :", len(all_main_data))
    print("TOTAL TRIM :", len(all_trim_data))
    print("TOTAL ACCESSORIES :", len(all_final_accessories))


    # ==================================================
    # CONVERT ALL DATA TO TITLE CASE
    # ==================================================

    for row in all_main_data:

        for key in row:
            row[key] = tc(row[key])

    for row in all_trim_data:

        for key in row:
            row[key] = tc(row[key])

    for row in all_final_accessories:

        for key in row:
            row[key] = tc(row[key])
  
# ==================================================
# FINAL SAVE
# ==================================================

with pd.ExcelWriter("final_output.xlsx") as writer:

    pd.DataFrame(all_main_data).to_excel(
        writer,
        "Main",
        index=False
    )

    pd.DataFrame(all_final_accessories).to_excel(
        writer,
        "Accessories",
        index=False
    )

    pd.DataFrame(all_trim_data).to_excel(
        writer,
        "Colours",
        index=False
    )

print("\n✅ FINAL MERGED OUTPUT READY")

driver.quit()

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
    model_success = False

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

    time.sleep(2)






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
    all_accessory_rows = []
    seen_accessories = set()

   

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

    try:
        wait.until(EC.invisibility_of_element_located((
            By.XPATH,
            '//*[@id="modeldt_processing"]'
        )))
    except:
        print(f"⚠ Loader timeout for {INPUT_MODEL}")

    time.sleep(2)





    # ---------------- COUNTRY ----------------
    country = driver.find_element(By.XPATH, '//*[@id="modeldt"]/tbody/tr[1]/td[2]').text.strip()

    # ---------------- CLICK MODEL ----------------
    for row in driver.find_elements(By.XPATH, '//*[@id="modeldt"]/tbody/tr'):
        try:
            if row.find_element(By.XPATH, "./td[3]").text.lower() == INPUT_MODEL:
                row.find_element(By.XPATH, "./td[3]/a").click()
                time.sleep(3)

                wait.until(EC.presence_of_element_located((
                    By.XPATH,
                    '//*[@id="nav-7-1"]'
                )))
                break
        except:
            continue

    # ---------------- MODEL DETAILS ----------------
    make = wait.until(EC.visibility_of_element_located((By.XPATH,'//*[@id="nav-7-1"]/div/div[2]/div[2]'))).text
    model = wait.until(EC.visibility_of_element_located((By.XPATH,'//*[@id="nav-7-1"]/div/div[3]/div[2]'))).text
    year = wait.until(EC.visibility_of_element_located((By.XPATH,'//*[@id="nav-7-1"]/div/div[5]/div[2]'))).text
    # DEFAULT ALIAS
    alias = "NA"


    # ==================================================
    # MODEL COLORS
    # ==================================================
    color_tab_model = wait.until(EC.element_to_be_clickable((
        By.XPATH,
        '/html/body/main/div[2]/div[2]/div/ui-view/div/div[2]/div[2]/header/div/div/ul/li[2]/a'
    )))
    driver.execute_script("arguments[0].click();", color_tab_model)

    wait.until(EC.presence_of_element_located((
        By.XPATH,
        '//*[@id="modelColorDt"]/tbody'
    )))

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

    wait.until(EC.presence_of_element_located((
        By.XPATH,
        '//*[@id="modelAccessories"]/tbody'
    )))

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

            # accessory_map[part] = {
            #     "Category": cols[2].text.strip(),
            #     "Status": cols[4].text.strip(),
            #     "Trims": {"ALL"}
            # }
            part_number = cols[0].text.strip()
            part = cols[1].text.strip()

            key = (
                part_number,
                part,
                "ALL"
            )

            if key not in seen_accessories:
                seen_accessories.add(key)

                all_accessory_rows.append({
                    "PartNumber": part_number,
                    "PartName": part,
                    "Category": cols[2].text.strip(),
                    "Status": cols[4].text.strip(),
                    "Trim": "ALL"
                })

            accessory_map[part] = {
                "PartNumber": part_number,
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
        all_accessory_rows.append({
        "PartNumber": "0",
        "PartName": "0",
        "Category": "0",
        "Status": "0",
        "Trim": "ALL"
    })

    print("✅ Model accessories done")

    # ==================================================
    # GO TO TRIMS
    # ==================================================
    trim_tab = wait.until(EC.element_to_be_clickable((
        By.XPATH,
        '/html/body/main/div[2]/div[2]/div/ui-view/div/div[2]/div[2]/header/div/div/ul/li[3]/a'
    )))
    driver.execute_script("arguments[0].click();", trim_tab)

    wait.until(EC.presence_of_element_located((
        By.XPATH,
        '//*[@id="trimDt"]/tbody'
    )))

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
                # acc_tab_trim = wait.until(EC.element_to_be_clickable((
                #     By.XPATH,
                #     '/html/body/main/div[2]/div[2]/div/ui-view/div/div[2]/div[2]/div[1]/header/div/div/ul/li[3]/a'
                # )))
                # driver.execute_script("arguments[0].click();", acc_tab_trim)

                # time.sleep(2)

                # rows = driver.find_elements(By.XPATH,'//*[@id="trimAccessoriesDt"]/tbody/tr')
                # print("TRIM ACCESSORY ROWS :", len(rows))
                

                # for r in rows:
                #     cols = r.find_elements(By.TAG_NAME,"td")
                #     if len(cols) < 5:
                #         continue
                #     part_number = cols[0].text.strip()
                #     part = cols[1].text.strip()

                #     key = (
                #         part_number,
                #         part,
                #         trim_name
                #     )

                #     if key not in seen_accessories:

                #         seen_accessories.add(key)

                #         all_accessory_rows.append({
                #             "PartNumber": part_number,
                #             "PartName": part,
                #             "Category": cols[2].text.strip(),
                #             "Status": cols[4].text.strip(),
                #             "Trim": trim_name
                #         })

                #     if part in accessory_map:

                #         accessory_map[part]["Trims"].add(trim_name)

                #     else:

                #         accessory_map[part] = {
                #             "PartNumber": part_number,
                #             "Category": cols[2].text.strip(),
                #             "Status": cols[4].text.strip(),
                #             "Trims": {trim_name}
                #         }



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
                    part_number = cols[0].text.strip()
                    part = cols[1].text.strip()

                    key = (
                        part_number,
                        part,
                        trim_name
                    )

                    if key not in seen_accessories:

                        seen_accessories.add(key)

                        all_accessory_rows.append({
                            "PartNumber": part_number,
                            "PartName": part,
                            "Category": cols[2].text.strip(),
                            "Status": cols[4].text.strip(),
                            "Trim": trim_name
                        })

                    if part in accessory_map:

                        accessory_map[part]["Trims"].add(trim_name)

                    else:

                        accessory_map[part] = {
                            "PartNumber": part_number,
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
    # FINAL ACCESSORY OUTPUT
    # ==================================================

    final_accessories = []

    for row in all_accessory_rows:

        final_accessories.append({
            "Country": country,
            "Year": year,
            "Make": make,
            "Model": model,
            "CG alias/MM alias": alias,
            "PartNumber": row["PartNumber"],
            "PartName": row["PartName"],
            "Category": row["Category"],
            "Status": row["Status"],
            "Trim": row["Trim"]
        
        })


    # ==================================================
    # FETCH ALIAS AFTER ALL STEPS
    # ==================================================

    try:

        # ---------- CLICK TAB ----------
        alias_tab = wait.until(EC.element_to_be_clickable((
            By.XPATH,
            '/html/body/main/div[2]/div[2]/div/ui-view/div/div[2]/div[2]/header/div/div/ul/li[5]/a'
        )))
        driver.execute_script("arguments[0].click();", alias_tab)

        time.sleep(2)

        # ---------- CLICK ACTION ICON ----------
        action_icon = wait.until(EC.element_to_be_clickable((
            By.XPATH,
            '/html/body/main/div[2]/div[2]/div/ui-view/div/div[2]/div[2]/div[1]/div/div/div[2]/div/div/table/tbody/tr[2]/td[2]/div/a/i'
        )))
        driver.execute_script("arguments[0].click();", action_icon)

        time.sleep(2)

        # ---------- CLICK DROPDOWN OPTION ----------
        alias_option = wait.until(EC.element_to_be_clickable((
            By.XPATH,
            '/html/body/main/div[2]/div[2]/div/ui-view/div/div[2]/div[2]/div[1]/div/div/div[2]/div/div/table/tbody/tr[2]/td[2]/div/div/ul/li[1]/a'
        )))
        driver.execute_script("arguments[0].click();", alias_option)

        time.sleep(3)

        # ---------- FETCH ALIAS ----------
        alias_full = wait.until(EC.presence_of_element_located((
            By.XPATH,
            '//*[@id="nav-7"]/div[3]/div/div/div/p/span[2]/a[1]'
        ))).text.strip()

        # example:
        # 26fordescapestlinesu4fa

        alias = alias_full.split("/")[-2]

        print("✅ Alias :", alias)

        # ==================================================
        # UPDATE ALIAS IN ALL DATA
        # ==================================================

        for tr in trim_data:
            tr["CG alias/MM alias"] = alias

        for acc in final_accessories:
            acc["CG alias/MM alias"] = alias

    except Exception as e:

        print("Alias Fetch Error :", e)

        alias = "NA"

    


    # ==================================================
    # MAIN SHEET CREATION
    # ==================================================

    main_data = []

    for tr in trim_data:

        trim_name = tr["Trim"]

        # count accessories for this trim
        acc_count = 0

        for a in final_accessories:

            trim_value = a["Trim"]

            # count ALL accessories also
            # if trim_value == "ALL" or trim_value == trim_name:
            #     acc_count += 1
            if (
                (trim_value == "ALL" or trim_value == trim_name)
                and a["PartNumber"] != "0"
            ):
                acc_count += 1
            

        main_data.append({
            "Country": tr["Country"],
            "Year": tr["Year"],
            "Make": make,
            "Model": model,
            "CG alias/MM alias": alias,
            "Trim": trim_name,
            "Accessories_Count": acc_count
        })





    # ==================================================
    # MERGE ALL RESULTS
    # ==================================================
    if main_data:
        model_success = True
    all_main_data.extend(main_data)
    all_trim_data.extend(trim_data)
    all_final_accessories.extend(final_accessories)



    print("TOTAL MAIN :", len(all_main_data))
    print("TOTAL TRIM :", len(all_trim_data))
    print("TOTAL ACCESSORIES :", len(all_final_accessories))


    if not model_success:

        print(f"⚠ No data captured for {INPUT_MODEL}")

        all_main_data.append({
            "Country": "NA",
            "Year": INPUT_YEAR,
            "Make": "NA",
            "Model": INPUT_MODEL,
            "CG alias/MM alias": "NA",
            "Trim": "NO DATA",
            "Accessories_Count": 0
        })

        all_trim_data.append({
            "Country": "NA",
            "Year": INPUT_YEAR,
            "Make": "NA",
            "Model": INPUT_MODEL,
            "CG alias/MM alias": "NA",
            "Trim": "NO DATA",
            "Trim_Status": "FAILED",
            "Color_Count": 0,
            "Colors": "NA"
        })

        all_final_accessories.append({
            "Country": "NA",
            "Year": INPUT_YEAR,
            "Make": "NA",
            "Model": INPUT_MODEL,
            "CG alias/MM alias": "NA",
            "PartNumber": "NA",
            "PartName": "NO DATA",
            "Category": "NA",
            "Status": "FAILED",
            "Trim": "NO DATA"
        })


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

with pd.ExcelWriter(
    "final_output.xlsx",
    engine="openpyxl"
) as writer:

    # ================= MAIN =================
    

    main_cols = [
        "Country",
        "Year",
        "Make",
        "Model",
        "CG alias/MM alias",
        "Trim",
        "Accessories_Count"
    ]

    main_df = pd.DataFrame(all_main_data)

    main_df = main_df.reindex(
        columns=main_cols,
        fill_value="NA"
    )

    print("MAIN DATA :", len(main_df))

    main_df.to_excel(
        excel_writer=writer,
        sheet_name="Main",
        index=False
    )

    # ================= ACCESSORIES =================
    
    acc_cols = [
        "Country",
        "Year",
        "Make",
        "Model",
        "CG alias/MM alias",
        "PartNumber",
        "PartName",
        "Category",
        "Status",
        "Trim"
    ]

    acc_df = pd.DataFrame(all_final_accessories)


    acc_df = acc_df.reindex(
        columns=acc_cols,
        fill_value="NA"
    )

    print("ACCESSORIES DATA :", len(acc_df))

    acc_df.to_excel(
        excel_writer=writer,
        sheet_name="Accessories",
        index=False
    )

    # ================= COLOURS =================
    
    color_cols = [
        "Country",
        "Year",
        "Make",
        "Model",
        "CG alias/MM alias",
        "Trim",
        "Trim_Status",
        "Color_Count",
        "Colors"
    ]

    color_df = pd.DataFrame(all_trim_data)

    color_df = color_df.reindex(
        columns=color_cols,
        fill_value="NA"
    )

    print("COLOUR DATA :", len(color_df))

    color_df.to_excel(
        excel_writer=writer,
        sheet_name="Colours",
        index=False
    )


print("\n✅ FINAL MERGED OUTPUT READY")

driver.quit()

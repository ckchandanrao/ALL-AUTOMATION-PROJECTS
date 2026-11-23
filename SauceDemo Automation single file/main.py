


from selenium import webdriver
from selenium.webdriver.common.by import By
import time

#disable this via Chrome options(save password and other popups) when creating the driver
from selenium.webdriver.chrome.options import Options

options = Options()
options.add_experimental_option("prefs", {
    "credentials_enable_service": False,
    "profile.password_manager_enabled": False,
    "profile.password_manager_leak_detection": False   # disables the "change your password" breach warning
    
})
options.add_argument("--disable-notifications")
options.add_argument("--disable-popup-blocking")
options.add_argument("--disable-infobars")


#Step 1: Invoking the browser
driver = webdriver.Chrome(options=options)
driver.get("https://www.saucedemo.com")
driver.maximize_window()
print(driver.title)
print(driver.current_url)


#Step 4: Handle Invalid Login credentials (test this FIRST, before valid login)
driver.find_element(By.ID,"user-name").send_keys("abc")
driver.find_element(By.ID,"password").send_keys("def")
driver.find_element(By.ID,"login-button").click()
errors=driver.find_element(By.XPATH,"//h3[@data-test='error']")
print(errors.text)
expected_error="Epic sadface: Username and password do not match any user in this service"
if errors.text==expected_error:
    print("Correct error message shown for Invalid Login")
else:
    print(f"Unexpected error: {errors.text}")

# Instead of clear() + retype, just reload the page fresh
driver.get("https://www.saucedemo.com")

from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

wait = WebDriverWait(driver, 10)
wait.until(EC.visibility_of_element_located((By.CLASS_NAME, "login_logo")))

#Step 2:Login using valid credentials
driver.find_element(By.ID,"user-name").send_keys("standard_user")
driver.find_element(By.ID,"password").send_keys("secret_sauce")
driver.find_element(By.ID,"login-button").click()

# DEBUG: see what's actually on screen right after clicking login
driver.save_screenshot("debug_after_valid_login.png")
print("Current URL right after click:", driver.current_url)

# Check if an error message is STILL showing
try:
    still_error = driver.find_element(By.XPATH, "//h3[@data-test='error']")
    print("ERROR STILL SHOWING:", still_error.text)
except:
    print("No error message present")

#Step 3:Verify login success - WAIT for the inventory page to actually load
wait = WebDriverWait(driver, 10)
wait.until(EC.presence_of_element_located((By.CLASS_NAME, "product_sort_container")))
print(driver.current_url)

#Step 5: Sort products
from selenium.webdriver.support.ui import Select
Select_element=Select(driver.find_element(By.CLASS_NAME,"product_sort_container")) #Selecting the dropdown
for i in Select_element.options: #Here options(a property Select gives) to get all actual <option> elements inside it
    print(i.text)   #Printing all the list of names in the dropdown
Select_element.select_by_index(2)

#Confirm all displayed prices in ascending order
price_div=driver.find_element(By.ID,"inventory_container") #Selecting the entire div
prices=price_div.find_elements(By.CLASS_NAME,"inventory_item_price") #Selecting the class name of the prices
for i in prices: #Looping through all prices
    print(i.text) #printing each price
price_text=[i.text.replace("$","") for i in prices] # remove $ sign
print(price_text)
price_float=[float(p) for p in price_text] # convert to numbers

if price_float == sorted(price_float):
    print("Prices are correctly sorted")
else:
    print("Prices are not sorted")


# Step 6: Add first product
driver.find_element(By.ID,"add-to-cart-sauce-labs-onesie").click()
print("First product added successfully")

wait = WebDriverWait(driver, 5)
cart_badge = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "shopping_cart_badge")))

#checking the add to cart count =1
cart_badge=driver.find_element(By.CLASS_NAME,"shopping_cart_badge")
cart_count=cart_badge.text
print("Cart count is",cart_count)
if cart_count == "1":
    print("One product is added")
else:
    print("Nothing added")

#Step 7:Add a second product, then remove one

driver.find_element(By.ID,"add-to-cart-sauce-labs-bike-light").click() #adding second product
print("second product added successfully")
cart_badge=driver.find_element(By.CLASS_NAME,"shopping_cart_badge")
cart_count=cart_badge.text
print("Cart count is",cart_count)
if cart_count == "2":
    print("second product is added")
else:
    print("Nothing added")

#Removing the one product
driver.find_element(By.ID,"remove-sauce-labs-onesie").click()
print("Removed one product")

wait = WebDriverWait(driver, 5)
cart_badge = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "shopping_cart_badge")))

cart_count = cart_badge.text
if cart_count == "1":
    print("one product is removed")
else:
    print("Nothing removed")


#Step 8: Go to the cart page
driver.find_element(By.CLASS_NAME,"shopping_cart_link").click()

#Step 9:Checkout flow (optional, more advanced)
driver.find_element(By.ID,"checkout").click()
driver.find_element(By.ID,"first-name").send_keys("Chandan")
driver.find_element(By.ID,"last-name").send_keys("Rao")
driver.find_element(By.ID,"postal-code").send_keys("572102")
driver.find_element(By.ID,"continue").click()
driver.find_element(By.ID,"finish").click()
verify=driver.find_element(By.CLASS_NAME,"complete-header")
verify_text=verify.text
if verify_text=="Thank you for your order!":
    print("Thank you for your order!")
else:
    print("Not ordered")

#Step 10: Logout
driver.find_element(By.ID,"react-burger-menu-btn").click()

#Wait for the Logout link specifically to be CLICKABLE, not just present
wait = WebDriverWait(driver, 5)
logout_link = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[text()='Logout']")))

lists=driver.find_elements(By.CLASS_NAME,"bm-item")
for i in lists:
    print(i.text)
lists_of_logout=[i.text for i in lists]
print(lists_of_logout)
driver.find_element(By.XPATH,"//a[text()='Logout']").click()
if driver.current_url == "https://www.saucedemo.com/":
    print("Successfully logged out")
else:
    print(f"Logout may have failed - current URL: {driver.current_url}")

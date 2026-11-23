from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC


class InventoryPage:
    SORT_DROPDOWN = (By.CLASS_NAME, "product_sort_container")
    PRICE_ELEMENTS = (By.CLASS_NAME, "inventory_item_price")
    CART_BADGE = (By.CLASS_NAME, "shopping_cart_badge")
    CART_LINK = (By.CLASS_NAME, "shopping_cart_link")

    ADD_TO_CART_ONESIE = (By.ID, "add-to-cart-sauce-labs-onesie")
    ADD_TO_CART_BIKE_LIGHT = (By.ID, "add-to-cart-sauce-labs-bike-light")
    REMOVE_ONESIE = (By.ID, "remove-sauce-labs-onesie")

    def __init__(self, driver):
        self.driver = driver
        self.wait = WebDriverWait(driver, 10)

    def wait_until_loaded(self):
        self.wait.until(EC.presence_of_element_located(self.SORT_DROPDOWN))

    def sort_by_index(self, index):
        dropdown = Select(self.driver.find_element(*self.SORT_DROPDOWN))
        dropdown.select_by_index(index)

    def get_all_prices(self):
        elements = self.driver.find_elements(*self.PRICE_ELEMENTS)
        prices_text = [e.text.replace("$", "") for e in elements]
        return [float(p) for p in prices_text]

    def add_onesie_to_cart(self):
        self.driver.find_element(*self.ADD_TO_CART_ONESIE).click()
        self.wait.until(EC.presence_of_element_located(self.CART_BADGE))

    def add_bike_light_to_cart(self):
        self.driver.find_element(*self.ADD_TO_CART_BIKE_LIGHT).click()

    def remove_onesie_from_cart(self):
        self.driver.find_element(*self.REMOVE_ONESIE).click()
        self.wait.until(EC.presence_of_element_located(self.CART_BADGE))

    def get_cart_count(self):
        return self.driver.find_element(*self.CART_BADGE).text

    def go_to_cart(self):
        self.driver.find_element(*self.CART_LINK).click()
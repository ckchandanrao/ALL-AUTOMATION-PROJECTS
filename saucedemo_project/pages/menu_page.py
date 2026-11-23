from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


class MenuPage:
    BURGER_MENU_BUTTON = (By.ID, "react-burger-menu-btn")
    LOGOUT_LINK = (By.XPATH, "//a[text()='Logout']")

    def __init__(self, driver):
        self.driver = driver
        self.wait = WebDriverWait(driver, 10)

    def logout(self):
        self.driver.find_element(*self.BURGER_MENU_BUTTON).click()
        logout_link = self.wait.until(EC.element_to_be_clickable(self.LOGOUT_LINK))
        logout_link.click()
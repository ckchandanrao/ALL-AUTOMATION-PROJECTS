import pytest
from selenium import webdriver
from selenium.webdriver.chrome.options import Options


@pytest.fixture
def driver():
    options = Options()
    options.add_experimental_option("prefs", {
        "credentials_enable_service": False,
        "profile.password_manager_enabled": False,
        "profile.password_manager_leak_detection": False
    })
    options.add_argument("--disable-notifications")
    options.add_argument("--disable-popup-blocking")
    options.add_argument("--disable-infobars")
    #options.add_argument("--headless=new")          # runs without a visible window
    options.add_argument("--window-size=1920,1080")   # fixes headless viewport issues
    options.add_argument("--no-sandbox")               # often required in CI environments

    driver = webdriver.Chrome(options=options)
    driver.maximize_window()

    yield driver   # this is what gets handed to each test function

    driver.quit()   # runs AFTER the test finishes, success or failure
from playwright.sync_api import expect
import pytest

@pytest.fixture
def loaded_page(page):
    page.goto("https://opensource-demo.orangehrmlive.com/web/index.php/auth/login")
    return page

def test_filter_locators(loaded_page):
    #loaded_page.locator("input[placeholder='Username']").fill("Admin")
    #loaded_page.locator("//input[@placeholder='Password']").fill("admin123")
    loaded_page.locator(".oxd-text.oxd-text--p.orangehrm-login-forgot-header").click()
    loaded_page.wait_for_timeout(5000)
    loaded_page.go_back()
    loaded_page.wait_for_timeout(3000)
    #product1=loaded_page.locator('.inventory_item').filter(has_text='Sauce Labs Bike Light')
    #product1.locator("button:has-text('Add to cart')").click()
    #loaded_page.wait_for_timeout(5000)
    
    
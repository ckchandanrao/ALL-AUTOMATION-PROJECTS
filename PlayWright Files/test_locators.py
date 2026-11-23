from playwright.sync_api import expect

def test_demo(page):
    page.goto("https://opensource-demo.orangehrmlive.com/web/index.php/auth/login")
    #page.get_by_placeholder("Username").fill("Admin")
    page.locator("input[placeholder='Username']").fill("Admin")
    page.locator("//input[@placeholder='Password']").fill("admin123")
    #page.get_by_placeholder("password").fill("admin123")
    page.get_by_role("button",name="Login").click()
    # page.get_by_alt_text("profile picture").click()
    # page.get_by_text("Logout").click()
    # page.wait_for_timeout(5000)

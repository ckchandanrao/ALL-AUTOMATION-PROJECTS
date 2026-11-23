from playwright.sync_api import sync_playwright

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False)
    context = browser.new_context()
    page = context.new_page()
    page.goto("https://www.google.com")
    print(page.title())
    browser.close()


from playwright.sync_api import sync_playwright

with sync_playwright() as p:
    browser=p.chromium.launch(headless=False)
    context=browser.new_context()
    page=context.new_page()
    page.goto("https://example.com")
    page.get_by_placeholder("Enter username").fill("admin")
    page.get_by_placeholder("Enter password").fill("12345")
    page.get_by_role("button", name="Login").click()

from playwright.sync_api import expect
def test_login(page):
    page.goto("https://example.com")
    page.get_by_placeholder("Enter username").fill("admin")
    page.get_by_placeholder("Enter password").fill("12345")
    page.get_by_role("button", name="Login").click()
    expect(page.get_by_role("button", name="Login")).to_be_visible()


def test_dashboard(logged_in_page):
    logged_in_page.goto("https://example.com")
    logged_in_page.get_by_placeholder("Enter username").fill("admin")
    logged_in_page.get_by_placeholder("Enter password").fill("12345")
    logged_in_page.get_by_role("button", name="Login").click()
    expect(logged_in_page.get_by_role("button", name="Login")).to_be_visible()

@pytest.fixture
def logged_in_page(page):
    page.goto("https://example.com")
    page.get_by_placeholder("Enter username").fill("admin")
    page.get_by_placeholder("Enter password").fill("12345")
    page.get_by_role("button", name="Login").click()

    return page


from playwright.sync_api import expect
def test_dashboard(logged_in_page):
    expect(logged_in_page).to_be_visible()


@pytest.fixture
def logged_in_page(page):
    page.get_by_placeholder("Enter username").fill("admin")
    page.get_by_placeholder("Enter password").fill("12345")
    page.get_by_role("button", name="Login").click()
    return page

from playwright.sync_api import expect
def test_dashboard(logged_in_page):
    expect(logged_in_page).to_be_visible


def test_profile(logged_in_page):
    expect(logged_in_page).to_be_visible
def test_logout(logged_in_page):
    expect(logged_in_page).to_be_visible

@pytest.fixture
def login(page):
    page.goto("https://www.saucedemo.com/")
    return page


def test_login(page):
    page.get_by_placeholder("Username").fill("standard_user")
    page.get_by_placeholder("Password").fill("secret_sauce")
    page.locator("#login-button").click()

    yield

    print("Test Completed")



import sys
import os
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from pages.login_page import LoginPage
from pages.inventory_page import InventoryPage
from pages.cart_page import CartPage
from pages.checkout_page import CheckoutPage
from pages.menu_page import MenuPage


def test_invalid_login(driver):
    login_page = LoginPage(driver)
    login_page.load()
    login_page.login("abc", "def")

    expected_error = "Epic sadface: Username and password do not match any user in this service"
    assert login_page.get_error_message() == expected_error


def test_full_shopping_flow(driver):
    # Login
    login_page = LoginPage(driver)
    login_page.load()
    login_page.login("standard_user", "secret_sauce")

    # Inventory - sort and verify prices
    inventory_page = InventoryPage(driver)
    inventory_page.wait_until_loaded()
    inventory_page.sort_by_index(2)

    prices = inventory_page.get_all_prices()
    assert prices == sorted(prices), "Prices are not sorted correctly"

    # Add first product
    inventory_page.add_onesie_to_cart()
    assert inventory_page.get_cart_count() == "1"

    # Add second product
    inventory_page.add_bike_light_to_cart()
    assert inventory_page.get_cart_count() == "2"

    # Remove one product
    inventory_page.remove_onesie_from_cart()
    assert inventory_page.get_cart_count() == "1"

    # Go to cart, checkout
    inventory_page.go_to_cart()
    cart_page = CartPage(driver)
    cart_page.click_checkout()

    checkout_page = CheckoutPage(driver)
    checkout_page.fill_checkout_info("Chandan", "Rao", "572102")
    checkout_page.finish_order()

    assert checkout_page.get_confirmation_message() == "Thank you for your order!"

    # Logout
    menu_page = MenuPage(driver)
    menu_page.logout()
    assert driver.current_url == "https://www.saucedemo.com/"
import re
import os
from playwright.sync_api import Playwright, sync_playwright, expect
from dotenv import load_dotenv
import time


load_dotenv()

def run(playwright: Playwright) -> None:
   
    username = os.getenv("SCRL_USERNAME")
    password = os.getenv("SCRL_PASSWORD")


    

    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page = context.new_page()
    page.goto("https://shrevecrumpandlow.com/account?vgfa_redirect_to_login=1")
    page.get_by_text("Username or email address *").click()
    page.get_by_label("Username or email address *").fill(username)
    page.get_by_label("Password *Required").click()
    page.get_by_label("Password *Required").fill(password)
    page.locator("p").filter(has_text="Password *Required").locator("span").nth(3).click()
    page.get_by_role("button", name="Log in").click()
    page.get_by_role("link", name="Admin Portal").click()
    page.get_by_role("link", name="Products", exact=True).click()
    page.locator("#content iframe").content_frame.get_by_role("link", name="Add new product").click()
    time.sleep(5)
 
    context.close()
    browser.close()


with sync_playwright() as playwright:
    run(playwright)
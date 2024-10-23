import os
import random
from datetime import datetime

import xlwings as xw
from faker import Faker
from playwright.sync_api import Page, expect, sync_playwright
from playwright.sync_api._generated import Browser, BrowserContext

from selector import home_page, signup_page


def test_signup(page: Page):
    expect(page.locator(signup_page.ENTER_ACCOUNT_INFO_TEXT)).to_contain_text("Enter Account Information")

    faker = Faker()

    # form account information
    # title
    page.get_by_label(random.choice(["Mr.", "Mrs."])).check()

    # password
    page.fill(signup_page.PASSWORD_INPUT_TEXT, "12345")

    # date of birth
    days: list[str] = [str(i) for i in range(1, 32)]
    page.select_option(signup_page.DATE_OF_BIRTH_DAY_DROPDOWN, random.choice(days))

    months: list[str] = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
    page.select_option(signup_page.DATE_OF_BIRTH_MONTH_DROPDOWN, random.choice(months))

    years: list[str] = [str(i) for i in range(1900, 2022)]
    page.select_option(signup_page.DATE_OF_BIRTH_YEAR_DROPDOWN, random.choice(years))

    # firstname
    page.fill(signup_page.FIRSTNAME_INPUT_TEXT, faker.first_name())

    # lastname
    page.fill(signup_page.LASTNAME_INPUT_TEXT, faker.last_name())

    # company
    page.fill(signup_page.COMPANY_INPUT_TEXT, faker.company())

    # address
    page.fill(signup_page.ADDRESS1_INPUT_TEXT, faker.address())
    page.fill(signup_page.ADDRESS2_INPUT_TEXT, faker.address())

    # country
    countries = ["India", "United States", "Canada", "Australia", "Israel", "New Zealand", "Singapore"]
    page.select_option(signup_page.COUNTRY_DROPDOWN, random.choice(countries))

    # state
    page.fill(signup_page.STATE_INPUT_TEXT, faker.state())

    # city
    page.fill(signup_page.CITY_INPUT_TEXT, faker.city())

    # zipcode
    page.fill(signup_page.ZIPCODE_INPUT_TEXT, faker.zipcode())

    # mobile number
    page.fill(signup_page.MOBILE_NUMBER_INPUT_TEXT, faker.phone_number())

    os.makedirs(name="../data/capture", exist_ok=True)
    datetimenow: datetime = datetime.now()
    page.screenshot(path="../data/capture/user_register_" + datetimenow.strftime("%Y-%m-%d_%H-%M-%S") + ".jpg", full_page=True)

    # signup button
    page.click(signup_page.CREATE_ACCOUNT_BUTTON)

    # continue button
    page.click(signup_page.CONTINUE_BUTTON)


with sync_playwright() as p:
    # setup browser and page
    browser: Browser = p.chromium.launch(headless=False, args=["--start-maximized"])
    context: BrowserContext = browser.new_context(no_viewport=True)
    page: Page = context.new_page()
    page.goto("https://automationexercise.com/")

    # homepage
    page.click(home_page.LOGIN_SIGNUP_LINK)
    book: xw.Book = xw.Book("../data/form.xlsx")
    sheet: xw.Sheet = book.sheets("signup")
    name: str = sheet["A2"].value
    email: str = sheet["B2"].value
    page.fill(home_page.NAME_INPUT_TEXT, name)
    page.fill(home_page.SIGNUP_EMAIL_ADDRESS_INPUT_TEXT, email)
    page.click(home_page.SIGNUP_BUTTON)

    test_signup(page)

    # close all resources
    page.pause()
    context.close()
    browser.close()

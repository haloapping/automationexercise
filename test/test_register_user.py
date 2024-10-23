import os
import random
from datetime import datetime

import xlwings as xw
from faker import Faker
from playwright.sync_api import Page, sync_playwright
from playwright.sync_api._generated import Browser, BrowserContext

from selector import home_page, signup_page


def signup(page: Page):
    faker = Faker("id_ID")
    title: str = random.choice(["Mr.", "Mrs."])
    password: str = "12345"
    day: str = random.choice([str(i) for i in range(1, 32)])
    month: str = random.choice(["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"])
    year: str = random.choice([str(i) for i in range(1900, 2022)])
    first_name: str = faker.first_name()
    last_name: str = faker.last_name()
    company: str = faker.company()
    address1: str = faker.address()
    address2: str = faker.address()
    country: str = random.choice(["India", "United States", "Canada", "Australia", "Israel", "New Zealand", "Singapore"])
    state: str = faker.state()
    city: str = faker.city()
    zipcode: str = faker.postcode()
    phone_number: str = faker.phone_number()

    # form account information
    # title
    page.get_by_label(title).check()

    # password
    page.fill(signup_page.PASSWORD_INPUT_TEXT, password)

    # date of birth
    page.select_option(signup_page.DATE_OF_BIRTH_DAY_DROPDOWN, day)

    page.select_option(signup_page.DATE_OF_BIRTH_MONTH_DROPDOWN, month)

    page.select_option(signup_page.DATE_OF_BIRTH_YEAR_DROPDOWN, year)

    # firstname
    page.fill(signup_page.FIRSTNAME_INPUT_TEXT, first_name)

    # lastname
    page.fill(signup_page.LASTNAME_INPUT_TEXT, last_name)

    # company
    page.fill(signup_page.COMPANY_INPUT_TEXT, company)

    # address
    page.fill(signup_page.ADDRESS1_INPUT_TEXT, address1)
    page.fill(signup_page.ADDRESS2_INPUT_TEXT, address2)

    # country
    page.select_option(signup_page.COUNTRY_DROPDOWN, country)

    # state
    page.fill(signup_page.STATE_INPUT_TEXT, state)

    # city
    page.fill(signup_page.CITY_INPUT_TEXT, city)

    # zipcode
    page.fill(signup_page.ZIPCODE_INPUT_TEXT, zipcode)

    # mobile number
    page.fill(signup_page.MOBILE_NUMBER_INPUT_TEXT, phone_number)

    os.makedirs(name="../data/capture", exist_ok=True)
    datetimenow: datetime = datetime.now()
    page.screenshot(path="../data/capture/user_register_" + datetimenow.strftime("%Y-%m-%d_%H-%M-%S") + ".jpg", full_page=True)

    # signup button
    page.click(signup_page.CREATE_ACCOUNT_BUTTON)

    # continue button
    page.click(signup_page.CONTINUE_BUTTON)


with sync_playwright() as p:
    # setup browser and page
    browser: Browser = p.chromium.launch(headless=False, args=["--start-maximized"], channel="chrome")
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

    signup(page)

    # close all resources
    page.close()
    context.close()
    browser.close()

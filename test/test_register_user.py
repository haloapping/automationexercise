import os
import random
from datetime import datetime

import xlwings as xw
from faker import Faker
from playwright.sync_api import Page, sync_playwright
from playwright.sync_api._generated import Browser, BrowserContext

from selector import home_page, signup_page


def create_fake_account() -> tuple[int, dict[str, str]]:
    faker = Faker("id_ID")
    name: str = faker.user_name()
    email: str = faker.free_email()
    password: str = faker.password(length=10)
    title: str = random.choice(["Mr.", "Mrs."])
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

    # save all form input account information
    # to excel file
    book = xw.Book("../data/form.xlsx")
    sheet = book.sheets["signup"]
    next_row: int = len(sheet.range("A1").expand().value) + 1
    sheet[f"A{next_row}"].value = name
    sheet[f"B{next_row}"].value = email
    sheet[f"C{next_row}"].value = title
    sheet[f"D{next_row}"].value = password
    sheet[f"E{next_row}"].value = day
    sheet[f"F{next_row}"].value = month
    sheet[f"G{next_row}"].value = year
    sheet[f"H{next_row}"].value = first_name
    sheet[f"I{next_row}"].value = last_name
    sheet[f"J{next_row}"].value = company
    sheet[f"K{next_row}"].value = address1
    sheet[f"L{next_row}"].value = address2
    sheet[f"M{next_row}"].value = country
    sheet[f"N{next_row}"].value = state
    sheet[f"O{next_row}"].value = city
    sheet[f"P{next_row}"].value = zipcode
    sheet[f"Q{next_row}"].value = phone_number
    sheet.range(f"A{next_row}:Q{next_row}").api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
    sheet.range(f"A{next_row}:Q{next_row}").api.VerticalAlignment = xw.constants.VAlign.xlVAlignCenter
    book.save()

    return next_row, {
        "name": name,
        "email": email,
        "password": password,
        "title": title,
        "day": day,
        "month": month,
        "year": year,
        "first_name": first_name,
        "last_name": last_name,
        "company": company,
        "address1": address1,
        "address2": address2,
        "country": country,
        "state": state,
        "city": city,
        "zipcode": zipcode,
        "phone_number": phone_number,
    }


def signup(page: Page, account: dict[str, str]):
    # form account information
    # title
    page.get_by_label(account["title"]).check()

    # password
    page.fill(signup_page.PASSWORD_INPUT_TEXT, account["password"])

    # date of birth
    page.select_option(signup_page.DATE_OF_BIRTH_DAY_DROPDOWN, account["day"])
    page.select_option(signup_page.DATE_OF_BIRTH_MONTH_DROPDOWN, account["month"])
    page.select_option(signup_page.DATE_OF_BIRTH_YEAR_DROPDOWN, account["year"])

    # firstname
    page.fill(signup_page.FIRSTNAME_INPUT_TEXT, account["first_name"])

    # lastname
    page.fill(signup_page.LASTNAME_INPUT_TEXT, account["last_name"])

    # company
    page.fill(signup_page.COMPANY_INPUT_TEXT, account["company"])

    # address
    page.fill(signup_page.ADDRESS1_INPUT_TEXT, account["address1"])
    page.fill(signup_page.ADDRESS2_INPUT_TEXT, account["address2"])

    # country
    page.select_option(signup_page.COUNTRY_DROPDOWN, account["country"])

    # state
    page.fill(signup_page.STATE_INPUT_TEXT, account["state"])

    # city
    page.fill(signup_page.CITY_INPUT_TEXT, account["city"])

    # zipcode
    page.fill(signup_page.ZIPCODE_INPUT_TEXT, account["zipcode"])

    # mobile number
    page.fill(signup_page.MOBILE_NUMBER_INPUT_TEXT, account["phone_number"])

    os.makedirs(name="../capture/image", exist_ok=True)
    datetimenow: datetime = datetime.now()
    page.screenshot(path="../capture/image/user_register_" + datetimenow.strftime("%Y-%m-%d_%H-%M-%S") + ".jpg", full_page=True)

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

    # create fake account
    next_row, account = create_fake_account()
    print(next_row)

    # homepage
    page.click(home_page.LOGIN_SIGNUP_LINK)
    book: xw.Book = xw.Book("../data/form.xlsx")
    sheet: xw.Sheet = book.sheets("signup")
    name: str = sheet[f"A{next_row}"].value
    email: str = sheet[f"B{next_row}"].value
    page.fill(home_page.NAME_INPUT_TEXT, account["name"])
    page.fill(home_page.SIGNUP_EMAIL_ADDRESS_INPUT_TEXT, account["email"])
    page.click(home_page.SIGNUP_BUTTON)

    signup(page, account)

    # close all process
    book.close()
    page.close()
    context.close()
    browser.close()

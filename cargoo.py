import asyncio
from playwright.async_api import async_playwright
import time
import os
import glob
from openpyxl import load_workbook
import xlwings as xw
from datetime import datetime

from support_module import support_function

CARGOO_WEBSITE = "https://app.cargoo.com/login"
EMAIL = os.getenv("CARGOO_EMAIL", "")
PASSWORD = os.getenv("CARGOO_PASSWORD", "")



def add_months(dt, months):
    month = dt.month - 1 + months
    year = dt.year + month // 12
    month = month % 12 + 1
    mdays = [31, 29 if (year % 4 == 0 and (year % 100 != 0 or year % 400 == 0)) else 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    day = min(dt.day, mdays[month - 1])
    return dt.replace(year=year, month=month, day=day)

async def cargoo_async(po_update_path):
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        context = await browser.new_context(accept_downloads=True)
        page = await context.new_page()
        await page.goto(CARGOO_WEBSITE)

        email = EMAIL
        password = PASSWORD  

        # Wait for all input and login
        await page.wait_for_selector('input[name="email"]')
        await page.fill('input[name="email"]', email)
        await page.click('#login-btn')
        print("button clicked!")
        await page.wait_for_selector('input[type="email"]')
        await page.fill('input[type="email"]', email)
        await page.click('#idSIButton9')
        print("email entered!")
        await page.wait_for_selector('input[type="password"]')
        await page.fill('input[type="password"]', password)
        await page.click('#idSIButton9')
        print("password entered!")
        await page.wait_for_selector('#idSIButton9')
        await page.click('#idSIButton9')
        print("stay signed in clicked!")

        # Wait for BI label and click
        await page.wait_for_selector('div[data-qa-id="nav-item-Business-Intelligence"]', timeout=100000)
        await page.click('div[data-qa-id="nav-item-Business-Intelligence"]')

        # Click Reports
        await page.wait_for_selector('a[data-qa-id="nav-item-business-intelligence-reports"]')
        await page.click('a[data-qa-id="nav-item-business-intelligence-reports"]')

        # Click Shipments
        await page.wait_for_selector("//div[contains(@class, 'group') and .//span[text()='Shipments']]", timeout=60000)
        await page.click("//div[contains(@class, 'group') and .//span[text()='Shipments']]", strict=True)
        print("shipments_item clicked")

        await page.wait_for_timeout(5000)

        # Date range
        today = datetime.today()
        start_dt = add_months(today, -4)
        end_dt = add_months(today, 3)
        start_str = start_dt.strftime("%d.%m.%Y")
        end_str = end_dt.strftime("%d.%m.%Y")

        # Click "Custom"
        await page.wait_for_selector("//button[.//div[contains(normalize-space(.),'Custom')]]")
        await page.click("//button[.//div[contains(normalize-space(.),'Custom')]]", strict=True)
        await page.wait_for_timeout(1000)

        # Fill date inputs
        date_inputs = await page.query_selector_all("input[name='date'].mx-input")
        if len(date_inputs) < 2:
            raise Exception("Expected two date inputs but found %d" % len(date_inputs))
        await date_inputs[0].click()
        await date_inputs[0].fill(start_str)
        await date_inputs[0].press("Enter")
        await page.wait_for_timeout(500)
        await date_inputs[1].click()
        await date_inputs[1].fill(end_str)
        await date_inputs[1].press("Enter")
        await page.wait_for_timeout(500)
        print(f"Date range set: {start_str} -> {end_str}")

        # Generate report
        await page.wait_for_selector("//button[.//div[contains(., 'Generate report')]]")
        await page.click("//button[.//div[contains(., 'Generate report')]]", strict=True)
        print("generate report clicked")

        # Export to Excel
        await page.wait_for_selector("#fm-tab-export", state="visible", timeout=20000)
        await page.hover("#fm-tab-export")
        await page.wait_for_timeout(2000)
        await page.wait_for_selector("#fm-tab-export-excel")
        async with page.expect_download() as download_info:
            await page.click("#fm-tab-export-excel")
            print("export to excel clicked")
        download = await download_info.value
        download_folder = os.path.expanduser("~\\Downloads")
        downloaded_file = os.path.join(download_folder, download.suggested_filename)
        await download.save_as(downloaded_file)
        print("Download complete, latest file:", downloaded_file, "copying to the main file...")


        #download_folder = os.path.expanduser("~\\Downloads")
        #downloaded_file = os.path.join(download_folder, "Shipments_16042026_164606.xlsx")
        #print("Using hardcoded downloaded file path for testing:", downloaded_file)
        # Open the source and target workbooks
        support_function.copy_sheets(source=downloaded_file, source_sheet="Shipments", target_path=po_update_path, target_sheet="Shipments")
        
        await browser.close()
        print(f"Workbook saved: {po_update_path}, finished cargoo module.")
        print("-" * 50)

def cargoo(po_update_path):
    print("Starting cargoo module...")
    asyncio.run(cargoo_async(po_update_path))



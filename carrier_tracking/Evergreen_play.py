from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
import time
import pandas as pd
import re
from bs4 import BeautifulSoup
from datetime import datetime

#booking_numbers = ["091500303414","11111","050500346507","091500303414"]
# booking_numbers = ["050500937921","050500350199", "050500824859", "091500303686","070500200175", "111111","070500113647","050500350199","091500266594"]
CARRIER = "Evergreen"

def Evergreen(booking_numbers):
    result_list = []
    search_cache = {}

    with sync_playwright() as pw:
        # Launch browser (headful so you can pass manual checks if required)
        browser = pw.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()

        # Open the target website
        page.goto("https://ct.shipmentlink.com/servlet/TDB1_CargoTracking.do", timeout=60000)

        cookie_banner_dismissed = False
        second_cookie_banner_dismissed = False
        first_time = True

        for bookingNo in booking_numbers:
            try:
                if bookingNo in search_cache:
                    result_list.append(search_cache[bookingNo])
                    continue

                # Wait until the page shows the expected radio or allow manual checks
                for _ in range(60):
                    if page.query_selector("#s_bk"):
                        break
                    print("Security page or cookie detected. Please complete the manual check...")
                    time.sleep(1)
                    try:
                        page.reload(timeout=30000)
                    except:
                        pass

                # Dismiss cookie banner if present
                if not cookie_banner_dismissed:
                    try:
                        page.locator("#btn_cookie_essential_only").click(timeout=5000)
                        cookie_banner_dismissed = True
                        print("cookie banner dismissed")
                    except PlaywrightTimeoutError:
                        cookie_banner_dismissed = True
                    except Exception:
                        cookie_banner_dismissed = True

                # Click radio and input booking number
                if first_time:
                    try:
                        page.click("#s_bk", timeout=5000)
                    except:
                        pass
                    try:
                        page.fill("#NO", bookingNo, timeout=5000)
                    except:
                        page.fill("input[name='NO']", bookingNo)
                else:
                    page.fill("input[name='NO']", bookingNo)

                print(f"booking number: {bookingNo} filled")
                time.sleep(2)


                # Prepare to handle possible alert/dialog quickly (booking not exists)
                dialog_seen = {"val": False}
                def on_dialog(dialog):
                    dialog_seen["val"] = True
                    try:
                        dialog.dismiss()
                    except:
                        try:
                            dialog.accept()
                        except:
                            pass
                # register once before the action that may trigger the dialog
                page.once("dialog", on_dialog)


                # Click submit (use JS click to handle hidden inputs)
                try:
                    #要先把隱藏的按鈕顯示出來再點擊(雖然我還是不知道她為什麼藏起來了)
                    page.eval_on_selector("input[type='button'][value='Submit']", "el => { el.style.display='block'; el.click(); }")
                    page.locator("input[type='button'][value='Submit']").click(force=True, timeout=5000)
                    #page.eval_on_selector("input[type='button'][value='Submit']", "el => el.click()", timeout=5000)
                    print("Submit button clicked")
                except Exception as e:
                    print("error:", e)
                    # Fallback: try a normal click
                    try:
                        page.locator("input[type='button'][value='Submit']").click(timeout=5000)
                        print("Submit button clicked(option 2)")
                    except:
                        print("Submit button click failed")
                        pass
                
                time.sleep(5)

                
                if dialog_seen["val"]:
                    print(f"booking number {bookingNo} triggered dialog -> skipping to next booking")
                    result = {
                        "booking No": bookingNo,
                        "carrier": CARRIER,
                        "Vessel": "No data",
                        "ETD": "No data",
                        "ETA": "No data"
                    }
                    result_list.append(result)
                    search_cache[bookingNo] = result
                    continue                

                # Wait for table presence (new page or updated content)
                try:
                    page.wait_for_selector("table", timeout=5000)
                except PlaywrightTimeoutError:
                    # If dialog was seen then booking doesn't exist
                    if dialog_seen["val"]:
                        print(f"booking number {bookingNo} not exists in Evergreen.")
                        result = {
                            "booking No": bookingNo,
                            "carrier": CARRIER,
                            "Vessel": "No data",
                            "ETD": "No data",
                            "ETA": "No data"
                        }
                        result_list.append(result)
                        search_cache[bookingNo] = result
                        continue
                    else:
                        # No table and no dialog -> treat as no data
                        result = {
                            "booking No": bookingNo,
                            "carrier": CARRIER,
                            "Vessel": "No data",
                            "ETD": "No data",
                            "ETA": "No data"
                        }
                        result_list.append(result)
                        search_cache[bookingNo] = result
                        continue

                first_time = False

                # Dismiss any second cookie banner on new content
                if not second_cookie_banner_dismissed:
                    try:
                        page.locator("#btn_cookie_essential_only").click(timeout=2000)
                        second_cookie_banner_dismissed = True
                    except:
                        second_cookie_banner_dismissed = True

                html = page.content()
                soup = BeautifulSoup(html, "html.parser")

                # Locate the table with the title "Basic Information"
                basic_info_table = None
                for table in soup.find_all("table", class_="ec-table ec-table-sm"):
                    title_cell = table.find("td", class_="#f13tabb2")
                    if title_cell and "Basic Information" in title_cell.text:
                        basic_info_table = table
                        break

                if basic_info_table:
                    # ETD
                    try:
                        port_of_loading_row = basic_info_table.find("th", string="Port of Loading").find_parent("tr")
                        etd_td = port_of_loading_row.find_all("td")[5]
                        etd_raw = etd_td.text.strip()
                        etd_date = datetime.strptime(etd_raw, "%b-%d-%Y")
                        etd_transformed = etd_date.strftime("%Y/%m/%d")
                    except Exception:
                        etd_transformed = "No data"

                    # ETA
                    try:
                        place_of_delivery_row = basic_info_table.find("th", string=re.compile(r"Place of Delivery", re.IGNORECASE)).find_parent("tr")
                        eta_td = place_of_delivery_row.find_all("td")[1]
                        eta_raw = eta_td.text.strip()
                        eta_date = datetime.strptime(eta_raw, "%b-%d-%Y")
                        eta_transformed = eta_date.strftime("%Y/%m/%d")
                    except Exception:
                        eta_transformed = "No data"

                    # Vessel: find the table that contains "Container Activity Information" and click first link
                    vessel = "No data"
                    try:
                        # Use Playwright to locate the specific table and click its first link so a new page/tab opens
                        container_table_locator = page.locator("table.ec-table.ec-table-sm:has-text('Container Activity Information')").first
                        if container_table_locator.count() > 0:
                            # If clicking opens a new page, wait for it
                            with context.expect_page(timeout=10000) as new_page_info:
                                # click the first link inside that table
                                try:
                                    container_table_locator.locator("a").first.click(timeout=5000)
                                except Exception:
                                    # fallback to evaluate JS click
                                    page.eval_on_selector("table.ec-table.ec-table-sm:has-text('Container Activity Information') a", "el => el.click()")
                            new_page = new_page_info.value
                            try:
                                new_page.wait_for_selector("table", timeout=10000)
                                html2 = new_page.content()
                                soup2 = BeautifulSoup(html2, "html.parser")
                                table2 = soup2.find("table", class_="ec-table ec-table-sm")
                                rows = table2.find_all("tr") if table2 else []
                                # last non-empty text from 4th column
                                for row in rows:
                                    cells = row.find_all("td")
                                    if len(cells) >= 4:
                                        text = cells[3].get_text(strip=True)
                                        if text:
                                            vessel = text
                                # close the new page and switch back
                            except Exception:
                                pass
                            try:
                                new_page.close()
                            except:
                                pass
                    except Exception:
                        vessel = "No data"

                    result = {
                        "booking No": bookingNo,
                        "carrier": CARRIER,
                        "Vessel": vessel,
                        "ETD": etd_transformed,
                        "ETA": eta_transformed
                    }
                    search_cache[bookingNo] = result
                    result_list.append(result)

                else:
                    # No basic info table found
                    result = {
                        "booking No": bookingNo,
                        "carrier": CARRIER,
                        "Vessel": "No data",
                        "ETD": "No data",
                        "ETA": "No data"
                    }
                    search_cache[bookingNo] = result
                    result_list.append(result)

            except Exception as e:
                print(f"Error when searching for booking number {bookingNo}: {e}")
                result_list.append({
                    "booking No": bookingNo,
                    "carrier": CARRIER,
                    "Vessel": "Error",
                    "ETD": "Error",
                    "ETA": "Error"
                })
                continue

        try:
            context.close()
            browser.close()
        except:
            pass

    return result_list
                

if __name__ == "__main__":
    booking_numbers = ["050500937921","050500350199", "050500824859", "091500303686","070500200175", "111111","070500113647","050500350199","091500266594"]
    result_list = Evergreen(booking_numbers)
    result = pd.DataFrame(result_list)
    print(result.to_string(index=False))
    #result.to_csv('evergreen_results.csv', index=False)

import asyncio
import pandas as pd
from bs4 import BeautifulSoup
from playwright.async_api import async_playwright

CARRIER = "Wan Hai"

def split_into_chunks(lst, chunk_size):
    for i in range(0, len(lst), chunk_size):
        yield lst[i:i + chunk_size]

async def WanHai(booking_numbers):
    chunked_booking_numbers = list(split_into_chunks(booking_numbers, 10))
    result_list = []
    search_cache = {}

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        context = await browser.new_context()
        page = await context.new_page()
        await page.goto("https://tw.wanhai.com/views/cargo_track_v2/tracking_query.xhtml")

        # Wait for manual security check
        while True:
            try:
                await page.wait_for_selector("#q_ref_no1", timeout=5000)
                print("Search box found, continuing automation.")
                break
            except Exception:
                print("Security page detected. Please complete the manual check...")
                await asyncio.sleep(5)

        for chunk in chunked_booking_numbers:
            try:
                print(f"Processing chunk: {chunk}")
                unique_booking_numbers = list(dict.fromkeys(chunk))

                # Fill in booking numbers
                for idx, bookingNo in enumerate(unique_booking_numbers):
                    box_id = f"#q_ref_no{idx+1}"
                    await page.fill(box_id, bookingNo)

                await asyncio.sleep(2)
                await page.click("#Query")
                print("clicked query")

                # Wait for new popup
                popup = await context.wait_for_event("page")
                await popup.wait_for_load_state("domcontentloaded")
                html = await popup.content()
                soup = BeautifulSoup(html, "html.parser")
                print("NEW page loaded")

                # Find all result rows
                rows = await popup.query_selector_all("table.tbl-list tr[align='center']")
                idx = 0
                print(f"Found {len(rows)} rows")
                for row in rows:
                    idx += 1
                    booking_no = chunk[idx - 1]
                    try:
                        tds = await row.query_selector_all("td")
                        if not tds or len(tds) < 6:
                            print("Insufficient data detected")
                            continue

                        # Click Booking Data link
                        booking_link_element = await tds[5].query_selector("a:text('Booking Data')")
                        print("Attempting to click Booking Data link")
                        if booking_link_element:
                            vessel_popup, _ = await asyncio.gather(
                                context.wait_for_event("page"),
                                booking_link_element.click()
                            )
                            print("Booking Data link clicked")
                            await vessel_popup.wait_for_load_state("networkidle")
                            print("Booking Data popup opened")
                            html = await vessel_popup.content()
                            print("html fetched")
                            soup = BeautifulSoup(html, "html.parser")
                            #print(soup.prettify().encode('utf-8', errors='replace').decode('utf-8'))

                            # Vessel name
                            vessel_tag = soup.find("th", string="船名")
                            vessel_name = vessel_tag.find_next("td").text.strip() if vessel_tag else ""

                            # Voyage
                            voyage_tag = soup.find("strong", string="航次")
                            voyage_td = voyage_tag.find_parent("th").find_next("td") if voyage_tag else None
                            voyage = voyage_td.text.strip() if voyage_td else ""
                            vessel = f"{vessel_name}-{voyage}"

                            # ETD/ETA
                            etd = "No data"
                            eta = "No data"
                            cytime = "No data"
                            trs = soup.find_all("tr")
                            for tr in trs:
                                if "Port of Loading" in tr.text:
                                    cells = tr.find_all("td")
                                    if len(cells) >= 2:
                                        etd = cells[1].text.strip()
                                if "Port of Discharging" in tr.text:
                                    cells = tr.find_all("td")
                                    if len(cells) >= 2:
                                        eta = cells[1].text.strip()
                                if "Place of Delivery" in tr.text:
                                    cells = tr.find_all("td")
                                    if len(cells) >= 2:
                                        cytime = cells[1].text.strip()

                            search_cache[booking_no] = {
                                "booking No": booking_no,
                                "carrier": "Wan Hai",
                                "Vessel": vessel,
                                "ETD": etd,
                                "ETA": eta,
                                "CYtime": cytime
                            }
                            await vessel_popup.close()
                        else:
                            raise Exception("Booking Data link not found")
                    except Exception as e:
                        print(f"Error processing booking number {booking_no}: {e}")
                        search_cache[booking_no] = {
                            "booking No": booking_no,
                            "carrier": "Wan Hai",
                            "Vessel": "No data",
                            "ETD": "No data",
                            "ETA": "No data",
                            "CYtime": "No data"
                        }
                        continue

                for bookingNo in chunk:
                    result_list.append(search_cache[bookingNo])

                await popup.close()

                # Clear search boxes
                for idx, bookingNo in enumerate(unique_booking_numbers):
                    box_id = f"#q_ref_no{idx+1}"
                    await page.fill(box_id, "")
            except Exception as e:
                print(f"Error when processing chunk {chunk}:{e}")

        # Final check to ensure all booking numbers are in result_list
        for booking_no in booking_numbers:
            if not any(result["booking No"] == booking_no for result in result_list):
                result_list.append({
                    "booking No": booking_no,
                    "carrier": "Wan Hai",
                    "Vessel": "Error",
                    "ETD": "Error",
                    "ETA": "Error",
                    "CYtime": "Error"
                })

        await browser.close()
        return result_list

if __name__ == "__main__":
    booking_numbers = ["030GG00198","030GG00135","091GA00680","0","aa","091GA00458"]
    #booking_numbers = ["030GG00198"]
    result_list = asyncio.run(WanHai(booking_numbers))
    result = pd.DataFrame(result_list)
    print(result.to_string(index=False))
import requests
import time
import logging
from typing import List, Dict
import re
import pandas as pd

#This is not recommended for production due to security risks.
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


#https://www.yangming.com/api/CargoTracking/GetTracking?paramTrackNo=I470260415&paramTrackPosition=SEARCH&paramRefNo=

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s:%(message)s')

YANGMING_API_URL = "https://www.yangming.com/api/CargoTracking/GetTracking"
CARRIER = "Yang Ming Line"
RATE_LIMIT = 0.5  # requests per second
RETRY_LIMIT = 3
TIMEOUT = 10  # seconds

def extract_port(port_str):
    """Extract port name before parenthesis."""
    return re.sub(r"\s*\(.*?\)", "", port_str).strip()

def extract_date(datetime_str):
    """Extract YYYY/MM/DD from datetime string."""
    return datetime_str.split()[0] if datetime_str else ""

def get_tracking_info(booking_number: str, session: requests.Session) -> Dict:
    params = {
        "paramTrackNo": booking_number,
        "paramTrackPosition": "SEARCH",
        "paramRefNo": ""
    }
    for attempt in range(1, RETRY_LIMIT + 1):
        try:
            response = session.get(YANGMING_API_URL, params=params, timeout=TIMEOUT, verify=False)
            if response.status_code == 429:
                logging.warning(f"Rate limited for {booking_number}, retrying after delay...")
                time.sleep(2)
                continue
            response.raise_for_status()
            data = response.json()
            if not data.get("blList"):
                logging.warning(f"No B/L info found for {booking_number}")
                return {
                    "booking No": booking_number,
                    "carrier": CARRIER,
                    "Vessel": "No data",
                    "ETD": "No data",
                    "ETA": "No data",
                    "CYtime": "No data"
                }
            bl = data["blList"][0]
            basic = bl.get("basicInfo", {})
            routing = bl.get("routingInfo", {}).get("routingSchedule", [])
            vessel = f"{basic.get('vesselName', '')}-{basic.get('vesselComn', '')}".strip("-")
            loading_port = extract_port(basic.get("loading", ""))
            discharge_port = extract_port(basic.get("discharge", ""))
            delivery_place = extract_port(basic.get("delivery", ""))
            etd, eta, cy_time = "No data", "No data", "No data"
            for stop in routing:
                if stop.get("placeName") == loading_port and etd == "No data":
                    etd = extract_date(stop.get("dateTime", ""))
                if stop.get("placeName") == discharge_port and eta == "No data":
                    eta = extract_date(stop.get("dateTime", ""))
                if stop.get("placeName") == delivery_place and cy_time == "No data":
                    cy_time = extract_date(stop.get("dateTime", ""))
            return {
                "booking No": booking_number,
                "carrier": CARRIER,
                "Vessel": vessel,
                "ETD": etd,
                "ETA": eta,
                "CYtime": cy_time
            }
        except requests.RequestException as e:
            logging.error(f"Error fetching {booking_number} (attempt {attempt}): {e}")
            if attempt < RETRY_LIMIT:
                time.sleep(2 ** attempt)
            else:
                return {
                    "booking No": booking_number,
                    "carrier": CARRIER,
                    "Vessel": "Error",
                    "ETD": "Error",
                    "ETA": "Error",
                    "CYtime": "Error"
                }

def YangMing(booking_numbers: List[str]) -> List[Dict]:
    session = requests.Session()
    results = []
    for idx, booking_number in enumerate(booking_numbers):
        logging.info(f"Tracking booking {booking_number} ({idx+1}/{len(booking_numbers)})")
        info = get_tracking_info(booking_number, session)
        results.append(info)
        if idx < len(booking_numbers) - 1:
            time.sleep(1 / RATE_LIMIT)
    return results


if __name__ == "__main__":
    # Example usage
    # booking_numbers = [
    #     "I470260415","I470271630","I488348602","I488348601","I488348586",
    #     "I488349311","I488351445","I488350067","I470271075","I488347073",
    #     "I488348441","I488347072","I488348367","I488348366","I470271345",
    #     "I470268859","I470270110","I488348343","I470270590","I470271258",
    #     "I470268836","I470270109","I470271481"
    #     # Add more booking numbers here
    # ]
    booking_numbers = [
        "I470260415","I470271630","I488348602"
        # Add more booking numbers here
    ]
    result_list = YangMing(booking_numbers)
    result = pd.DataFrame(result_list)
    print(result.to_string(index=False))
    result.to_csv('yangming_results.csv', index=False)
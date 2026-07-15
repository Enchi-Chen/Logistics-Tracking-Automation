import requests
import time
import logging
import pandas as pd
import urllib3

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s:%(message)s')
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

#Network 中名稱為search!
API_URL = "https://ecomm.one-line.com/api/v1/edh/containers/track-and-trace/search"
CARRIER = "Ocean Network Express"
HEADERS = {
    "Content-Type": "application/json",
    "User-Agent": "Mozilla/5.0"
}
RATE_LIMIT = 1  # requests per second
RETRY_LIMIT = 3
TIMEOUT = 10  # seconds

def extract_date(events, matrix_id):
    for event in events:
        if event.get("matrixId") == matrix_id:
            date_str = event.get("date", "")
            if date_str and len(date_str) >= 10:
                return date_str[:10].replace("-", "/")
    return "No data"



def get_vessel_info(booking_no: str, session: requests.Session, timeout: int = 10) -> str:
    vessel_url = f"https://ecomm.one-line.com/api/v1/edh/vessel/track-and-trace/voyage-list?booking_no={booking_no}"
    try:
        response = session.get(vessel_url, timeout=timeout, verify=False)
        if response.status_code in (200, 201):
            data = response.json().get("data", [])
            if not data:
                return "No data"
            is_actual_list = [item.get("pol", {}).get("isActual", True) for item in data]
            if all(is_actual_list):
                # All are actual, use the last item
                last_item = data[-1]
                vessel = last_item.get("vesselEngName", "No data")
                voyage = last_item.get("outboundConsortiumVoyage", "No data")
                return f"{vessel} {voyage}"
            elif not any(is_actual_list):
                # All are non-actual, use the first item
                first_item = data[0]
                vessel = first_item.get("vesselEngName", "No data")
                voyage = first_item.get("outboundConsortiumVoyage", "No data")
                return f"{vessel} {voyage}"
            else:
                # Use the last actual item in the list
                last_actual_idx = None
                for idx in reversed(range(len(is_actual_list))):
                    if is_actual_list[idx]:
                        last_actual_idx = idx
                        break
                if last_actual_idx is not None:
                    item = data[last_actual_idx]
                    vessel = item.get("vesselEngName", "No data")
                    voyage = item.get("outboundConsortiumVoyage", "No data")
                    return f"{vessel} {voyage}"
                # Fallback: if somehow no actual found, use the first item
                vessel = data[0].get("vesselEngName", "No data")
                voyage = data[0].get("outboundConsortiumVoyage", "No data")
                return f"{vessel} {voyage}"
        else:    
            return "Error"
    except Exception as e:
        logging.error(f"Error getting vessel info for booking {booking_no}: {e}")
        return "Error"

def track_booking(booking_no, session, retries=RETRY_LIMIT, timeout=TIMEOUT):
    payload = {
        "page": 1,
        "page_length": 10,
        "filters": {
            "search_text": booking_no,
            "search_type": "BKG_NO"
        }
    }
    for attempt in range(1, retries + 1):
        try:
            response = session.post(API_URL, json=payload, headers=HEADERS, timeout=timeout, verify=False)
            if response.status_code in (200, 201):
                data = response.json()
                if data.get("status") == 200 and data.get("data"):
                    entry = data["data"][0]
                    cargo_events = entry.get("cargoEvents", [])
                    etd = extract_date(cargo_events, "E061")
                    eta = extract_date(cargo_events, "E089")
                    vessel = get_vessel_info(booking_no, session)
                    return {
                        "booking No": booking_no,
                        "carrier": CARRIER,
                        "Vessel": vessel,
                        "ETD": etd,
                        "ETA": eta,
                        "CYtime": "Not finished yet"
                    }
                else:
                    logging.warning(f"No data found for booking {booking_no}")
                    return {
                        "booking No": booking_no,
                        "carrier": CARRIER,
                        "Vessel": "No data",
                        "ETD": "No data",
                        "ETA": "No data",
                        "CYtime": "Not finished yet"
                    }
            else:
                logging.warning(f"Status code {response.status_code} for booking {booking_no}")
        except Exception as e:
            logging.error(f"Error for booking {booking_no} on attempt {attempt}: {e}")
        time.sleep(1)  # Wait before retry
    # If all retries failed
    return {
        "booking No": booking_no,
        "carrier": CARRIER,
        "Vessel": "Error",
        "ETD": "Error",
        "ETA": "Error",
        "CYtime": "Error"

    }

def ONE(booking_numbers):
    results = []
    with requests.Session() as session:
        for idx, booking_no in enumerate(booking_numbers):
            logging.info(f"Tracking booking {booking_no} ({idx+1}/{len(booking_numbers)})")
            result = track_booking(booking_no, session)
            results.append(result)
            if idx < len(booking_numbers) - 1:
                time.sleep(1 / RATE_LIMIT)  # Rate limit: 2 requests per second
    return results

# Example usage:
if __name__ == "__main__":
    # booking_numbers = ["BIOF00756400","OSAF78208800","BASF06853400","OSAF74211500","OSAF70418300",
    #                    "BCNF15692900","BCNF15694400","BASF05610800","BCNF14236900","HAMF93922600"]
    booking_numbers = ["BIOF00756400","OSAF78208800","HAMF93922600"]
    result_list = ONE(booking_numbers)
    result = pd.DataFrame(result_list)
    print(result.to_string())


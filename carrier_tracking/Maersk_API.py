
from typing import Dict, List

import os
import requests
import time
import logging
from requests.adapters import HTTPAdapter, Retry
from datetime import datetime
import pandas as pd

#This is not recommended for production due to security risks.
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

CLIENT_ID = os.getenv("MAERSK_CLIENT_ID", "")
CLIENT_SECRET = os.getenv("MAERSK_CLIENT_SECRET", "")
CLIENT_ID_2 = os.getenv("MAERSK_CLIENT_ID_2", "")
CLIENT_SECRET_2 = os.getenv("MAERSK_CLIENT_SECRET_2", "")


RETRY_LIMIT = 3
RATE_LIMIT = 1  # requests per second
CARRIER = "Maersk Line"

class MaerskAuth:
    def __init__(self, client_id, client_secret):
        self.client_id = client_id
        self.client_secret = client_secret
        self.token = None
        self.expire_at = 0

    def get_token(self):
        if self.token and time.time() < self.expire_at:
            return self.token

        response = requests.post(
            "https://api.maersk.com/customer-identity/oauth/v2/access_token",
            headers={
                "Cache-Control": "no-cache",
                "Content-Type": "application/x-www-form-urlencoded",
                "Consumer-Key": self.client_id
            },
            data={
                "grant_type": "client_credentials",
                "client_id": self.client_id,
                "client_secret": self.client_secret
            },
            verify=False #待改掉
        )
        response.raise_for_status()

        payload = response.json()
        self.token = payload["access_token"]
        self.expire_at = time.time() + payload["expires_in"] - 60
        return self.token



# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s')

def get_session_with_retries():
    session = requests.Session()
    retries = Retry(
        total=RETRY_LIMIT,
        backoff_factor=1,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["HEAD", "GET", "OPTIONS", "POST"]
    )
    adapter = HTTPAdapter(max_retries=retries)
    session.mount('https://', adapter)
    session.mount('http://', adapter)
    return session



def parse_events(events):
    vessel = "No data"
    etd = "No data"
    eta = "No data"
    CYtime = "No data"
    CY_arrived = False

    # Find last ACT/DEPA event for vessel
    act_depa_events = [e for e in events if e.get("eventClassifierCode") == "ACT" and e.get("transportEventTypeCode") == "DEPA"]
    if act_depa_events:
        #print(act_depa_events[-1])
        last_act_depa = act_depa_events[-1]
        vessel_name = last_act_depa.get("transportCall", {}).get("vessel", {}).get("vesselName", "")
        voyage = last_act_depa.get("transportCall", {}).get("exportVoyageNumber", "")
        vessel = f"{vessel_name} {voyage}".strip()
    # ETD: first ACT/DEPA event, else first EST/DEPA event
    if act_depa_events:
        etd_event = act_depa_events[0]
        etd = etd_event.get("eventDateTime")
    else:
        est_depa_events = [e for e in events if e.get("eventClassifierCode") == "EST" and e.get("transportEventTypeCode") == "DEPA"]
        if est_depa_events:
            etd_event = est_depa_events[0]
            etd = etd_event.get("eventDateTime")
    if etd:
        etd = datetime.fromisoformat(etd[:10]).strftime("%Y/%m/%d")

    # ETA: last ARRI event for Keelung/Kaohsiung
    arri_events = [e for e in events if e.get("transportEventTypeCode") == "ARRI" and e.get("transportCall", {}).get("location", {}).get("locationName") and ("Keelung Sea Terminal" in e["transportCall"]["location"]["locationName"] or "Kaohsiung" in e["transportCall"]["location"]["locationName"])]
    if arri_events:
        eta_event = arri_events[-1]
        eta = eta_event.get("eventDateTime")
        if eta:
            eta = datetime.fromisoformat(eta[:10]).strftime("%Y/%m/%d")

    # CYtime step 1 :Find containers from first SHIPMENT/DRFT event
    containers = []
    #for e in events:
        # print(e)
        # #if e.get("eventType") == "SHIPMENT":
        # if e.get("eventType") == "SHIPMENT" and e.get("shipmentEventTypeCode") == "DRFT":
        #     refs = e.get("references", [])
        #     print(f"Found DRFT event with references: {refs}")
        #     for ref in refs:
        #         if ref.get("referenceType") == "EQ":
        #             containers.append(ref.get("referenceValue"))
        #    break  # Only first DRFT event
    containers = list({
        e.get("equipmentReference")
        for e in events
        if e.get("eventType") == "EQUIPMENT" and e.get("equipmentReference")
    })
    print(f"Found containers: {containers}")


    # CYtime step 2 : Find GTIN and Laden equipment event for "CMT Logistics Co. Ltd"，並且考慮可能有多個貨櫃
    cy_times = []
    GTIN_events = [e for e in events if e.get("equipmentEventTypeCode") == "GTIN" and e.get("emptyIndicatorCode") == "LADEN" and e.get("transportCall", {}).get("location", {}).get("locationName") == "CMT Logistics Co. Ltd"]
    if not GTIN_events: # 還沒到 -> 不會記錄在equipment event裡(equipment event沒有EST只有ACT)，但會有統一的transport event來預估GATE IN，也就是最後一個ARRI (EST)
        arri_events = [e for e in events if e.get("transportEventTypeCode") == "ARRI"]
        if arri_events:
            cy_event = arri_events[-1]
            CYtime_val = cy_event.get("eventDateTime")
            if CYtime_val:
                CYtime_val = datetime.fromisoformat(CYtime_val[:10]).strftime("%Y/%m/%d")
            cy_times.append(f"{CYtime_val} (EST)")
    else: #至少有一個貨櫃已到櫃場，檢查每個貨櫃到的時間，已到的後方寫(ACT)，沒到的仍用ARRI event(EST)
        for container in containers:
            #print(f"Checking container {container} for GTIN events...")
            found = False
            for gtin_event in GTIN_events:
                if gtin_event.get("equipmentReference") == container:
                    CYtime_val = gtin_event.get("eventDateTime")
                    if CYtime_val:
                        CYtime_val = datetime.fromisoformat(CYtime_val[:10]).strftime("%Y/%m/%d")
                        if len(containers   ) > 1:
                            cy_times.append(f"{container}: {CYtime_val} (ACT)")
                        else:                            
                            cy_times.append(f"{CYtime_val} (ACT)")
                        CY_arrived = True
                        found = True
                    break
            if not found:
                # Use last ARRI event as EST
                arri_events = [e for e in events if e.get("transportEventTypeCode") == "ARRI"]
                if arri_events:
                    cy_event = arri_events[-1]
                    CYtime_val = cy_event.get("eventDateTime")
                    if CYtime_val:
                        CYtime_val = datetime.fromisoformat(CYtime_val[:10]).strftime("%Y/%m/%d")
                        cy_times.append(f"{container}: {CYtime_val} (EST)")
                else:
                    cy_times.append(f"{container}: No data")
            CY_arrived = True
    print(f"CY_arrived: {CY_arrived}")
    if not cy_times:
        CYtime = "No data"
    else:
        CYtime = ", ".join(cy_times) if cy_times else "No data"

    return vessel, etd, eta, CYtime

def track_maersk_bookings(booking_numbers, token, client_id):
    session = get_session_with_retries()
    url = "https://api.maersk.com/track-and-trace-private/events"
    results = []
    for idx, booking_no in enumerate(booking_numbers):
        params = {
            "carrierBookingReference": booking_no,
            "eventType": ["SHIPMENT","TRANSPORT", "EQUIPMENT"],
            "limit": 500
        }
        headers = {
            "Authorization": f"Bearer {token}",
            "Consumer-Key": client_id,
            "API-Version": "1"
        }
        try:
            logging.info(f"Tracking booking {booking_no} ({idx+1}/{len(booking_numbers)})")
            resp = session.get(url, params=params, headers=headers, timeout=15, verify=False)
            resp.raise_for_status()
            data = resp.json()
            events = data.get("events", [])
            vessel, etd, eta, CYtime = parse_events(events)
            results.append({
                "booking No": booking_no,
                "carrier": CARRIER,
                "Vessel": vessel,
                "ETD": etd,
                "ETA": eta,
                "CYtime": CYtime
            })
            #logging.info(f"Success for {booking_no}: {results[booking_no]}")
        except Exception as e:
            logging.error(f"Error for {booking_no}: {e}")
            results.append({
                "booking No": booking_no,
                "carrier": CARRIER,
                "Vessel": "Error",
                "ETD": "Error",
                "ETA": "Error",
                "CYtime": "Error"
            })
        time.sleep(1 / RATE_LIMIT)
    return results



def Maersk(booking_numbers: List[str]) -> List[Dict]:
    auth = MaerskAuth(CLIENT_ID, CLIENT_SECRET)
    auth2 = MaerskAuth(CLIENT_ID_2, CLIENT_SECRET_2)
    token = auth.get_token()
    token2 = auth2.get_token()
    result_list = track_maersk_bookings(booking_numbers, token2, CLIENT_ID_2)

    # Find bookings with "No data" in ETD, try with second token
    no_data_bookings = [r["booking No"] for r in result_list if (r["ETD"] == "No data" or r["ETD"] == "Error")]
    if no_data_bookings:
        logging.info(f"Retrying {len(no_data_bookings)} bookings with another CLIENT_ID")
        retry_results = track_maersk_bookings(no_data_bookings, token, CLIENT_ID)
        # Update result_list with retry_results where ETD is not "No data"
        for retry in retry_results:
            if retry["ETD"] != "No data" and retry["ETD"] != "Error":
                for orig in result_list:
                    if orig["booking No"] == retry["booking No"]:
                        orig.update(retry)

    return result_list

# Example usage
if __name__ == "__main__":
    # booking_numbers = [
    # "263764413","263525573","263342255","263342256","262599863",
    # "262295323","262252444","262251710","261762945","259762947",
    # "259761479","265012417","264110216"]
    booking_numbers = ["264110216"]
    result_list = Maersk(booking_numbers)
    result = pd.DataFrame(result_list)
    print(result.to_string())


'''
262252444
261762945
262651386
262251710
259762947
261230426
259761479
262705603
259756421
261807299
261807520
260620901
259920047
260117044
'''


'''
263764413
263525573
263342255
263342256
262599863
262295323
262252444
262251710
261762945
259762947
259761479
265012417
264110216
'''
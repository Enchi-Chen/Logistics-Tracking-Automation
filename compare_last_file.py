import os
import pandas as pd
from pyparsing import col
#from openpyxl import load_workbook
from support_module import support_function
import xlwings as xw
from datetime import time

#此函式會: 1.重新寫入"ETA History"以紀錄每筆資料的ETA變化歷史；2.將舊資料中已經沒有的紀錄移到"歷史"工作表，並加上日期標記

def compare_last_file(po_update_path):
    print("Starting compare_last_file module...")

    #recalculate workbook to ensure we have the latest data before comparison
    support_function.recalculate_workbook(po_update_path)
    print("Workbook recalculated before comparison.")
    
    # Find most recent old file (excluding new file)
    folder = os.path.dirname(po_update_path)
    files = [f for f in os.listdir(folder) if f.endswith('.xlsx') and f != os.path.basename(po_update_path)]
    old_file = max([os.path.join(folder, f) for f in files], key=os.path.getmtime)
    old_po_path = old_file
    print("Old PO path:", old_po_path)

    # Read "大表" with header in 2nd row (index 1)
    new_df = pd.read_excel(po_update_path, sheet_name="大表", header=1)
    old_df = pd.read_excel(old_po_path, sheet_name="大表", header=1)

    # Helper: Convert date to MM/DD
    def convert_mmdd(val):
        try:
            dt = pd.to_datetime(val)
            return dt.strftime("%m/%d")
        except Exception:
            return ""

    # Generate ETA History
    eta_history = []
    for idx, row in new_df.iterrows():
        item_code = row['Material']
        po_doc = row['Purchasing Document']
        new_eta = row['Cargoo\nETA/ATA']
        match = old_df[(old_df['Material'] == item_code) & (old_df['Purchasing Document'] == po_doc)]
        if match.empty: #整筆為新資料，直接新增ETA History為當週cargoo ETA
            eta_history.append(convert_mmdd(new_eta))
        else:
            old_row = match.iloc[0]
            old_eta = old_row['Cargoo\nETA/ATA']
            old_history = old_row.get('ETA History', "")
            if pd.isna(old_history):
                old_history = ""
            if str(new_eta) == str(old_eta): #新舊cargoo ETA相同，直接沿用舊的ETA history
                eta_history.append(old_history)
            else: #新舊cargoo ETA不同
                new_mmdd = convert_mmdd(new_eta)
                if old_history == "": #舊ETA history無歷史紀錄，直接新增新的ETA history
                    eta_history.append(new_mmdd)
                else: #舊ETA history有歷史紀錄，新的cargoo ETA接在後方
                    eta_history.append(f"{old_history} -> {new_mmdd}")


    new_df['ETA History'] = eta_history

    #drop rows where the 3rd column (material) is empty in new_df to avoid mistakes
    third_col_name = new_df.columns[2]
    empty_mask = new_df[third_col_name].isnull() | (new_df[third_col_name].astype(str).str.strip() == "")
    if empty_mask.any():
        first_empty_idx = empty_mask.idxmax()
        new_df = new_df.loc[:first_empty_idx-1]

    print("ETA History generated.")
    print(f"Length of new_df: {len(new_df)}")



    with xw.App(visible=False) as app:
        wb = xw.Book(po_update_path)
        ws = wb.sheets["大表"]
        print("Workbook opened with xlwings for writing ETA History.")

        # Find the header row and the "ETA History" column index
        header_row = 2  # 1-based in Excel
        headers = ws.range((header_row, 1)).expand('right').value
        #print(f"Headers found: {headers}")
        eta_history_col = headers.index("ETA History") + 1  # 1-based
        print(f"ETA History column found at index: {eta_history_col}")

        # Write ETA History values (data starts from row 3)
        # for i, value in enumerate(eta_history, start=3):
        #     ws.range((i, eta_history_col)).value = value
        #     print(f"Written ETA History for row {i}: {value}")

        # Prepare the values as a column (list of lists)
        eta_history_column = [[v] for v in eta_history]  # Each value in its own list for a column
        ws.range((3, eta_history_col), (2 + len(eta_history), eta_history_col)).value = eta_history_column
        print("ETA History updated in 大表.")



        # Move outdated records to "歷史"
        history_rows = []
        for idx, row in old_df.iterrows():
            item_code = row['Material']
            po_doc = row['Purchasing Document']
            if pd.isna(item_code) or str(item_code).strip() == "": # Skip if Material is empty or NaN
                continue
            
            match = new_df[(new_df['Material'] == item_code) & (new_df['Purchasing Document'] == po_doc)]
            if match.empty:
                history_rows.append(row)

        if history_rows:
            history_df = pd.DataFrame(history_rows)
            # Add today's date column
            from datetime import datetime
            today_str = datetime.today().strftime('%Y-%m-%d')
            history_df['Date'] = today_str
            

            #drop rows where the 3rd column (material) is empty in history_df to avoid mistakes
            third_col_name = history_df.columns[2]
            empty_mask = history_df[third_col_name].isnull() | (history_df[third_col_name].astype(str).str.strip() == "")
            if empty_mask.any():
                first_empty_idx = empty_mask.idxmax()
                history_df = history_df.loc[:first_empty_idx-1]
            print(f"Length of history_df (outdated records): {len(history_df)}")

            # Convert any datetime.time columns to string to avoid Excel issues
            for col in history_df.columns:
                if history_df[col].apply(lambda x: isinstance(x, time)).any():
                    history_df[col] = history_df[col].astype(str)

            try:
                if "歷史" in [s.name for s in wb.sheets]:
                    ws_history = wb.sheets["歷史"]
                    start_row = ws_history.range("C2").end("down").row + 1
                    print(ws_history.range("C2:C20").value)
                    print(f"Existing 歷史 sheet found. Appending data starting at row {start_row}.")
                else:
                    ws_history = wb.sheets.add("歷史")
                    start_row = 1
                    # Write headers if needed
                    ws_history.range((start_row, 1)).value = history_df.columns.tolist()
                    start_row += 1
            except Exception as e:
                print("Error updating history:", e)

        # Write data
        if history_rows:
            ws_history.range((start_row, 1)).value = history_df.values.tolist()
        print(f"{len(history_rows)} outdated records moved to 歷史.")

        wb.save()
        wb.close()

    print("Finished compare_last_file module.")
    print("-" * 50)


if __name__ == "__main__":
    # File paths
    po_update_path = r"YOUR_NETWORK_DRIVE\ShipmentTracking\PO update\temp\PO update 20260225_experiment_new_sample.xlsx"
    compare_last_file(po_update_path)
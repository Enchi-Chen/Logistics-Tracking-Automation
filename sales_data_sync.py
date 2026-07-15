import os
import shutil
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime
from support_module import support_function
import xlwings as xw


#----------由於SalesData為第一個更新模組，因此可以設置要清除內容的區域-----------#
RANGE1 = "A3:AB"
RANGE2 = "AO3:AZ"
SALES_DATA_FOLDER = r"YOUR_NETWORK_DRIVE\ShipmentTracking\SalesData\PO LIST" 

def sales_data_sync(po_update_path):
    print("-" * 50)
    print("Starting sales_data_sync module...")

    # Step 2: Erase content from 3rd row, columns indicated by RANGE1 and RANGE2, until the last row with data in column C
    with xw.App(visible=False) as app:
        wb = app.books.open(po_update_path)
        ws = wb.sheets["大表"]
        last_row = ws.range('C' + str(ws.cells.last_cell.row)).end('up').row
        print(f"last_row: {last_row}")
        ws.range(f"{RANGE1}{last_row}").clear_contents()
        ws.range(f"{RANGE2}{last_row}").clear_contents()

        # Step 3: Open latest_data from SalesData\PO LIST
        latest_data_folder = SALES_DATA_FOLDER
        latest_data_file = support_function.get_latest_file(latest_data_folder)
        df = pd.read_excel(latest_data_file, header=None)

        # Step 4: Copy content from A2 and paste to A3 of PO update, only rows with [Still to be delivered (qty)] > 0
        header_row = df.iloc[0]
        try:
            qty_col_idx = header_row.tolist().index('Still to be delivered (qty)')
        except ValueError:
            raise Exception("Column 'Still to be delivered (qty)' not found in source file.")

        filtered_rows = [row.tolist() for row in df.values[1:] if row[qty_col_idx] > 0]
        if filtered_rows:
            ws.range(f"A3").value = filtered_rows

        # Step 5: Save the file
        wb.save()
    print(f"PO update saved: {po_update_path}, finished sales_data_sync module.")
    print("-" * 50)



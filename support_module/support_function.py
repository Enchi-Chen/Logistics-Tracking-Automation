import os
import shutil
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime
import xlwings as xw



def get_latest_file(folder, ext=".xlsx", contains_text=None):
    files = [f for f in os.listdir(folder) if f.endswith(ext)]
    if contains_text:
        files = [f for f in files if contains_text in f]
    files = [os.path.join(folder, f) for f in files]
    if not files:
        return None  # or raise an exception if preferred
    latest_file = max(files, key=os.path.getmtime)
    return latest_file


def create_new_file(src_folder):
    today_str = datetime.today().strftime("%Y%m%d")
    
    dst_folder = src_folder
    base_name = f"PO update {today_str}.xlsx"
    po_update_path = os.path.join(dst_folder, base_name)

    # Step 1: Copy latest file
    latest_file = get_latest_file(src_folder)

    # If target exists, find a free name by appending _1, _2, ...
    if os.path.exists(po_update_path):
        counter = 1
        while True:
            candidate = f"PO update {today_str}_{counter}.xlsx"
            candidate_path = os.path.join(dst_folder, candidate)
            if not os.path.exists(candidate_path):
                po_update_path = candidate_path
                break
            counter += 1

    shutil.copy2(latest_file, po_update_path)
    return latest_file, po_update_path



def copy_sheets(source , source_sheet, target_path = "", target_sheet = "",src_range="", tgt_start_col="", values_only=False, new_workbook=False):
    with xw.App(visible=False, add_book=False) as app:
        app.api.AskToUpdateLinks = False  # Disable update links prompt
        src_wb = app.books.open(source,update_links=False)
        src_ws = src_wb.sheets[source_sheet]
        print(f"Opened source file: {source}")
        print(f"Active sheet in source workbook: {src_ws.name}")
        # tgt_wb = app.books.open(target_path, update_links=False)
        # print(f"Opened target file: {target_path}")

        if (not target_path) or (new_workbook): #未指定貼上路徑或指定要貼到尚未創建的檔案，貼到新檔
            print("No target path specified or new_workbook=True, creating new workbook for copied sheet.")
            if not target_path:
                TEMP_FOLDER = r"YOUR_NETWORK_DRIVE\ShipmentTracking\PO update\temp"
                target_path = os.path.join(TEMP_FOLDER, f"Copied_{source_sheet}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
                #target_path = os.path.join(os.path.dirname(source), f"Copied_{source_sheet}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
            if values_only:
                print("values_only=True, copying only values from source range.")
                data = src_ws.range(src_range).value if src_range else src_ws.used_range.value
                new_wb = xw.Book()
                new_ws = new_wb.sheets[0]
                new_ws.name = source_sheet
                new_ws.range("A1:ZZ1000").number_format = "@"  # Set number format to text to preserve leading zeros if any
                new_ws.range("A1").options(expand='table').value = data
                print(f"target_path for new workbook: {target_path}")
                new_wb.save(target_path)
                new_wb.close()
                src_wb.close()            
            else:
                src_ws.api.Copy()
                new_wb = app.books.active  # The new workbook created by the copy
                new_wb.save(target_path)
                new_wb.close()
                src_wb.close()
            print(f"Copied sheet '{source_sheet}' saved as new file: {target_path}")

        else:   #貼到既有檔
            tgt_wb = app.books.open(target_path, update_links=False)
            tgt_ws = tgt_wb.sheets[target_sheet] if target_sheet else tgt_wb.sheets[0]

            if src_range and tgt_start_col:
                if values_only:
                    data = src_ws.range(src_range).value

                    print(f"Type of data: {type(data)}")
                    if isinstance(data, list):
                        print(f"Length of data: {len(data)}")
                        if len(data) > 0 and isinstance(data[0], list):
                            print(f"Length of first row: {len(data[0])}")
                            print("First 3 rows of data:")
                            for row in data[:3]:
                                print(row)
                        else:
                            print("Data is 1D list, first 10 elements:")
                            print(data[:10])
                    else:
                        print("Data is not a list:", data)

                    tgt_ws.range(tgt_start_col).options(expand='table').value = data
                    print(f"Copied values from range '{src_range}' in sheet '{source_sheet}' to '{tgt_start_col}' in target sheet '{tgt_ws.name}'.(values only)")
                else:
                    src_ws.api.Range(src_range).Copy(Destination=tgt_ws.api.Range(tgt_start_col))
                    print(f"Copied range '{src_range}' from sheet '{source_sheet}' to '{tgt_start_col}' in target sheet '{tgt_ws.name}'.")
            else:
                if values_only:
                    data = src_ws.used_range.value
                    tgt_ws.range("A1").options(expand='table').value = data
                    print(f"Copied all values from sheet '{source_sheet}' to target sheet '{tgt_ws.name}'.(values only)")
                else:
                    src_ws.api.UsedRange.Copy(Destination=tgt_ws.api.Range("A1"))
                    print(f"Copied entire sheet '{source_sheet}' to target workbook '{target_path}'.")
            tgt_wb.save()
            tgt_wb.close()


    print(f"Copied sheets saved as: {target_path}")


def create_temp_file(po_update_path, temp_folder):
    print("Copy original file to a temporary file for followed steps...")
    new_file = os.path.join(temp_folder, f"Copied_大表_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx") #暫時檔位於TEMP_FOLDER，檔名會是 "Copied_原檔名_時間.xlsx"，之後會從這個暫時檔把資料複製回主檔
    copy_sheets(source=po_update_path, source_sheet="大表", target_path=new_file, values_only=True, new_workbook=True) #貼到temp資料夾創建新檔
    temp_po_update = get_latest_file(temp_folder) #取得剛剛創建的暫時擋，接下來的程式都在temp_po_update這個檔案上操作，最後再把結果貼回主檔
    print("temp_po_update:", temp_po_update)
    print("Finished copying to temporary file. Starting next modules...")
    print("-" * 50)
    return temp_po_update


def temp_to_main(temp_po_update, po_update_path, copy_range):
    print("Copy data from temporary file back to the main file...")
    with xw.App(visible=False) as app: # #決定要複製到底幾列，若人工操作不佳導致列數過多，則只複製到第4000列以內
        wb = app.books.open(temp_po_update)
        ws = wb.sheets["大表"]
        last_row = ws.range("AO" + str(ws.cells.last_cell.row)).end("up").row
        end_row = min(4000, last_row)
        src_range = f"{copy_range}{end_row}"
        wb.close()
    copy_sheets(source=temp_po_update, source_sheet="大表", target_path=po_update_path, target_sheet="大表", src_range=src_range, tgt_start_col="AO3", values_only=True)
    print("Data copied back to main file")
    print("-" * 50)

def recalculate_workbook(file_path):
    # Recalculate workbook with xlwings to update any formulas that depend on the "Shipments" sheet (Open real excel)
    print("recalculating workbook...")
    wb = xw.Book(file_path)
    wb.app.calculate()
    wb.save()
    wb.close()
    print("recalculation complete and closed workbook.")
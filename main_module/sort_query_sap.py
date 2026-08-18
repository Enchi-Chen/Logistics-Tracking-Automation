import win32com.client

def rgb_to_bgr_int(hex_color):
    # Convert hex string (e.g., 'C6EFCE') to BGR integer for Excel
    hex_color = hex_color.lstrip('#')
    r = int(hex_color[0:2], 16)
    g = int(hex_color[2:4], 16)
    b = int(hex_color[4:6], 16)
    return b << 16 | g << 8 | r


def sort_query_sap(file_path):
    print(f"Starting sort_query_sap module...")

    # Re-apply the existing sort rule
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        wb = excel.Workbooks.Open(file_path)
        ws = wb.Worksheets("大表")

        # Find the "Status" column index in row 2
        for col in range(1, ws.UsedRange.Columns.Count + 1):
            if ws.Cells(2, col).Value == "Status":
                status_col = col
                break
        else:
            raise Exception("Status column not found.")

        # ws.Sort.SortFields.Clear()
        # # Add sort fields for each color in order
        # ws.Sort.SortFields.Clear()
        # for color in ["C6EFCE", "FFEB9C", "FFC7CE"]:
        #     sf = ws.Sort.SortFields.Add(
        #         Key=ws.Range(ws.Cells(3, status_col), ws.Cells(ws.UsedRange.Rows.Count, status_col)),
        #         SortOn=win32com.client.constants.xlSortOnCellColor,
        #         Order=win32com.client.constants.xlDescending,
        #         DataOption=win32com.client.constants.xlSortNormal
        #     )
        # sf.SortOnValue.Color = rgb_to_bgr_int(color)

        # # Set the range to sort (from row 3 to last row, all columns)
        # last_row = ws.UsedRange.Rows.Count
        # last_col = ws.UsedRange.Columns.Count #要改成BA!!!!!!!!!
        # ws.Sort.SetRange(ws.Range(ws.Cells(3, 1), ws.Cells(last_row, last_col)))
        # ws.Sort.Header = win32com.client.constants.xlYes
        # ws.Sort.Apply()
        # print("Sorted by cell color in Status column.")

    except Exception as e:
        print("Error opening workbook or applying sort:", e)
        try:
            wb.Close()
        except:
            pass
        try:
            excel.Quit()
        except:
            pass
        print(f"Ended sort_query_sap module")
        return
    
    # Refresh all queries 
    try:
        wb.RefreshAll()
        print("RefreshAll called, now waiting for async queries to complete...")
        excel.CalculateUntilAsyncQueriesDone()
        print("RefreshAll completed, now running SAPExecuteCommand to refresh data...")
    except Exception as e:
        print("Error during RefreshAll or opening workbook:", e)
        excel.Quit()
        print(f"Ended sort_query_sap module")
        return

    wb.Save()

    # Refresh SAP data using SAPExecuteCommand macro
    try:
        excel.Application.Run("SAPExecuteCommand", "RefreshData")
        print("SAPExecuteCommand called, now waiting for async queries to complete...")
        excel.CalculateUntilAsyncQueriesDone()
        print("SAPExecuteCommand completed.")
    except Exception as e:
        print("Error during SAPExecuteCommand:", e)
        excel.Quit()
        print(f"Ended sort_query_sap module")
        return

    wb.Save()
    wb.Close()
    excel.Quit()
    print(f"Finished sort_query_sap module, workbook saved")
    print("-" * 50)
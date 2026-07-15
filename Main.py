import sales_data_sync
import cargoo
import customs
import carrier_tracking_main as carrier_tracking_main
from support_module import support_function
import warehouse
import compare_last_file
import sort_query_sap
from datetime import datetime

import os
import shutil
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime
import xlwings as xw

## UI

#-----------------------------------------------流程設定區-------------------------------------------------#
SRC_FOLDER = r"YOUR_NETWORK_DRIVE\ShipmentTracking\PO update" #欲更新的主檔所在地
TEMP_FOLDER = r"YOUR_NETWORK_DRIVE\ShipmentTracking\PO update\temp" #暫時檔案存放位址，執行模組3-5時會在這裡創建暫時檔
COPY_RANGE = "AO3:AY"  #由TEMP檔案貼回主檔時需要複製貼上的範圍(模組3,4,5的結果填寫處)

### 一、 下方區塊二選一: 選擇創建新檔或使用現有檔，之後的流程都在 po_update_path 這個檔案上進行

#創建新檔
latest_file, po_update_path = support_function.create_new_file(SRC_FOLDER)
print("latest_file:", latest_file)
print("po_update_path:", po_update_path)

#使用現有檔
# po_update_path = support_function.get_latest_file(SRC_FOLDER)
# print("po_update_path:", po_update_path)


### 二、 從七個模組中選擇要執行的模組（若要取消某個模組，直接把對應那行註解掉即可，使用CTRL + / ）

# 以下兩個模組會直接在主檔做更改
# 1. SalesData 模組
sales_data_sync.sales_data_sync(po_update_path) #會清空後方資料

# 2. Cargoo 模組
cargoo.cargoo(po_update_path) 


# 若要進行模組3-5，程式會複製大表到TEMP_FOLDER存為暫時新檔，此列需為非註解狀態
temp_po_update = support_function.create_temp_file(po_update_path, TEMP_FOLDER) #在TEMP_FOLDER創建主檔的暫時副本，之後模組3-5都在這個暫時檔上操作
# 3. Carrier tracking 模組
carrier_tracking_main.tracker_main(temp_po_update) #取消重算 
# 4. Customs 模組
customs.customs(temp_po_update) 
# 5. Warehouse 模組
warehouse.warehouse(temp_po_update) 
#執行上方步驟完畢後，會將位於COPY_RANGE的資料重新貼到主檔，並且儲存主檔，此列需為非註解狀態
support_function.temp_to_main(temp_po_update, po_update_path, COPY_RANGE)

# 以下兩個模組也會直接在主檔做更改
# 6. Compare last file 模組
compare_last_file.compare_last_file(po_update_path) #視情況取消重算 
# 7. Sort query and SAP 模組
sort_query_sap.sort_query_sap(po_update_path) #最後可能需要人工驗證 #約三分鐘 #完成




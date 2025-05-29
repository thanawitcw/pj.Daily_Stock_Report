import os
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook

# Folder path and file type
folder_path = r'D:\Data for Stock Report\Completed Daily Stock Report'
file_type = '.xlsx'

# Load latest file in the directory
def get_latest_file(directory, file_prefix, file_extension):
    latest_file = [f for f in os.listdir(directory) if f.startswith(file_prefix) and f.endswith(file_extension)]

    # Extract date from filenames:
    dates = []
    for file in latest_file:
        try:
            date_str = file.replace(file_prefix, '').replace(file_extension, '').strip('_')
            date_obj = datetime.strptime(date_str, '%d-%m-%Y')
            dates.append((file, date_obj))
        except ValueError:
            continue

    if dates:
        latest_file = max(dates, key=lambda x: x[1])[0]
        return os.path.join(directory, latest_file)  # Return full file path
    else:
        return None

# Get latest file
latest_file = get_latest_file(folder_path, 'Sahamit Daily Stock Report_', '.xlsx')
if latest_file:
    print(f"Loading file: {latest_file}")
else:
    print("No file found")
    exit()

# Read the two required sheets
df_data_by_cartons = pd.read_excel(latest_file, sheet_name='Data by Cartons', engine='openpyxl')
df_all_po_pending = pd.read_excel(latest_file, sheet_name='All PO Pending', engine='openpyxl')

# Cast SHM / CJ Item to string
df_data_by_cartons = df_data_by_cartons.astype({'SHM_Item': str, 'CJ_Item': str})
df_all_po_pending = df_all_po_pending.astype({'SHM_Item': str, 'CJ_Item': str})

# Create new column in df_data_by_cartons
df_data_by_cartons['ReportDate'] = latest_file.split('_')[-1].replace('.xlsx', '')
df_data_by_cartons['ReportDate'] = pd.to_datetime(df_data_by_cartons['ReportDate'], format='%d-%m-%Y')

# Columns to fill for 'Data by Cartons' sheet
columns_to_fill = [
    'DC1_OOSAssort', 'DC1_CountOKROOS', 'DC1_PercOOS', 'DC1_StoreStockCTN', 'DC1_DOHStore',
    'DC1_AvgSaleCTN_Last90Days', '%Ratio_AvgSalesQty90D_DC1', 'DC1_Remain_CTN', 'Current_DC1_DOH',
    'DC2_OOSAssort', 'DC2_CountOKROOS', 'DC2_PercOOS', 'DC2_StoreStockCTN', 'DC2_DOHStore',
    'DC2_AvgSaleCTN_Last90Days', '%Ratio_AvgSalesQty90D_DC2', 'DC2_Remain_CTN', 'Current_DC2_DOH',
    'DC4_OOSAssort', 'DC4_CountOKROOS', 'DC4_PercOOS', 'DC4_StoreStockCTN', 'DC4_DOHStore',
    'DC4_AvgSaleCTN_Last90Days', '%Ratio_AvgSalesQty90D_DC4', 'DC4_Remain_CTN', 'Current_DC4_DOH'
]

# Fill missing values for 'Data by Cartons' sheet
df_data_by_cartons[columns_to_fill] = df_data_by_cartons.groupby('CJ_Item')[columns_to_fill].transform(
    lambda x: x.replace(0, pd.NA).ffill().bfill().fillna(0)
)

# Merge df_data_by_cartons to get data by MOQ
moq_file_path = r'C:\Users\Thanawit C\OneDrive - Sahamit Product Co.,Ltd\Data for Stock Report\COPY_MasterLeadTime.xlsx'
df_moq_data = pd.read_excel(moq_file_path, sheet_name='All_Product', header=1, engine='openpyxl')

# Handle with CJ_Item column
df_moq_data['CJ_Item'] = df_moq_data['CJ_Item'].astype(str).str.split('.').str[0]

# Merge 2 DF to get data MOQ
df_data_by_cartons = df_data_by_cartons.merge(
    df_moq_data[
        ['SHM_Item', 'CJ_Item',
         'PO_Type_DC1', 'DC1_Group_Code', 'DC1_Max_Pallet_per_group', 'DC1_MOQ_per_group[CTN]', 'DC1_MOQ_per_SKU[CTN]',
         'PO_Type_DC2', 'DC2_Group_Code', 'DC2_Max_Pallet_per_group', 'DC2_MOQ_per_group[CTN]', 'DC2_MOQ_per_SKU[CTN]',
         'PO_Type_DC4', 'DC4_Group_Code', 'DC4_Max_Pallet_per_group', 'DC4_MOQ_per_group[CTN]', 'DC4_MOQ_per_SKU[CTN]']
    ],
    on=['SHM_Item', 'CJ_Item'],
    how='left'
)

# Save path for output file
save_path = r'C:\Users\Thanawit C\Sahamit Product Co.,Ltd\SCM - SCM\15.SCM_Report\1.Stock All Chanel\2.Calculate Forecast by MOQ.xlsx'

# Save both sheets to the existing Excel file
with pd.ExcelWriter(save_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    df_data_by_cartons.to_excel(writer, sheet_name='Data by CTN', index=False)
    df_all_po_pending.to_excel(writer, sheet_name='All PO Pending', index=False)  

print(f"Data cleaning and saving completed. File saved to: {save_path}")

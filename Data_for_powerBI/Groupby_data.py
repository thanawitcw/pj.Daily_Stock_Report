import pandas as pd
import os

# Folder path and file type
folder_path = r'D:\Data for Stock Report\test1'
file_type = '.xlsx'
sheet_name = 'Sheet1'

# Columns to fill
columns_to_fill = [
    'DC1_ScmAssort', 'DC1_OOSAssort', 'DC1_CountOKROOS', 'DC1_PercOOS', 'DC1_StoreStockQty', 'DC1_DOHStore', 'DC1_AvgSaleQty90D',
    '%Ratio_AvgSalesQty90D_DC1', 'DC1_Remain_StockQty', 'Current_DC1_DOH', 'DC2_ScmAssort', 'DC2_OOSAssort', 'DC2_CountOKROOS',
    'DC2_PercOOS', 'DC2_StoreStockQty', 'DC2_DOHStore', 'DC2_AvgSaleQty90D', '%Ratio_AvgSalesQty90D_DC2',
    'DC2_Remain_StockQty', 'Current_DC2_DOH','DC4_ScmAssort', 'DC4_OOSAssort', 'DC4_CountOKROOS', 'DC4_PercOOS', 'DC4_StoreStockQty',
    'DC4_DOHStore', 'DC4_AvgSaleQty90D', '%Ratio_AvgSalesQty90D_DC4', 'DC4_Remain_StockQty', 'Current_DC4_DOH'
]

# Check if the folder path exists
if not os.path.exists(folder_path):
    print(f"The folder path {folder_path} does not exist.")
else:
    # Iterate over each file in the folder
    for filename in os.listdir(folder_path):
        if filename.endswith(file_type):
            file_path = os.path.join(folder_path, filename)
            
            # Read the Excel file
            df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
            
            # Fill the data for specified columns where CJ_Item is the same
            for column in columns_to_fill:
                df[column] = df.groupby('CJ_Item')[column].transform(lambda x: x.replace(0, pd.NA).ffill().bfill().fillna(0))
            
            # Save the cleaned data to a new Excel file
            cleaned_file_path = os.path.join(folder_path, f'cleaned_{filename}')
            df.to_excel(cleaned_file_path, index=False, engine='openpyxl')

    print("Data cleaning and saving completed.")

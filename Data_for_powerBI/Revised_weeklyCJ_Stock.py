import pandas as pd
import os
from datetime import datetime

# Define a function to get the latest Excel file
def get_latest_file(file_path, file_prefix, file_extension):
    files = [file for file in os.listdir(file_path)
             if file.startswith(file_prefix) and file.endswith(file_extension)]
    
    dated_files = []
    for file in files:
        try:
            date_str = file.replace(file_prefix, '').replace(file_extension, '').strip()
            date_obj = datetime.strptime(date_str, '%d-%m-%Y')
            dated_files.append((file, date_obj))
        except ValueError:
            continue

    if dated_files:
        return max(dated_files, key=lambda x: x[1])[0]
    else:
        return None

# Process the latest file
def clean_cj_stock(file_path, file_prefix, file_extension, output_folder):
    latest_file = get_latest_file(file_path, file_prefix, file_extension)

    if latest_file:
        print(f"Loading latest file: {latest_file}")
        full_path = os.path.join(file_path, latest_file)
        cj_stock = pd.read_excel(full_path, sheet_name='Sahamit Report', header=2)

        # Fill NA with 0 in specific columns
        columns_to_fill = [
            'DC1_DCStockQty', 'DC1_DCStockValue', 'DC1_DOHDC',
            'DC2_DCStockQty', 'DC2_DCStockValue', 'DC2_DOHDC',
            'DC4_DCStockQty', 'DC4_DCStockValue', 'DC4_DOHDC'
        ]
        cj_stock[columns_to_fill] = cj_stock[columns_to_fill].fillna(0)

        # Load DC daily stock data
        dc_stock_path = r'D:\Data for Stock Report\cleaned_DC_daily_stock.xlsx'
        dc_stock = pd.read_excel(dc_stock_path, sheet_name='Pivot_DC_stock')
        print(f"Loading DC stock data from: {dc_stock_path}")
        # Change the column name for mergeing
        dc_stock = dc_stock.rename(columns={'CJ_Item': 'Product'})

        # Merge dataframes
        cj_stock = cj_stock.merge(dc_stock, on='Product', how='left')

        # Replace target columns if data exists in merged ones
        replace_map = {
            'DC1_DCStockQty': 'DC1_Remain_StockQty',
            'DC1_DCStockValue': 'DC1_Remain_StockValue',
            'DC2_DCStockQty': 'DC2_Remain_StockQty',
            'DC2_DCStockValue': 'DC2_Remain_StockValue',
            'DC4_DCStockQty': 'DC4_Remain_StockQty',
            'DC4_DCStockValue': 'DC4_Remain_StockValue'
        }

        for main_col, merged_col in replace_map.items():
            if merged_col in cj_stock.columns:
                cj_stock[main_col] = cj_stock[main_col].where(cj_stock[merged_col].isnull(), cj_stock[merged_col])


        # Replace null value in avg_col with 0
        for dc in ['DC1', 'DC2']:
            avg_col = f'{dc}_AvgSaleQty90D'
            stock_col = f'{dc}_DCStockQty'
            doh_col = f'{dc}_DOHDC'

            if avg_col in cj_stock.columns:
                cj_stock[avg_col] = cj_stock[avg_col].fillna(0) 

            if stock_col in cj_stock.columns and avg_col in cj_stock.columns:
                cj_stock[doh_col] = cj_stock[stock_col] / cj_stock[avg_col]
                
                # Replace infinite values with 365 in column doh_col
                cj_stock[doh_col] = cj_stock[doh_col].replace([float('inf'), -float('inf')], 365)
                cj_stock[doh_col] = cj_stock[doh_col].fillna(0)

        # Drop the merged columns
        cj_stock = cj_stock.drop(columns=list(replace_map.values()))

        # Define saved file name
        output_file_name = latest_file
        output_file_path = os.path.join(output_folder, output_file_name)

        # Save the result
        cj_stock.to_excel(output_file_path, index=False)
        print(f"Cleaned file saved to: {output_file_path}")

    else:
        print("No valid files found.")

# Run the function
clean_cj_stock(
    file_path= r'T:\SCM Data\Data For Stock\DC_Store',
    file_prefix= 'Sahamit Report ',
    file_extension= '.xlsx',
    output_folder= r'T:\SCM Data\Data for PowerBI\cleaned_weekly'
    )

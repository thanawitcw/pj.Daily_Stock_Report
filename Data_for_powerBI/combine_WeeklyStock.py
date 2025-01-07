# Hold this python code because of all files already cleaned and saved in myenv/Data_for_powerBI/Weekly_Stock
import os
import pandas as pd

# Define files path
input_path = r'T:\SCM Data\Data For Stock\DC_Store'
output_path = r'T:\SCM Data\Data for PowerBI\cleaned_weekly'
subfolders = ['2025']

# Loop through subfolders
for subfolder in subfolders:
    subfolder_path = os.path.join(input_path, subfolder)
    try:
        # Walk through directory and subdirectories
        for root, dirs, files in os.walk(subfolder_path):
            for file in files:
                if file.endswith('.xlsx'):
                    file_path = os.path.join(root, file)
                    output_file = os.path.join(output_path, f'{subfolder}_{file}')
                    try:
                        # Read specific sheet and skip first 2 rows
                        df = pd.read_excel(file_path, sheet_name='Sahamit Report',header=2)

                        # Save the cleaned data using openpyxl for speed
                        df.to_excel(output_file, index=False, engine='openpyxl')
                        print(f"Successfully processed:{output_file}")
                    except ValueError as ve: # Catch if sheet name does not exist
                        print(f"Error: Sheet 'Sahamit Report' not found in {file_path}. Error: {ve}")
                    except Exception as e:
                        print(f'Error processing {file}: {e}')
    except FileNotFoundError:
        print(f"Subfolder not found: {subfolder_path}")
    except Exception as e:
        print(f"An unexpected error occurred processing year {subfolder}: {e}")

print(f"Finished processing. Cleaned Excel files saved to '{output_path}'")
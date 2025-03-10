import os
import pandas as pd
from datetime import datetime

# Define files path
input_path = r'T:\SCM Data\Data For Stock\DC_Store'
output_path = r'T:\SCM Data\Data for PowerBI\cleaned_weekly'
subfolders = ['2025']

# Create output directory if it does not exist
if not os.path.exists(output_path):
    os.makedirs(output_path)

# Load latest file in the directory
def get_latest_file(directory, file_prefix, file_extension):
    latest_file = [f for f in os.listdir(directory) if f.startswith(file_prefix) and f.endswith(file_extension)]

    # extract date from filenames:
    dates = []
    for file in latest_file:
        try:
            date_str = file.replace(file_prefix, '').replace(file_extension, '').strip()
            date_obj = datetime.strptime(date_str, '%d-%m-%Y')
            dates.append((file,date_obj))
        except ValueError:
            continue

    if dates:
        latest_file = max(dates, key=lambda x: x[1])[0]
        return latest_file
    else:
        return None

# Get latest file
latest_file = get_latest_file(input_path, 'Sahamit Report', '.xlsx')
if latest_file:
    print(f"Load the: {latest_file}")
else:
    print("No file found")
    

# Load the latest file
latest_file_path = os.path.join(input_path, latest_file)
df = pd.read_excel(latest_file_path, sheet_name='Sahamit Report', header=2)
# Save the cleaned data using openpyxl for speed
output_file = os.path.join(output_path, f'{latest_file}')
df.to_excel(output_file, index=False, engine='openpyxl')
print(f"Successfully processed:{output_file}")


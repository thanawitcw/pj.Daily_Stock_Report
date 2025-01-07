import os
import shutil

# Define files path
input_path = r'T:\SCM Data\Sale Out\Daily_SO'
output_path = r'T:\SCM Data\Data for PowerBI\Daily_SellOut'
subfolders = ['2024','2025']

# Loop through subfolders
for subfolder in subfolders:
    subfolder_path = os.path.join(input_path, subfolder)
    try:
        files = os.listdir(subfolder_path)
        for file in files:
            if file.endswith('.xlsx'):
                file_path = os.path.join(subfolder_path, file)
                output_file = os.path.join(output_path, f'{subfolder}_{file}')
                try:
                    shutil.copy2(file_path, output_file)  # Copy file with metadata
                    print(f"Successfully processed: {file_path} -> {output_file}")
                except Exception as e:
                    print(f'Error copying {file}: {e}')
    except FileNotFoundError:
        print(f"Subfolder not found: {subfolder_path}")

print(f"All Excel files has been save to '{output_path}'")




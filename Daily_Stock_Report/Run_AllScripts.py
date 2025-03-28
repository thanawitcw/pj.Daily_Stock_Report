import os
from datetime import datetime
from nbconvert import PythonExporter
import nbformat
#import schedule
#import time

# Paths to the Jupyter Notebook scripts
notebook_paths = [
    r'C:\Users\Thanawit C\OneDrive - Sahamit Product Co.,Ltd\Desktop\MyPython\myenv\Daily_Stock_Report\1.cleaned_CJ Stock Report.ipynb',
    r'C:\Users\Thanawit C\OneDrive - Sahamit Product Co.,Ltd\Desktop\MyPython\myenv\Daily_Stock_Report\2.cleaned_DC-Stock.ipynb',
    r'C:\Users\Thanawit C\OneDrive - Sahamit Product Co.,Ltd\Desktop\MyPython\myenv\Daily_Stock_Report\3.cleaned-SellOut.ipynb',
    r'C:\Users\Thanawit C\OneDrive - Sahamit Product Co.,Ltd\Desktop\MyPython\myenv\Daily_Stock_Report\4.cleaned-PO_HBA.ipynb',
    r'C:\Users\Thanawit C\OneDrive - Sahamit Product Co.,Ltd\Desktop\MyPython\myenv\Daily_Stock_Report\5.cleaned-PO_Access.ipynb',
    r'C:\Users\Thanawit C\OneDrive - Sahamit Product Co.,Ltd\Desktop\MyPython\myenv\Daily_Stock_Report\6.cleaned-PO_Import.ipynb',
    r'C:\Users\Thanawit C\OneDrive - Sahamit Product Co.,Ltd\Desktop\MyPython\myenv\Daily_Stock_Report\7.merged-AllPO.ipynb',
    r'C:\Users\Thanawit C\OneDrive - Sahamit Product Co.,Ltd\Desktop\MyPython\myenv\Daily_Stock_Report\8.merged-Stock_Report.ipynb'
]

# Function to convert Jupyter Notebook to Python script and execute it
def run_notebook(notebook_path):
    print(f"Running notebook: {notebook_path}")
    # Handle with .xlsb
    if '4.Copy&Cleaned_PO Pending H&BA.ipynb' in notebook_path:
        with open(notebook_path, 'rb') as f:
            nb = nbformat.read(f, as_version=4)
    else:
        with open(notebook_path, 'r', encoding='utf-8') as f:
            nb = nbformat.read(f, as_version=4)
    
    exporter = PythonExporter()
    source, _ = exporter.from_notebook_node(nb)
    
    code = compile(source, notebook_path, 'exec')
    exec(code, globals())
    print(f"Finished running notebook: {notebook_path}")

# Function to run all notebooks
def run_all_notebooks():
    for notebook_path in notebook_paths:
        run_notebook(notebook_path)

# Run the code Right now
run_all_notebooks()

# Schedule the task
#schedule.every().day.at("09:30").do(run_all_notebooks)

#print("Scheduler started. Waiting for the scheduled time...")

# Keep the script running
#while True:
    #schedule.run_pending()
    #time.sleep(1)

import pandas as pd
import sys
sys.stdout.reconfigure(encoding='utf-8')

file_path = r"d:\LQP\Project\ghn-gtalk-user\[Gtalk Expansion] Workforce Analysis.xlsx"
xl = pd.ExcelFile(file_path)

sheets = ['data1', 'data history', 'Pivot analysis', 'Pivot Table']

for sheet in sheets:
    print(f"\n--- Sheet: {sheet} ---")
    df = xl.parse(sheet, nrows=5)
    print("Columns:", list(df.columns))
    print(df.head())
    

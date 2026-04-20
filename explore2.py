import pandas as pd
import sys
sys.stdout.reconfigure(encoding='utf-8')

file_path = r"d:\LQP\Project\ghn-gtalk-user\[Gtalk Expansion] Workforce Analysis.xlsx"
xl = pd.ExcelFile(file_path)
df = xl.parse("data history", nrows=0)
print([str(c) for c in df.columns])

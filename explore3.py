import pandas as pd
import sys
sys.stdout.reconfigure(encoding='utf-8')

file_path = r"d:\LQP\Project\ghn-gtalk-user\[Gtalk Expansion] Workforce Analysis.xlsx"
df_h = pd.read_excel(file_path, sheet_name='data history')

print('Total active per date:')
for c in df_h.columns:
    print(f'  {c}: {df_h[c].dropna().nunique()} unique IDs')

print()
df_s = pd.read_excel(file_path, sheet_name='data1')
print('Departments per division:')
for div in df_s['division_name_vn'].unique():
    depts = df_s[df_s['division_name_vn']==div]['department_name_vn'].nunique()
    print(f'  {div}: {depts} departments, {len(df_s[df_s["division_name_vn"]==div])} staff')

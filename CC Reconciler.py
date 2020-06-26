import pandas as pd
import xlrd

df1 = pd.read_excel(r'C:\Users\rkeenan\OneDrive - Aurora Cooperative\Documents\Development\Account Reconciler\Copy of 01-05 Banking Entries.xlsx', skiprows=1, usecols=11)

print(df1)

df2 = pd.read_excel(r'C:\Users\rkeenan\OneDrive - Aurora Cooperative\Documents\Development\Account Reconciler\Copy of May Credit Card Reconciliation to GL.xlsx', sheet_name=0, skiprows=3, usecols=18)

print(df2)

with open(r'C:\Users\rkeenan\OneDrive - Aurora Cooperative\Documents\Development\Account Reconciler\WriteFile1.csv', 'w') as f1:
    df1.to_csv(f1)

with open(r'C:\Users\rkeenan\OneDrive - Aurora Cooperative\Documents\Development\Account Reconciler\WriteFile2.csv', 'w') as f2:
    df2.to_csv(f2)


INNER_JOIN = pd.merge(df1, df2, how='inner', left_on=['Amount', 'Location Name'], right_on=['Debit', 'Location'])
#outer_join_df = pd.merge(df1, df2, on='')

print(INNER_JOIN)

with open(r'C:\Users\rkeenan\OneDrive - Aurora Cooperative\Documents\Development\Account Reconciler\WriteFile3.csv', 'w') as f3:
    INNER_JOIN.to_csv(f3)
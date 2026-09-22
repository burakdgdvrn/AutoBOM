import pandas as pd

df = pd.read_excel('E-26-RNS-01_BOM_Listesi (1).xlsx', nrows=5)
print('Headers:')
print(df.columns.tolist())
print('\nFirst row:')
for col in df.columns:
    print(f"  {col}: {df.iloc[0][col]}")

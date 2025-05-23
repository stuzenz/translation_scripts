import pandas as pd
try:
    df = pd.read_excel('../input_files/Draft.xlsx', sheet_name=None)
    print('Pandas can read it!')
    print('Sheets:', list(df.keys()))
    for name, sheet_df in df.items():
        print(f'Sheet {name}: {sheet_df.shape}')
except Exception as e:
    print('Pandas also fails:', e)

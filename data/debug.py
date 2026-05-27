import pandas as pd
df = pd.read_excel('/Users/streiten/customs/bao/data/示例发票.xlsx', dtype=str)
print('Columns:', list(df.columns))
print('Shape:', df.shape)
print()
print(df.head(16).to_string())

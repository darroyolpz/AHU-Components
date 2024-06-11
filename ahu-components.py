import pandas as pd
import numpy as np

# Import sales file
sales_file = 'Sales.xlsx'
cols = ['Order no', 'Line no', 'Item no', 'Date', 'Amount', 'Cost amount', 'Ref order no']
df_sales = pd.read_excel(sales_file, usecols = cols)

# Check formats
for col in cols:
    df_sales[col] = df_sales[col].astype(str)

# Rename columns
df_sales = df_sales.rename(columns={'Ref order no': 'MO no'})

# Import MOs
mo_file = '645-100.xlsx'
cols = ['MO no', 'Component no']
df_mo = pd.read_excel(mo_file, usecols = cols)

# Check formats
for col in cols:
    df_mo[col] = df_mo[col].astype(str)
    
# Import Components
comp_file = 'Components.xlsx'
df_comp = pd.read_excel(comp_file)
df_comp['Component no'] = df_comp['Component no'].astype(str)

# Merge MO with Component
df_map = pd.merge(df_mo, df_comp, on='Component no')

# Get Pivot
df_pivot = pd.pivot_table(df_map, values = 'Component no', index = 'MO no',
                          columns = 'Component', aggfunc = 'count')
df_pivot = pd.DataFrame(df_pivot.to_records())

# Create only one check for each component
div = df_pivot.iloc[:, 1:]/df_pivot.iloc[:, 1:]
df_pivot = pd.concat([df_pivot['MO no'], div], axis = 1)
df_pivot = df_pivot.fillna(0)

# Merge to the final dataframe and remove zero value rows. Rearrange the columns
df_final = pd.merge(df_sales, df_pivot, on='MO no')
df_final = df_final.fillna(0)
#df_final = df_final[df_final['Order no'] != 0]
print(df_final.head())

# Export to Excel
name = 'AHU-Components - Results.xlsx'
df_final.to_excel(name, index = False)
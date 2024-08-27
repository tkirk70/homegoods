import pandas as pd
import streamlit as st

from datetime import date
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils.dataframe import dataframe_to_rows  # Import the missing function
from io import BytesIO


st.title("The Home Good Database App 🗂️ ")

st.caption("""
         **In Extensiv**:
         \n1. Export the transaction as an excel file.
         \n2. Upload to this application.
         \n3. Download the transformed file.
         \n4. Replace the current_homegoods file in the tds folder with the downloaded one.
         \n5. Open BarTender file and make sure it is connected to the DataBase.
         \n6. Run print preview before printing.
         \n7. Enjoy!
         """)

uploaded_file = st.file_uploader("Choose a file")
# Can be used wherever a "file-like" object is accepted:
df = pd.read_excel(uploaded_file, dytpe='str')
st.write(df)

df['Ship To Zip'] = df['Ship To Zip'].astype(str)

df['Ship To Zip'] = df['Ship To Zip'].str.zfill(5)

df_split = df['SKU/Quantity'].str.split(",", expand=True)

df_split.columns.to_list()

split_list = [item+1 for item in df_split.columns.to_list()]

# Adding these columns to the existing dataframe. 
df[split_list] = df['SKU/Quantity'].apply(lambda x: pd.Series(str(x).split(","))) 

df_items = df[split_list]

df_T = df_items.T

df_T = df_T[0].str.split("(", expand=True)

df_T[1] = df_T[1].str.replace(")", "")

# Repeating each row
result = df_T.loc[df_T.index.repeat(df_T[1])]

needed_columns = ['Transaction ID',
         'Reference Number',
         'Ship To Company',
         'Customer',
         'Ship To Name',
         'Purchase Order',
         'Ship To Address',
         'Ship To Address 2',
         'Ship To City',
         'Ship To Country',
         'Ship To State',
         'Ship To Zip',
         'Total Item Qty'
         ]
        

for item in needed_columns:
    result[item] = df[item].iloc[0]

# Rename Multiple Columns by Index
result.rename(columns={result.columns[0]: 'SKU', result.columns[1]: 'Quantity'},inplace=True)

result['Dept'] = 54

# add dictionary and column to translate wms sku to customer sku

SKU = [
'CS9001',
'CS9002',
'CS9009',
'CS9010',
'CS9020',
'CS9030',
'CS9040',
'CS9050',
'CS9070',
'CS9080',
'CS9090',
'CS3040',
'CS3041',
'CS2015',
'CS2016',
'CS2017',
'CS2018',
'CS2020',
'CS2021',
'CS2031',
'CS2032',
'CS4020',
'CS4022',
'CS4023',
'CS4024',
'CS4026',
'CS6020',
'CS6021',
'CS7001',
'CS7002',
'CS7003',
'CS7004',
'CSMK100',
]

Cust_SKU = [
'320948',
'326519',
'326526',
'320943',
'326527',
'326514',
'',
'320958',
'326531',
'326522',
'320956',
'320982',
'320986',
'326536',
'320960',
'320969',
'326537',
'326553',
'320965',
'',
'326562',
'320972',
'326558',
'326554',
'320979',
'320974',
'',
'',
'',
'',
'',
'',
'320951',
]

cust_sku_dict = dict(zip(SKU, Cust_SKU))

# Remap the values of the dataframe
result['Cust_SKU'] = result['SKU'].map(cust_sku_dict)

excel = result.to_excel('current_homegoods.xlsx', index=None)


st.download_button(label="Download Home Goods template as excel file",
                       data=excel,
                       file_name=f"current_homegoods.xlsx",
                       mime='application/octet-stream')
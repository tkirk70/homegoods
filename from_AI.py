import pandas as pd
import streamlit as st

from datetime import date
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils.dataframe import dataframe_to_rows
from io import BytesIO

image = 'https://www.homegoods.com/img/header/homegoods-logo.svg'
st.image(image, use_column_width='auto', clamp=False, channels="RGB", output_format="auto")

st.title("The Home Goods Database App 🗂️")

st.caption("""
    **In Extensiv**:
    \n1. Export the transaction as an excel file.
    \n2. Upload to this application.
    \n3. Download the transformed file.
    \n4. Replace the current_homegoods file in this folder: C:\Users\PickPack\Desktop\HomeGoods.
    \n5. Open BarTender file C:\Users\PickPack\Desktop\BarTender Suite\HomeGoods\homegoods_test.btw.
    \n6. Make sure it is connected to the DataBase and refresh.
    \n7. Run print preview before printing.
    \n8. Enjoy!
""")

uploaded_file = st.file_uploader("Choose a file", type=["xlsx"])

if uploaded_file is not None:
    st.write('Preview of Upload')
    try:
        df = pd.read_excel(uploaded_file)
        # Assuming df is your DataFrame
        df['Transaction ID'] = df['Transaction ID'].apply(lambda x: f'{x:.0f}')
        st.write(df)
    except Exception as e:
        st.write('An Error Occurred.')
        st.error(f"Error reading the file: {e}")

    df['Ship To Zip'] = df['Ship To Zip'].astype(str)
    df['Ship To Zip'] = df['Ship To Zip'].str.zfill(5)

    df_split = df['SKU/Quantity'].str.split(",", expand=True)
    df_split.columns = [f'Item_{i+1}' for i in df_split.columns]

    df = df.join(df_split)

    df_items = df[df_split.columns]
    df_T = df_items.T
    df_T = df_T[0].str.split("(", expand=True)
    df_T[1] = df_T[1].str.replace(")", "")

    result = df_T.loc[df_T.index.repeat(df_T[1].astype(int))]

    needed_columns = [
        'Transaction ID', 'Reference Number', 'Ship To Company', 'Customer',
        'Ship To Name', 'Purchase Order', 'Ship To Address', 'Ship To Address 2',
        'Ship To City', 'Ship To Country', 'Ship To State', 'Ship To Zip', 'Total Item Qty'
    ]

    for item in needed_columns:
        result[item] = df[item].iloc[0]

    result.rename(columns={result.columns[0]: 'SKU', result.columns[1]: 'Quantity'}, inplace=True)
    result['Dept'] = 54

    SKU = [
        'CS9001', 'CS9002', 'CS9009', 'CS9010', 'CS9020', 'CS9030', 'CS9040', 'CS9050',
        'CS9070', 'CS9080', 'CS9090', 'CS3040', 'CS3041', 'CS2015', 'CS2016', 'CS2017',
        'CS2018', 'CS2020', 'CS2021', 'CS2031', 'CS2032', 'CS4020', 'CS4022', 'CS4023',
        'CS4024', 'CS4026', 'CS6020', 'CS6021', 'CS7001', 'CS7002', 'CS7003', 'CS7004',
        'CSMK100'
    ]

    Cust_SKU = [
        '320948', '326519', '326526', '320943', '326527', '326514', '', '320958', '326531',
        '326522', '320956', '320982', '320986', '326536', '320960', '320969', '326537',
        '326553', '320965', '', '326562', '320972', '326558', '326554', '320979', '320974',
        '', '', '', '', '', '320951'
    ]

    cust_sku_dict = dict(zip(SKU, Cust_SKU))
    result['Cust_SKU'] = result['SKU'].map(cust_sku_dict)

    st.write('Preview of Download')
    st.write(result)

    # Convert DataFrame to BytesIO object
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        result.to_excel(writer, index=False)
    output.seek(0)

    st.download_button(
        label="Download Home Goods template as excel file",
        data=output,
        file_name="current_homegoods.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# Custom CSS style for the text
custom_style = '<div style="text-align: right; font-size: 20px;">✨ A TDS Application ✨</div>'

# Render the styled text using st.markdown
st.markdown(custom_style, unsafe_allow_html=True)

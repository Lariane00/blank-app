import pandas as pd
import streamlit as st
from io import BytesIO

import openpyxl
# def process_rmt_data(input_path, output_path):
#     # 读取输入 Excel 文件
#     RMT_full_excel = pd.read_excel(input_path)
#     RMT_full_list = RMT_full_excel.values.tolist()

#     RMT = []
#     switch = 0
#     Project = ''
#     Power = ''
#     Energy = ''

#     # 解析输入数据
#     for i in RMT_full_list:
#         if 'Project name' in str(i[0]):
#             Project = i[1]
#         elif 'Capacity' in str(i[0]):
#             Capacity = i[1]
#             Power, Energy = Capacity.split('/')
#         if switch == 1:
#             RMT.append(i[2])
#         if 'Requirements' in str(i[2]):
#             switch = 1

#     # 创建数据字典
#     data = {
#         'Region': [''] * len(RMT),
#         'Project': [Project] * len(RMT),
#         'Power (MW)': [Power] * len(RMT),
#         'Capacity (MWh)': [Energy] * len(RMT),
#         'Product': [''] * len(RMT),
#         'Category': [''] * len(RMT),
#         'RMT': RMT
#     }

#     # 创建 DataFrame
#     df = pd.DataFrame(data)

#     # 读取输出 Excel 文件中的已有数据
#     RMT_towrite = pd.read_excel(output_path, sheet_name='RMT database')
#     start_row = RMT_towrite.shape[0]

#     # 将新的数据追加到 Excel 文件中
#     with pd.ExcelWriter(output_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
#         df.to_excel(writer, sheet_name='RMT database', startrow=start_row + 1, index=False, header=False)

# # 调用函数
# input_path = 'Excel tool/RMT - Kallista.xlsx'
# output_path = 'Excel tool/OR and RMT database_to test.xlsx'
# process_rmt_data(input_path, output_path)

def process_rmt_data(uploaded_files,Baseline):
    
    # 读取输入 Excel 文件
    ini_data ={'Region': [''] ,
            'Project': [''],
            'Power (MW)': [''],
            'Capacity (MWh)': [''],
            'Product':[''],
            'Category': [''],
            'RMT': [''],
            'RMT details': ['']
        }
    long_df = pd.DataFrame(ini_data)
    for RMT_full_excel in uploaded_files:
        RMT_full_read_excel = pd.read_excel(RMT_full_excel)
        RMT_full_list = RMT_full_read_excel.values.tolist()

        RMT = []
        switch = 0
        Project = ''
        Power = ''
        Energy = ''
        RMT_details = []
        # 解析输入数据
        for i in RMT_full_list:
            if 'Project name' in str(i[0]):
                Project = i[1]
            elif 'Capacity' in str(i[0]):
                Capacity = i[1]
                Power, Energy = Capacity.split('/')
            if switch == 1:
                RMT.append(i[2])
                RMT_details.append(i[3])
            if 'Requirements' in str(i[2]):
                switch = 1

        data = {
            'Region': [''] * len(RMT),
            'Project': [Project] * len(RMT),
            'Power (MW)': [Power] * len(RMT),
            'Capacity (MWh)': [Energy] * len(RMT),
            'Product': [''] * len(RMT),
            'Category': [''] * len(RMT),
            'RMT': RMT,
            'RMT details': RMT_details
        }

        # 创建 DataFrame
        df = pd.DataFrame(data)
        long_df = pd.concat([long_df,df],axis = 0)
    # 读取输出 Excel 文件中的已有数据
    RMT_towrite = pd.read_excel(Baseline, sheet_name='RMT database')
    start_row = RMT_towrite.shape[0]
    output = BytesIO()
    # 将新的数据追加到 Excel 文件中
    with pd.ExcelWriter(Baseline, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        long_df.to_excel(writer, sheet_name='RMT database', startrow=start_row + 1, index=False, header=False)
  
    output.seek(0)
    st.download_button(
    label="Download combined RMT as xlsx",
    data=Baseline,
    file_name='RMT database.xlsx',
    mime='text/csv'
    )

st.title("File Processor")
uploaded_files = st.file_uploader("Choose the new RMT template",accept_multiple_files=True)

Baseline = st.file_uploader("Choose the baseline to be added")


if uploaded_files is not None and Baseline is not None:
    process_rmt_data(uploaded_files,Baseline)


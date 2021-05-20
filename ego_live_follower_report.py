# To: LLZ ChildWolf
# Author: 最上静香
# API impletemented by https://vtbs.moe/about.
# Version: v0.0.1

import requests as r
import pandas as pd
import os
import json
import time
from datetime import date
from openpyxl import load_workbook

def append_df_to_excel(filename, df, sheet_name='Sheet1', startrow=None,
                       truncate_sheet=False, 
                       **to_excel_kwargs):
    # Excel file doesn't exist 
    # Or the excel file is empty
    # Save and exit
    excel_df = pd.read_excel(filename)
    if not os.path.isfile(filename) or excel_df.empty:
        df.to_excel(
            filename,
            index=False,
            sheet_name=sheet_name, 
            startrow=startrow if startrow is not None else 0, 
            **to_excel_kwargs)
        return
    
    # ignore [engine] parameter if it was passed
    if 'engine' in to_excel_kwargs:
        to_excel_kwargs.pop('engine')

    writer = pd.ExcelWriter(filename, engine='openpyxl', mode='a')

    # try to open an existing workbook
    writer.book = load_workbook(filename)
    
    # get the last row in the existing Excel sheet
    # if it was not specified explicitly
    if startrow is None and sheet_name in writer.book.sheetnames:
        startrow = writer.book[sheet_name].max_row

    # truncate sheet
    if truncate_sheet and sheet_name in writer.book.sheetnames:
        # index of [sheet_name] sheet
        idx = writer.book.sheetnames.index(sheet_name)
        # remove [sheet_name]
        writer.book.remove(writer.book.worksheets[idx])
        # create an empty sheet [sheet_name] using old index
        writer.book.create_sheet(sheet_name, idx)
    
    # copy existing sheets
    writer.sheets = {ws.title:ws for ws in writer.book.worksheets}

    if startrow is None:
        startrow = 0

    # erase title when start appending
    # write out the new sheet
    df.to_excel(writer, sheet_name, header=False, index=False, startrow=startrow, **to_excel_kwargs, )

    # save the workbook
    writer.save()





vtb_full_info_url = "https://api.vtbs.moe/v1/info"
ego_room_ids = [7194103, 1086621, 22631364, 475577, 22588330, 11312, 10413051, 673595, 22572737, 22580086, 3923305, 22800243, 22595698, 3000303, 22605289]

def decode_ego_info():
    with open('vtb_info.json', encoding='utf_8_sig') as file:
        vtbjson = file.read()

    decoded_ego_live_members =  json.loads(vtbjson, object_hook=ego_decoder)
    ego_live_info_list = list(filter(None, decoded_ego_live_members))
    for ego_members_info in ego_live_info_list:
        ego_members_info['DateTime'] = date.today().strftime('%Y-%m-%d')
    return ego_live_info_list

def get_vtb_info():
    vtb = r.get(vtb_full_info_url).json()
    vtb_json_file = json.dumps(vtb)
    write_file_to_local(vtb_json_file)

def write_file_to_local(vtb):
    with open('vtb_info.json', 'w', encoding='utf_8_sig') as file:
        file.write(vtb)


# def delete_local_file():
#     if os.path.exists('vtb_info.txt'):
#         os.remove('vtb_info.txt')

def ego_decoder(dct):
    if 'roomid' in dct:
        if dct['roomid'] in ego_room_ids:
            return dct

def create_dataframe(df_list):
    df = pd.DataFrame(df_list)
    return df


get_vtb_info()
ego_list = decode_ego_info()
df = create_dataframe(ego_list)
append_df_to_excel('ego_live_report.xlsx', df)




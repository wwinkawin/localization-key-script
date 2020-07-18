import pandas as pd
import json
from datetime import datetime
import os.path
from os import path
import xlsxwriter

def main():
    while True:
        file_name = input("Excel file name: ")
        if path.exists(file_name):
            break
        else:
            print ("File path does not exist")

    # sheet_name = input("Excel sheet name (Leave empty if default): ")
    # if not sheet_name:
    #     sheet_name = 'NEXT MVP2 Delta Localization Ke'

    while True:
        fileVersion = input("Version: ")
        if fileVersion.isnumeric():
            break
        else:
            print ("Plase enter a version number")
    

    while True:
        try:
            keys = pd.read_excel(file_name,sheet_name = ['FE Localization Keys','BE Status Codes'])
        except PermissionError:
            decision = input('Please close file: ' + str(file_name) + '\nPress Enter')
            if decision == "":
                        continue
        except Exception as e:
            print('Unexpected error:' + str(e))
        break
        
    #Duplicate checking
    keys_name = {
        'FE Localization Keys' : 'String Name',
        'BE Status Codes': 'Status Code'}
    
    duplicate_dict = duplicate_check(keys, keys_name)
    FE_processor(keys.get("FE Localization Keys"), fileVersion, duplicate_dict.get("FE Localization Keys"))
    
    
def duplicate_check(keys,keys_name):
    
    df_dict = {}
    duplicate_dict = {}
    
    #Get BE and FE sheets
    for sheet_name, column_name in keys_name.items():        
        sheet = keys.get(sheet_name)
        
        #Create isDuplucated column
        isDuplicated = pd.DataFrame(data = sheet.duplicated([column_name],keep=False),columns=['isDuplicated'])
        
        duplicate_dict[sheet_name] = isDuplicated
        
        #check if there are any duplicated rows
        if (isDuplicated.values.sum() != 0):
            #If there is, append dict with key = sheet_name, values = dataframe
            df_dict[sheet_name] = pd.concat([sheet,isDuplicated],axis=1)

    #Create excel file if there are duplicated rows
    if df_dict:
        while True:
            try:
                with pd.ExcelWriter('Localization_key_isDuplicated.xlsx') as writer:
                    for sheet_name, df in df_dict.items():
                        df.to_excel(writer, sheet_name=sheet_name, index = False)
                        
            except xlsxwriter.exceptions.FileCreateError:
                decision = input("Cannot write to Excel file.\nPlease close "
                                  "'Localization_key_isDuplicated.xlsx'\n"
                                  "Press Enter")
                if decision == "":
                    continue
            break
    return duplicate_dict

def BE_processor(keys):
    column_list = list(range(6))
    for i in column_list:
        keys.iloc[:,i] = keys.iloc[:,i].str.strip()
        
    #Create isDuplucated check column
    isDuplicated = pd.DataFrame(data = keys.duplicated(['Status Code'],keep=False),columns=['isDuplicated'])
        
        
def FE_processor(keys, fileVersion, isDuplicated):
    #remove whitespace
    column_list = [0,1,2]
    for i in column_list:
        keys.iloc[:,i] = keys.iloc[:,i].str.strip()

    #Get Current DateTime in specified format
    currentTime = str(datetime.now().strftime("%Y-%m-%dT%H:%M:%S+07:00"))
    
    #Remove duplicated rows
    keys = keys[isDuplicated.values == False]
    
    #Split dataframe
    en = keys.iloc[:,[0,1]].set_index('String Name')
    en.rename(columns={en.columns[0]:'en'},inplace=True)
    
    th = keys.iloc[:,[0,2]].set_index('String Name')
    th.rename(columns={th.columns[0]:'th'},inplace=True)
    
    #Convert dataframe to dictionaries
    en_dict = en.to_dict()
    th_dict = th.to_dict()
    
    #Combine dictionaries
    data = {
        "versionNumber":fileVersion,
        "languagePackLastModified":currentTime,
        "content":{**en_dict, **th_dict}
    }
    
    #Gen file name
    file_name = 'language_pack_' + str(datetime.now().strftime("%Y%m%d")) + "_1700_v" + str(fileVersion) + ".json"
    
    #Wrtie JSON file
    json_writer(data, file_name)
    

def json_writer(data,file_name):
    with open(file_name, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=4)
    
if __name__ == '__main__':
    main()
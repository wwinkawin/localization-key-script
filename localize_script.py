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
    

    try:
        keys = pd.read_excel(file_name,sheet_name = ['FE Localization Keys','BE Status Codes'])
    except Exception as e:
        print('Unexpected error:' + str(e))
        
    FE_processor(keys.get("FE Localization Keys"), fileVersion)
    
    # #remove whitespace
    # column_list = [0,1,2]
    # for i in column_list:
    #     keys.iloc[:,i] = keys.iloc[:,i].str.strip()
    
    # #Create isDuplucated check column
    # isDuplicated = pd.DataFrame(data = keys.duplicated(['String Name'],keep=False),columns=['isDuplicated'])
    
    
    # #Concat isDuplucated check column
    # keys = pd.concat([keys,isDuplicated],axis=1)
    
    # while True:
    #     try:
    #         keys.to_excel('Localization_key_isDuplicated.xlsx', sheet_name = sheet_name,index = False)
    #     except xlsxwriter.exceptions.FileCreateError as e:
    #         decision = input("Cannot write to Excel file.\nPlease close "
    #                          "'Localization_key_isDuplicated.xlsx'\n"
    #                          "Press Enter")
    #         if decision == "":
    #             continue
    #     break

    # #Get Current DateTime
    
    # currentTime = str(datetime.now().strftime("%Y-%m-%dT%H:%M:%S+07:00"))
    
    # #Remove duplicated rows
    # keys = keys[keys['isDuplicated'] == False]
    
    
    # #Split dataframe
    # en = keys.iloc[:,[0,1]].set_index('String Name')
    # en.rename(columns={en.columns[0]:'en'},inplace=True)
    
    # th = keys.iloc[:,[0,2]].set_index('String Name')
    # th.rename(columns={th.columns[0]:'th'},inplace=True)
    
    # #Convert dataframe to dictionaries
    # en_dict = en.to_dict()
    # th_dict = th.to_dict()
    
    # #Combine dictionaries
    # data = {
    #     "versionNumber":fileVersion,
    #     "languagePackLastModified":currentTime,
    #     "content":{**en_dict, **th_dict}
    # }
    
    
    # #Write JSON file
    
    # with open('data.json', 'w', encoding='utf-8') as f:
    #     json.dump(data, f, ensure_ascii=False, indent=4)

def FE_processor(keys, fileVersion):
    #remove whitespace
    column_list = [0,1,2]
    for i in column_list:
        keys.iloc[:,i] = keys.iloc[:,i].str.strip()
    
    #Create isDuplucated check column
    isDuplicated = pd.DataFrame(data = keys.duplicated(['String Name'],keep=False),columns=['isDuplicated'])
    
    
    if (isDuplicated.sum() != 0):
        
        #Concat isDuplucated check column
        keys = pd.concat([keys,isDuplicated],axis=1)
        
        #Write isDuplicated excel file
        while True:
            try:
                keys.to_excel('Localization_key_isDuplicated.xlsx', sheet_name = 'FE Localization Keys',index = False)
            except xlsxwriter.exceptions.FileCreateError:
                decision = input("Cannot write to Excel file.\nPlease close "
                                 "'Localization_key_isDuplicated.xlsx'\n"
                                 "Press Enter")
                if decision == "":
                    continue
            break

    #Get Current DateTime in specified format
    currentTime = str(datetime.now().strftime("%Y-%m-%dT%H:%M:%S+07:00"))
    
    #Remove duplicated rows
    keys = keys[keys['isDuplicated'] == False]
    
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
    
    #Write JSON file
    fileName = 'language_pack_' + str(datetime.now().strftime("%Y%m%d")) + "_1700_v" + str(fileVersion) + ".json"
    with open(fileName, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=4)
    
if __name__ == '__main__':
    main()
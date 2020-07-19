# -*- coding: utf-8 -*-
"""
Created on Sun Jul 19 15:47:29 2020

@author: wwink
"""
import pandas as pd
import json
from datetime import datetime
import os.path
from os import path
import xlsxwriter



class ExcelProcessor:
    
    #Duplicate checking
    keys_name = {
        'FE Localization Keys' : 'String Name',
        'BE Status Codes': 'Status Code'}
    sheet_name = ['FE Localization Keys','BE Status Codes']
   
    def __init__(self):
        pass
        
       
        
    
        # while True:
        #     self.file_version = input("Version: ")
        #     if self.file_version.isnumeric():
        #         break
        #     else:
        #         print ("Plase enter a version number")
        
    
        
            
        
        
        # duplicate_dict = self.duplicate_check(keys, keys_name)
        # self.FE_processor(keys.get("FE Localization Keys"), self.file_version, duplicate_dict.get("FE Localization Keys"))
        # self.BE_processor(keys.get('BE Status Codes'),self.file_version, duplicate_dict.get('BE Status Codes'))
    
    def set_attr(self,file_path, file_version):
       self.file_path = file_path
       self.file_version = file_version
       print(file_path)
    
    def read_excel(self):
        while True:
            try:
                self.keys = pd.read_excel(self.file_path, sheet_name = self.sheet_name)
            except PermissionError:
                decision = input('Please close file: ' + str(self.file_path) + '\nPress Enter')
                if decision == "":
                            continue
            except Exception as e:
                print('Unexpected error:' + str(e))
            break
    
    def duplicate_check(self):
        
        df_dict = {}
        duplicate_dict = {}
        keys = self.keys
        
        #Get BE and FE sheets
        for sheet_name, column_name in self.keys_name.items():        
            sheet = keys.get(sheet_name)
            
            #Create isDuplicated column
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
        self.duplicate_dict = duplicate_dict
        return
    
    def BE_processor(self, isDuplicated):
        keys = self.keys
        
        #remove whitespace
        column_list = list(range(6))
        for i in column_list:
            keys.iloc[:,i] = keys.iloc[:,i].str.strip()
            
        #Remove duplicated rows
        keys = keys[self.duplicate_dict.get('BE Status Codes').values == False]
        
        #Rename first 6 columns
        new_columns = keys.columns.values
        new_columns[1:6] = ['description','titleEn','titleTh','messageEn','messageTh']
        keys.columns  = new_columns
        
        keys = keys[keys.iloc[:,0].astype(str).str.isnumeric()]
        
        #Change first column to be index
        keys.set_index(str(keys.columns.values[0]), inplace=True, drop=True)
        
        keys = keys.loc[:,['description','titleEn','titleTh','messageEn','messageTh']]
    
        data = keys.to_dict('index')
        
        file_name = 'beStatusCodes-' + str(datetime.now().strftime("%Y%m%d")) + "_" + str(datetime.now().strftime("%H%M")) +"_v" + self.file_version
        
        #Wrtie JSON file
        self.json_writer(data, file_name)
            
    def FE_processor(self, isDuplicated):
        keys = self.keys
        
        #remove whitespace
        column_list = [0,1,2]
        for i in column_list:
            keys.iloc[:,i] = keys.iloc[:,i].str.strip()
    
        #Get Current DateTime in specified format
        currentTime = str(datetime.now().strftime("%Y-%m-%dT%H:%M:%S+07:00"))
        
        #Remove duplicated rows
        keys = keys[self.duplicate_dict.get("FE Localization Keys").values == False]
        
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
            "versionNumber":self.file_version,
            "languagePackLastModified":currentTime,
            "content":{**en_dict, **th_dict}
        }
        
        #Gen file name
        file_name = 'language_pack_' + str(datetime.now().strftime("%Y%m%d")) + "_" + str(datetime.now().strftime("%H%M")) + "_v" + str(self.file_version)
        
        #Wrtie JSON file
        self.json_writer(data, file_name)
        
    
    def json_writer(data,file_name):
        with open(file_name + ".json", 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=4)
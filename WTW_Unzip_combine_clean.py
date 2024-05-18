# -*- coding: utf-8 -*-
"""
Created on Wed May 31 15:30:47 2023

@author: Jackson_Aquino
"""
import os
import zipfile
import shutil

# INSERT YOUR ZIP FILE FULL PATH HERE (INCLUDING EXTENSION):
zip_file_path = r"C:\Users\MyUserName\Downloads\MyFile.zip"


def extract_xlsx_files(zip_file_path):
    with zipfile.ZipFile(zip_file_path, 'r') as zip_file:
        output_folder = os.path.join(os.path.dirname(zip_file_path), 'WTW Files')
        os.makedirs(output_folder, exist_ok=True)
        
        for inner_zip_file_name in zip_file.namelist():
            if inner_zip_file_name.endswith('.zip'):
                country_name = os.path.splitext(inner_zip_file_name)[0].split(' — ')[-1]
                with zip_file.open(inner_zip_file_name) as inner_zip_file:
                    with zipfile.ZipFile(inner_zip_file, 'r') as inner_zip:
                        for file_name in inner_zip.namelist():
                            if file_name.startswith('Compensation Report/') and file_name.endswith('.xlsx'):
                                if "Function, Discipline, Career Level, Survey Grade " in file_name:
                                    output_file_name = file_name.replace(".xlsx","") + f"{country_name}.xlsx"
                                    output_file_name = output_file_name.replace(r'Compensation Report/','')
#                                    if file_name != "Compensation Report/Incumbent-Weighted Results.xlsx":
#                                        output_file_name = f"Incumbent-Weighted Results - {country_name}" + file_name.replace("Compensation Report/Incumbent-Weighted Results","").replace(".xlsx","").strip() +".xlsx"
                                    output_file_name = output_file_name.replace("/","\\")                                    
                                    output_file_path = os.path.join(output_folder, output_file_name)
                                    print(output_file_path)
                                    with inner_zip.open(file_name) as xlsx_file:
                                        with open(output_file_path, 'wb') as output_file:
                                            shutil.copyfileobj(xlsx_file, output_file)


# Specify the path to your main zip file
# Get this from a text file to integrate with Excel, parse from the ZipFile= param and add back an OutputFile= param that Excel will open afterward

extract_xlsx_files(zip_file_path)

import pandas as pd
import numpy

folder_path = os.path.join(os.path.dirname(zip_file_path), 'WTW Files')  # Specify the path to your folder containing Excel files

# Create an empty DataFrame to store the consolidated results
AllResults = pd.DataFrame()

ColumnsOfInterest = ["Effective Date", "Scope", "Currency", "Job Code", "Job Title", "Base Salary #Incs", "Base Salary #Orgs", "Base Salary Average", "Base Salary 25th", "Base Salary 50th", "Base Salary 75th", "Base Salary 90th", "Target Total Annual Incentives #Incs", "Target Total Annual Incentives #Orgs", "Target Total Annual Incentives Average", "Target Total Annual Incentives 25th", "Target Total Annual Incentives 50th", "Target Total Annual Incentives 75th", "Target Total Annual Incentives 90th", "Target Total Annual Compensation #Incs", "Target Total Annual Compensation #Orgs", "Target Total Annual Compensation Average", "Target Total Annual Compensation 25th", "Target Total Annual Compensation 50th", "Target Total Annual Compensation 75th", "Target Total Annual Compensation 90th", "Long-Term Incentive #Incs", "Long-Term Incentive #Orgs", "Long-Term Incentive Average", "Long-Term Incentive 25th", "Long-Term Incentive 50th", "Long-Term Incentive 75th", "Long-Term Incentive 90th", "Target Total Direct Compensation #Incs", "Target Total Direct Compensation #Orgs", "Target Total Direct Compensation Average", "Target Total Direct Compensation 25th", "Target Total Direct Compensation 50th", "Target Total Direct Compensation 75th", "Target Total Direct Compensation 90th"]

# Loop through the files in the folder
for file_name in os.listdir(folder_path):
    if file_name.endswith('.xlsx'):
        file_path = os.path.join(folder_path, file_name)
        
        # Read the Excel file and get the "Results" sheet (adjust the sheet name if necessary)
        df = pd.read_excel(file_path, sheet_name='Results')
        dfinfo = pd.read_excel(file_path, sheet_name='Information')
        dfinfo = dfinfo.set_index("Product Name")
        currency = dfinfo.loc["Currency Displayed"].values[0]
        effdate = dfinfo.loc["Effective Date"].values[0]
        weighting = dfinfo.loc["Weighting"].values[0]
        if weighting == 'Incumbent':
            print("\n",file_name.replace("Incumbent-Weighted Results - ","").replace(".xlsx",""))
            if "Geographic Scope" in df.columns:
                pass
            else:
                if "Scope" in df.columns:
                    df["Geographic Scope"] = df["Scope"]
                    print("Using Scope instead of Geographic Scope")
                else:
                    print("Scope column not found")
                    df["Geographic Scope"] = 'All - Geographic Scope'
            df["Scope"] = file_name.replace("Incumbent-Weighted Results - ","").replace(".xlsx","").strip()
            df["Currency"] = currency
            df["Effective Date"] = effdate
            for col in ColumnsOfInterest:
                if col in df.columns:
                    pass
                else:
                    print(col, "not found")
            # Append the data to the consolidated DataFrame
            listdfs = [AllResults,df]
            AllResults = pd.concat(listdfs, ignore_index=True)
        else:
            print('ignoring',file_path,":",weighting)

folder_path = os.path.dirname(zip_file_path)
Backup = AllResults
# AllResults = AllResults[AllResults["Geographic Scope"]=="All - Geographic Scope"]
AllResults = AllResults[AllResults['Geographic Scope'].isin(['Total Sample', 'All - Geographic Scope'])]
#print(Backup["Geographic Scope"].unique())

AllResults = AllResults[ColumnsOfInterest]

AllResults.replace("--", "", inplace=True)

AllResults.to_excel(folder_path + "\\WTW Results.xlsx",index=False)

# -*-coding = utf-8 -*-
import pandas as pd
import re
import os
import datetime

def readFile(filename, usecols, skiprows=0):
    print("开始读取并处理: " + filename)
    if filename.split(".")[-1] == "csv":
        for i in ["UTF-8", "GBK"]:
            try:
                data = pd.read_csv(filename, encoding=i, skiprows=skiprows, low_memory=False, error_bad_lines=False,header=None)
            except Exception as e:
                print(e)
            else:
                break
        else:
            print(filename+"有问题，请保存为xlsx格式后重试。")
    else:
        data = pd.read_excel(filename, skiprows=skiprows ,header=None)
    data = data.dropna(axis=0, how="all")
    data.columns = data.iloc[0,:]

    data=data.drop([0]).reset_index(drop=True)
    print("总共"+str(data.shape[0])+"行数据")
    if usecols:
        return data[usecols]
    else:
        return data



day = datetime.datetime.now().strftime('%m.%d')

path = os.getcwd() + "\\"
dirs = os.listdir(path)
filelist=[]
typeList=["csv","xlsx","xls"]
for file in dirs:
    filetype = file.split(".")[-1]
    if filetype in typeList:
        filelist.append(file)
try:
    os.mkdir("结果")
except:
    pass
for file in filelist:
    data = readFile(file,[])
    data.rename(columns={'售价':'price'}, inplace = True)
    if 'price' not in data.columns:
        data.rename(columns={'vip_price':'price'}, inplace = True)
        data = data.astype({"price": 'float'})
    else:
        data = data.astype({"price": 'float'})
        data["price"]=data["price"]/100
    if 'fullProductName' in data.columns:
        data.rename(columns={'product_name':'product_name_origin'}, inplace = True)
        data.rename(columns={'fullProductName':'product_name'}, inplace = True)
    data["规格"] = 1
    for i in range(data.shape[0]):
        r= re.findall("\d[0-9a-zA-Z*/.±+-×]*[^0-9]*",data.loc[i,"product_name"])
        if r:
            data.loc[i,"规格"] = r[-1]
    
    data.to_excel(path+"结果"+"\\"+file.split(".")[0]+day+".xlsx",index=False)




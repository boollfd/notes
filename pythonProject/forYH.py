import datetime
import numpy as np
import pandas as pd
import re
import os

def readFile(filename):

    return pd.read_csv if filename.split(".")[-1] == "csv" else pd.read_excel

# def yhDataProcessing(filename,num,deletename=""):
#     yHProductDatatype = readFile(filename)
    
#     if yHProductDatatype == pd.read_csv:
#         yHProductData = readFile(filename)(filename,usecols=[0,1,2,3,4,5,6,7],low_memory=False)
#     else:
#         yHProductData = readFile(filename)(filename,usecols=[0,1,2,3,4,5,6,7])
#     yHProductData = yHProductData.dropna(axis=0, how="all")
#     yHProductData.columns = ["友商编码","友商品项","友商价","秒杀价","友商规格","活动信息","店铺名称","友商价执行门店数"]
#     if deletename:
#         yHProductData = yHProductData[yHProductData["店铺名称"].str.contains(deletename)==False]
#     # yHProductData = yHProductData[~yHProductData["友商品项"].isin(["清仓"])]
#     # yHProductData = yHProductData[~yHProductData["友商编码"].isin(["N","S"])]
#     yHProductData = yHProductData.astype({"友商规格":'str'})
#     yHProductData = yHProductData[["友商编码","友商品项","友商价","友商规格","友商价执行门店数"]].dropna(axis=0, how="any")
#     yHProductData = yHProductData[yHProductData["友商品项"].str.contains("清仓")==False]
#     yHProductData = yHProductData[yHProductData["友商编码"].str.contains("N")==False]
#     yHProductData = yHProductData[yHProductData["友商编码"].str.contains("S")==False]
#     yHProductData["友商价"] = yHProductData["友商价"].apply(lambda x: int(re.findall("\d+",str(x))[0])).apply(lambda x: int(x)/100)
#     yHProductData["友商编码"] = yHProductData["友商编码"].apply(lambda x: int(re.findall("\d+",str(x))[0]))
    
#     pvt = pd.pivot_table(yHProductData, index=["友商编码", "友商品项", "友商规格", "友商价"], values=["友商价执行门店数"], aggfunc="count").reset_index()
    
#     pvt = pvt.sort_values(by='友商价执行门店数', ascending=False).drop_duplicates(["友商编码"])
#     pvt = pvt[pvt["友商价执行门店数"] > num]
    
#     return pvt
def yhDataProcessing(filename,num,deletename=""):
    yHProductDatatype = readFile(filename)
    
    if yHProductDatatype == pd.read_csv:
        yHProductData = readFile(filename)(filename,usecols=[0,1,2,3,4,5,6,7],low_memory=False)
    else:
        yHProductData = readFile(filename)(filename,usecols=[0,1,2,3,4,5,6,7])
    writer = pd.ExcelWriter(path+filename.split(".")[0]+".xlsx")
    yHProductData = yHProductData.dropna(axis=0, how="all")
    yHProductData.columns = ["友商编码","友商品项","友商价","秒杀价","友商规格","活动信息","店铺名称","友商价执行门店数"]
    if deletename:
        yHProductData = yHProductData[yHProductData["店铺名称"].str.contains(deletename)==False]
    # yHProductData = yHProductData[~yHProductData["友商品项"].isin(["清仓"])]
    # yHProductData = yHProductData[~yHProductData["友商编码"].isin(["N","S"])]
    yHProductData = yHProductData.dropna(axis=0, how="all")
    
    yHProductData["ascii编码"] = [ord(str(i)[0]) for i in yHProductData["友商编码"]]
    yHProductData = yHProductData[yHProductData["ascii编码"]<255]
    yHProductData.drop(["ascii编码"], axis=1, inplace=True)
    
    yHProductData = yHProductData.astype({"友商规格":'str'})
    yHProductData = yHProductData[yHProductData["友商品项"].str.contains("清仓")==False]
    yHProductData = yHProductData[yHProductData["友商编码"].str.contains("N")==False]
    yHProductData = yHProductData[yHProductData["友商编码"].str.contains("S")==False]
    
    yHProductData["友商编码"] = yHProductData["友商编码"].apply(lambda x: int(re.findall("\d+",str(x))[0]))
    yHProductData.to_excel(writer,sheet_name = filename.split(".")[0],index = False)
    yHProductData["友商价"] = yHProductData["友商价"].apply(lambda x: int(re.findall("\d+",str(x))[0])).apply(lambda x: int(x)/100)
    yHProductData = yHProductData[["友商编码","友商品项","友商价","友商规格","友商价执行门店数"]].dropna(axis=0, how="any")

    pvt = pd.pivot_table(yHProductData, index=["友商编码", "友商品项", "友商规格", "友商价"], values=["友商价执行门店数"], aggfunc="count").reset_index()
    
    pvt = pvt.sort_values(by='友商价执行门店数', ascending=False).drop_duplicates(["友商编码"])
    pvt = pvt[pvt["友商价执行门店数"] > num]
    pvt.to_excel(writer,sheet_name = filename.split(".")[0]+"透视",index = False)
    writer.save()
yHProduct = ""
yHProduct2 = ""
path = os.getcwd()+"\\"
files = os.listdir(path)
# 有一个库可以使用正则匹配文件
day = datetime.datetime.now().strftime('%Y%m%d')

for i in files:
    if "_永辉商户产品" in i and i[0] != "~":
        if yHProduct == "":
            yHProduct = i
    elif day in i and i[0] != "~":
        if yHProduct2 == "":
            yHProduct2 = i
    else:
        pass



if yHProduct:
    # yhDataProcessing(yHProduct,3,deletename="").to_excel(writer,sheet_name = yHProduct.split(".")[0],index = False)
    yhDataProcessing(yHProduct,0,deletename="")
    
if yHProduct2:
    # yhDataProcessing(yHProduct2,5,deletename="万象").to_excel(writer,sheet_name = yHProduct2.split(".")[0],index = False)
    yhDataProcessing(yHProduct2,0,deletename="万象")



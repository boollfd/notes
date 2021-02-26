# -*-coding = utf-8 -*-
import datetime
import numpy as np
import pandas as pd
import re
import os




# def readFile(filename):
#     return pd.read_csv if filename.split(".")[-1] == "csv" else pd.read_excel


# def readFile(filename, usecols, skiprows=0):
#     print("开始读取并处理: " + filename)
#     if filename.split(".")[-1] == "csv":
#         for i in ["UTF-8", "GBK"]:
#             try:
#                 if usecols:
#                     data = pd.read_csv(filename, usecols=usecols, encoding=i, skiprows=skiprows, low_memory=False)
#                 else:
#                     data = pd.read_csv(filename, encoding=i, skiprows=skiprows, low_memory=False)
#             except Exception as e:
#                 print(e)
#             else:
#                 return data
#         else:
#             print(filename+"的编码有问题，请保存为xlsx格式后重试。")
#     else:
#         if usecols:
#             data = pd.read_excel(filename, usecols=usecols, skiprows=skiprows)
#         else:
#             data = pd.read_excel(filename, skiprows=skiprows)
#         return data
    
    
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


def merge(left, right):
    return pd.merge(left, right, how="left")


def yhDataProcessing(filename):
    num=0
    deletename=""
    if "永辉" in filename:
        num = 3
    else:
        num = 5
        deletename = "万象"
    
    data = readFile(filename,[])
    data = data.iloc[:,:8]
    data = data.dropna(axis=0, how="all")
    data.columns = ["友商编码", "友商品项", "友商价", "秒杀价", "友商规格", "活动信息", "店铺名称", "友商价执行门店数"]
    data["ascii编码"] = [ord(str(i)[0]) for i in data["友商编码"]]
    data = data[data["ascii编码"]<255]
    data.drop(["ascii编码"], axis=1, inplace=True)
    if deletename:
        data = data[data["店铺名称"].str.contains(deletename) == False]
    data = data.astype({"友商规格": 'str', "友商编码": 'str'})
    data = data.dropna(axis=0, subset = ["友商编码", "友商品项", "友商价", "友商规格", "友商价执行门店数"])
    data = data[data["友商品项"].str.contains("清仓") == False]
    data = data[data["友商编码"].str.contains("N") == False]
    data = data[data["友商编码"].str.contains("S") == False]
    data["友商价"] = data["友商价"].apply(lambda x: int(re.findall("\d+", str(x))[0])).apply(
        lambda x: int(x) / 100)
    data["友商编码"] = data["友商编码"].apply(lambda x: int(re.findall("\d+", str(x))[0]))
    
    
    if "永辉" in filename:
        global yHCashBack
        global yHFlashSale
        yHData= data[data["活动信息"].notnull()]
        yHCashBack = yHData[yHData["活动信息"] != "秒杀价"]["友商编码"].drop_duplicates()
        yHFlashSale = yHData[yHData["活动信息"] == "秒杀价"]["友商编码"].drop_duplicates()
        yHData = ""
    
    
    
    pvt = pd.pivot_table(data, index=["友商编码", "友商品项", "友商价", "友商规格"], values=["友商价执行门店数"], aggfunc="count").reset_index()
    # pvt = pd.merge(pvt, data[["友商编码", "活动信息"]], how="left")
    pvt = pvt.sort_values(by='友商价执行门店数', ascending=False).drop_duplicates(["友商编码"])
    pvt = pvt[pvt["友商价执行门店数"] > num]
    if "永辉" in filename:
        pvt["仓或门店"] = "仓"
    else:
        pvt["仓或门店"] = "门店"
    return pvt

# 无价格换算
def pupuExcelDataProcessing(filename):
    
    data = readFile(filename, ["名称", "售价", "券额", " 满", "减", "活动类型","活动介绍"])
    # data = readFile(filename, ["名称", "售价", "活动介绍", " 满", "减", "活动类型","领券"])
    data.columns = ["品名", "当日平台价格", "朴朴优惠券", "购买数量", "赠品数量", "标签","领券"]
    data["购买数量"].fillna(0,inplace = True)
    data["赠品数量"].fillna(0,inplace = True)
    data = data.astype({ "购买数量": 'float',"赠品数量":'float',"当日平台价格":'float'})
    data["朴朴优惠券"].fillna(data["领券"], inplace=True)
    data["购买金额"] = data["购买数量"]
    data["折扣金额"] = data["赠品数量"]
    for i in range(len(data["朴朴优惠券"])):
        if not pd.isna(data.loc[i, "朴朴优惠券"]):
            if "满" in data.loc[i, "朴朴优惠券"] and "减" in data.loc[i, "朴朴优惠券"]:
                data.loc[i, "标签"] = "满减"
            elif "买" in data.loc[i, "朴朴优惠券"] and "赠" in data.loc[i, "朴朴优惠券"]:
                data.loc[i, "标签"] = "买赠"
            elif "领" in data.loc[i, "朴朴优惠券"] and "券" in data.loc[i, "朴朴优惠券"]:
                data.loc[i, "标签"] = "领券"
                
            # if pd.isna(data.loc[i, "标签"]) and "满" in data.loc[i, "朴朴优惠券"] and "减" in data.loc[i, "朴朴优惠券"]:
            #     data.loc[i, "标签"] = "满减"
            # elif pd.isna(data.loc[i, "标签"]) and "买" in data.loc[i, "朴朴优惠券"] and "赠" in data.loc[i, "朴朴优惠券"]:
            #     data.loc[i, "标签"] = "买赠"
            # elif pd.isna(data.loc[i, "标签"]) and "领" in data.loc[i, "朴朴优惠券"] and "券" in data.loc[i, "朴朴优惠券"]:
            #     data.loc[i, "标签"] = "领券"
            
        if data["标签"][i] == "满减":
            data.loc[i, "购买数量"] = ""
            data.loc[i, "赠品数量"] = ""
            reNum = re.findall(r'\d+\.?\d*', data["朴朴优惠券"][i])
            if len(reNum) > 1:
                data.loc[i, "购买金额"] = int(reNum[-2])
                data.loc[i, "折扣金额"] = int(reNum[-1])
        elif data["标签"][i] == "买赠":
            data.loc[i, "购买金额"] = ""
            data.loc[i, "折扣金额"] = ""
        else:
            data.loc[i, "购买数量"] = ""
            data.loc[i, "赠品数量"] = ""
            data.loc[i, "购买金额"] = ""
            data.loc[i, "折扣金额"] = ""
        
    return data[["品名", "当日平台价格", "朴朴优惠券", "标签", "购买数量", "赠品数量", "购买金额", "折扣金额"]]


# 从初始数据就换算
# def pupuExcelDataProcessing(filename):
    
#     data = readFile(filename, ["名称", "售价", "券额", " 满", "减", "活动类型","活动介绍"])
#     # data = readFile(filename, ["名称", "售价", "活动介绍", " 满", "减", "活动类型","领券"])
#     data.columns = ["品名", "当日平台价格", "朴朴优惠券", "购买数量", "赠品数量", "标签","领券"]
#     data["购买数量"].fillna(0,inplace = True)
#     data["赠品数量"].fillna(0,inplace = True)
#     data = data.astype({ "购买数量": 'float',"赠品数量":'float',"当日平台价格":'float'})
#     data["朴朴优惠券"].fillna(data["领券"], inplace=True)
#     data["购买金额"] = data["购买数量"]
#     data["折扣金额"] = data["赠品数量"]
    
#     print("开始计算优惠券信息和换算价格")
#     for i in range(len(data["朴朴优惠券"])):
#         if not pd.isna(data.loc[i, "朴朴优惠券"]):
#             if "满" in data.loc[i, "朴朴优惠券"] and "减" in data.loc[i, "朴朴优惠券"]:
#                 data.loc[i, "标签"] = "满减"
#             elif "买" in data.loc[i, "朴朴优惠券"] and "赠" in data.loc[i, "朴朴优惠券"]:
#                 data.loc[i, "标签"] = "买赠"
#             elif "领" in data.loc[i, "朴朴优惠券"] and "券" in data.loc[i, "朴朴优惠券"]:
#                 data.loc[i, "标签"] = "领券"
                
#         if data["标签"][i] == "满减":
#             data.loc[i, "购买数量"] = ""
#             data.loc[i, "赠品数量"] = ""
#             reNum = re.findall(r'\d+\.?\d*', data["朴朴优惠券"][i])
#             if len(reNum) > 1:
#                 data.loc[i, "购买金额"] = int(reNum[-2])
#                 data.loc[i, "折扣金额"] = int(reNum[-1])
#                 if data.loc[i, "购买金额"]<=data.loc[i, "当日平台价格"]:
#                     data.loc[i, "当日平台价格"] = data.loc[i, "当日平台价格"]-data.loc[i, "折扣金额"]
#                 else:
#                     data.loc[i, "当日平台价格"] = data.loc[i, "当日平台价格"]*((data.loc[i, "购买金额"]-data.loc[i, "折扣金额"])/data.loc[i, "购买金额"])
#         elif data["标签"][i] == "买赠":
#             data.loc[i, "购买金额"] = ""
#             data.loc[i, "折扣金额"] = ""
#             data.loc[i, "当日平台价格"] = data.loc[i, "当日平台价格"]*data.loc[i, "购买数量"]/(data.loc[i, "购买数量"]+data.loc[i, "赠品数量"])
#         else:
#             data.loc[i, "购买数量"] = ""
#             data.loc[i, "赠品数量"] = ""
#             data.loc[i, "购买金额"] = ""
#             data.loc[i, "折扣金额"] = ""
        
#     return data[["品名", "当日平台价格", "朴朴优惠券", "标签", "购买数量", "赠品数量", "购买金额", "折扣金额"]]



def storeCheckDataProcessing(filename):
    data = readFile(filename, ["商品编号", "批次成本"])
    data.columns = ["商品编码", "当日成本"]
    data = data.astype({ "商品编码": 'int',"当日成本":'float'})
    return data


def cityProductDataProcessing(filename):
    data = readFile(filename, [], skiprows=1)
    data = data.iloc[:,:17]
    data = data[["商品编码", "最新售价"]]
    data.columns = ["商品编码", "当日平台价格"]
    data = data.astype({ "商品编码": 'int',"当日平台价格":'float'})
    return data


def yHBenefitDataProcessing(filename):
    data = readFile(filename, [])
    data = data.iloc[:,:7]
    data.columns = ["友商编码","品名","价格", "原价","规格","友商优惠券","店铺ID"]
    
    data["友商编码"] = data["友商编码"].apply(lambda x: int(re.findall("\d+", str(x))[0]))
    return data[["友商编码","友商优惠券"]]


def storeBalanceDataProcessing(filename):
    data = readFile(filename, [])
    col = data.columns
    lengthOfCol = len(col)
    city = col[-1].split("-")[0]
    storeList = [city + "-" + str(i).rjust(2, "0") for i in range(lengthOfCol) if city + "-" + str(i).rjust(2, "0") in col]
    print("门店列表：" + ",".join(storeList))
    
    data["库存为0门店数"] = [[str(j) for j in data.loc[i, storeList]].count("0") for i in range(data.shape[0])]
    data["2级财务类别"] = data["2级财务类别"].apply(lambda x: x.split("-")[-1])
    data["2级财务类别"] = data["2级财务类别"].replace({"品类原名": "改后名",
                                                    })
    data["3级财务类别"] = data["3级财务类别"].apply(lambda x: int(x.split("-")[0]))
    data["库存合计"] = data["库存合计"].replace({"--":"0"})
    data["库存合计"] = data["库存合计"].apply(lambda x:int(str(x).replace(",", "")))
    data = data[["商品编码", "商品名", "2级财务类别", "销售规格", "库存合计", "是否上线", "是否计划下架", "库存为0门店数", "3级财务类别"]]
    data.columns = ["商品编码", "品名", "品类", "规格", "库存合计", "是否上线", "是否计划下架", "库存为0门店", "3级编号"]
    
    
    data = data.astype({ "商品编码": 'int',"库存合计": 'int',"是否上线": 'int',"是否计划下架": 'int',})
    return data


def productListDataDataProcessing(filename):
    data = readFile(filename, ["商品编码", "合并编码"], skiprows=2)
    data = data.T.drop_duplicates().T
    data.columns = ["商品编码", "友商编码"]
    data = data.astype({ "商品编码": 'int',"友商编码": 'int'})
    return data

def productsPUPUProcessing(filename):
    data = readFile(filename, ["商品编码","场景"])
    data = data.astype({ "商品编码": 'int'})
    return data

def outputData(writer,data,sheet_name):
    data.to_excel(writer,sheet_name = sheet_name,index = False)

def dropduplicates(data):
    if "场景" in data.columns:
        return data.drop_duplicates(["商品编码","场景"])
    else:
        return data.drop_duplicates(["商品编码"])


def dataProcessing(files, flag):
    
    functionDic = {"朴朴": pupuExcelDataProcessing,
                    "永辉商户产品": yhDataProcessing,
                    "城市商品": cityProductDataProcessing,
                    "带优惠产品信息": yHBenefitDataProcessing,
                    "statistics": storeBalanceDataProcessing,
                    "商品库存盘点": storeCheckDataProcessing,
                    "清单": productsPUPUProcessing,
                    "友商竞品": productListDataDataProcessing,
                    day: yhDataProcessing
    }
    
    fileList = ["statistics","友商竞品","朴朴","商品库存盘点","带优惠产品信息","永辉商户产品"]
    flag1 = 0
    data = 0
    for i in range(len(fileList)):
        if files[fileList[i]] and flag1 == 0:
            data = functionDic[fileList[i]](files[fileList[i]])
            flag1=1
        elif files[fileList[i]]:
            data = merge(data,functionDic[fileList[i]](files[fileList[i]]))
    if files[day]:
        data1 = data[data["友商品项"].isnull()]
        data2 = data[data["友商品项"].notnull()]
        data = ""
        data1.drop(["友商品项", "友商价", "友商规格", "友商价执行门店数", "仓或门店"], axis=1, inplace=True)
        data1 = merge(data1,functionDic[day](files[day]))
        data = data1.append(data2)
        data1 = ""
        data2 = ""
    data["跟价"] = "/"
    data = data.sort_values(["友商价"],ascending=False)
    writer = pd.ExcelWriter("result"+day+".xlsx")
    wordsOrder = ["品类","场景","商品编码","品名","规格","当日成本","当日平台价格","朴朴优惠券","标签","购买数量","赠品数量","购买金额","折扣金额","友商编码","友商品项","友商规格","友商价","友商价执行门店数","友商优惠券","跟价","仓或门店","库存合计","是否上线","是否计划下架","库存为0门店","3级编号"]



    
    if files["清单"]:
        productsData = functionDic["清单"](files["清单"])
        dataForProducts = merge(productsData,data)
        # if files["城市商品"] and "场景" in dataForProducts.columns and len(dataForProducts[dataForProducts["场景"] == "S级商品"])>0:
        #     data1 = dataForProducts[dataForProducts["场景"] == "S级商品"]
        #     data2 = dataForProducts[dataForProducts["场景"] != "S级商品"]
        #     dataForProducts = ""
        #     data1.drop(["当日平台价格"], axis=1, inplace=True)
        #     data1 = merge(data1,functionDic["城市商品"](files["城市商品"]))
        #     dataForProducts = data1.append(data2)
        #     data1 = ""
        #     data2 = ""
        wordsSensitive = ["品类","商品编码","品名","规格","当日成本","当日平台价格","朴朴优惠券","标签","购买数量","赠品数量","购买金额","折扣金额","友商编码","友商品项","友商规格","友商价","友商价执行门店数","友商优惠券","跟价","是否上线","是否计划下架","库存为0门店"]
        wordsAProducts = ["品类","3级编号","商品编码","品名","规格","当日成本","当日平台价格","朴朴优惠券","标签","购买数量","赠品数量","购买金额","折扣金额","友商编码","友商品项","友商规格","友商价","友商价执行门店数","友商优惠券","跟价","是否上线","是否计划下架","库存为0门店"]
        wordsOther = ["品类","场景","商品编码","品名","规格","当日成本","当日平台价格","朴朴优惠券","标签","购买数量","赠品数量","购买金额","折扣金额","友商编码","友商品项","友商规格","友商价","友商价执行门店数","友商优惠券","跟价","是否上线","是否计划下架","库存为0门店","3级编号"]
        # dataForProducts = merge(productsData,dataForProducts)
        outputData(writer,dropduplicates(dataForProducts[dataForProducts["场景"]=="敏感商品"][wordsSensitive]),"敏感商品")
        dataAproducts = dropduplicates(dataForProducts[(dataForProducts["场景"]=="A类") &(dataForProducts["库存合计"]>46) & (dataForProducts["是否上线"] == 1) & (dataForProducts["友商价"] > 0) ][wordsAProducts])
        outputData(writer,dataAproducts,"A类")
        
        outputData(writer,dropduplicates(dataForProducts[wordsOther]),"营销或其他（不删）")
        
        data1 = dataForProducts[dataForProducts["当日平台价格"].isnull()]
        data2 = dataForProducts[dataForProducts["当日平台价格"].notnull()]
        dataForProducts = ""
        data1 = data1[data1["友商编码"].isnull()]
        dataForProducts = data1.append(data2)
        data1 = ""
        data2 = ""
        dataForProducts = merge(productsData,dataForProducts)
        
        
        dataMarketing = dropduplicates(dataForProducts[(dataForProducts["库存合计"]>46) & (dataForProducts["是否上线"] == 1)][wordsOther])
        outputData(writer,dataMarketing,"营销")
    
    
    
    data1 = data[data["当日平台价格"].isnull()]
    data2 = data[data["当日平台价格"].notnull()]
    data = ""
    data1["友商价"] = ""
    data = data1.append(data2)
    data1 = ""
    data2 = ""
    
    
    # yHCashBack = ""
    # yHFlashSale = ""
    
    
    global yHCashBack
    global yHFlashSale
    # yHCashBack.to_frame().to_excel(writer,sheet_name = "非秒杀",index = False)
    # yHFlashSale.to_frame().to_excel(writer,sheet_name = "秒杀",index = False)
    YHCB = ["品类","场景","商品编码","品名","规格","当日成本","当日平台价格","朴朴优惠券","标签","购买数量","赠品数量","购买金额","折扣金额","友商编码","友商品项","友商规格","友商价","友商价执行门店数","友商优惠券","跟价"]
    dataForYH = data[(data["仓或门店"] == "仓")&(data["当日成本"].notnull())&(data["当日平台价格"].notnull())&(data["库存合计"]>46) & (data["是否上线"] == 1)]
    yHCashBack = pd.merge(yHCashBack,dataForYH,how="inner")
    yHCashBack["场景"] = "永辉折扣"
    yHCashBack = dropduplicates(yHCashBack[YHCB])
    outputData(writer,yHCashBack,"友商非秒杀")
    
    YHFS = ["品类","场景","友商编码","友商品项","友商规格","友商价","友商价执行门店数","商品编码","品名","当日平台价格","当日成本"]
    yHFlashSale = pd.merge(yHFlashSale,dataForYH,how="inner")
    yHFlashSale["场景"] = "限时抢购"
    outputData(writer,dropduplicates(yHFlashSale[YHFS]),"友商秒杀")
    
    
    
    
    wordsList = [i for i in wordsOrder if i in data.columns]
    outputData(writer,data[wordsList],"合并数据")
    print("开始输出数据")
    writer.save()



def filesCheck():

    filedic = {"朴朴": "",
                    "永辉商户产品": "",
                    # "城市商品": "",
                    "带优惠产品信息": "",
                    "statistics": "",
                    "商品库存盘点": "",
                    "清单": "",
                    "友商竞品": "",
                    day: ""
    }
    print("开始定位文件。")
    for i in files:
        if "朴朴" in i and i[0] != "~":
            if filedic["朴朴"] == "":
                filedic["朴朴"] = i
        elif "永辉商户产品" in i and i[0] != "~":
            if filedic["永辉商户产品"] == "":
                filedic["永辉商户产品"] = i
        # elif "城市商品" in i and i[0] != "~":
        #     if filedic["城市商品"] == "":
        #         filedic["城市商品"] = i
        elif "带优惠产品信息" in i and i[0] != "~":
            if filedic["带优惠产品信息"] == "":
                filedic["带优惠产品信息"] = i
        elif "statistics" in i and i[0] != "~":
            if filedic["statistics"] == "":
                filedic["statistics"] = i
        elif "商品库存盘点" in i and i[0] != "~":
            if filedic["商品库存盘点"] == "":
                filedic["商品库存盘点"] = i
        elif "清单" in i and i[0] != "~":
            if filedic["清单"] == "":
                filedic["清单"] = i
        elif "友商竞品" in i and i[0] != "~":
            if filedic["友商竞品"] == "":
                filedic["友商竞品"] = i
        elif day in i and i[0] != "~" and len(i)<18:
            if filedic[day] == "":
                filedic[day] = i
        else:
            pass
    print("各表获取结果")
    count = 0
    for k, v in filedic.items():
        if v:
            count += 1
            print("名字带有【"+k+"】的表:"+v)
        else:
            print("名字带有【" + k + "】的表: 没有该表")
    print("总共确定"+str(count)+"张表")
    print()
    flag = 0
    # if not filedic["友商竞品"] or not filedic["商品库存盘点"] or not filedic["statistics"] or not filedic["永辉商户产品"] or not filedic["朴朴"]:
    # 	print("五张张必要的表不齐全，请凑齐后重试:", "友商竞品,", "商品库存盘点,", "statistics,", "永辉商户产品,", "朴朴,")
    # elif count == 9:
    # 	print("""9表齐全，请选择制作内容:
    # 	直接回车或输入1后回车:制作营销表或A类或敏感商品，
    # 	输入2后回车：制作友商表，
    # 	输入3后回车：同时制作以上两张表
    # 	""")
    # 	while True:
    # 		answer = input()
    # 		if answer == "" or answer == "1":
    # 			flag = 1
    # 		elif answer == "2":
    # 			flag = 2
    # 		elif answer == "3":
    # 			flag = 3
    # 		elif answer == "q":
    # 			break
    # 		else:
    # 			print("输入有误，请重新尝试，输入q或点击右上角×退出")
    # 		if flag:
    # 			break
    # elif count == 8:
    # 	if not filedic["城市商品"]:
    # 		print("开始制作其他表，类型包括：A类、敏感商品等")
    # 		flag = 4
    # else:
    # 	print("开始制作友商表")
    # 	flag = 2
    dataProcessing(filedic, flag)

yHCashBack = ""
yHFlashSale = ""
day = datetime.datetime.now().strftime('%Y%m%d')
# day = "20210127"
path = os.getcwd() + "\\"
# path = r"D:\数据\2021\2月\0205数据\原始"
# os.chdir(path)
files = os.listdir(path)

filesCheck()


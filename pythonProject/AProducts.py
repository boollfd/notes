# -*-coding = utf-8 -*-
import datetime
import numpy as np
import pandas as pd
import re
import os

# 若需要，可改品类名
def categoryReplace(data,columnName):
    data[columnName] = data[columnName].replace({"宠物用品": "家居日用",
                                                    "纺织品": "家居日用",
                                                    "家庭用品": "家居日用",
                                                    "家用清洁用品": "家居日用",
                                                    "生活电器": "家居日用",
                                                    "文体休闲": "家居日用",
                                                    "鞋与配件": "家居日用",
                                                    "日用清洁": "家居日用",
                                                    "粮油米面": "粮油干货",
                                                    "调味干货": "粮油干货",
                                                    "常温奶": "奶制品",
                                                    "低温奶": "奶制品",
                                                    "个人清洁用品": "个护美妆",
                                                    "计生用品": "个护美妆",
                                                    "婴儿用品": "母婴类",
                                                    "婴儿玩具": "母婴类",
                                                    "婴童寝居": "母婴类",
                                                    "孕妇产品": "母婴类",
                                                    "婴儿食品": "母婴类"})
    
# 品类排序并新增相应维度的累计列,最好只传入必要字段
def categoryDataProcessing(data,targetColumn,sumOfCategory):
    data = data.sort_values(by=targetColumn, ascending=False)
    li = [0]+list(data[targetColumn])
    for i in range(1,len(li)):
        li[i] += li[i-1]
    li = li[1:]
    data[targetColumn+"累计"] = li
    li = [i/sumOfCategory for i in li]
    data[targetColumn+"累计占比"] = li
    data[targetColumn+"评级"] = ["A" if i<limits[0] else ( "C" if i>limits[1] else "B" ) for i in li]
    
    return data

def categoryLoop(data,targetColumn,categoryPvt):
    li = []
    for categery in categoryPvt.index:
        li.append(categoryDataProcessing(data[data["2级类别名"] == categery],targetColumn,categoryPvt.loc[categery,targetColumn]))
    return pd.concat(li)

def entryPoint(filenameBefore,filenameNow,targetColumns,finalType,writer):
    dataNow = pd.read_excel(filenameNow,usecols=["商品编号","商品名","2级类别名","销售数量","实际收入","实收毛利"])
    dataBefore = pd.read_excel(filenameBefore, usecols=["商品编号","商品名","2级类别名","销售数量","实际收入","实收毛利"])
    print("删除特选商品")
    dataNow = dataNow[dataNow["商品名"].str.contains("特选")==False]
    dataBefore = dataBefore[dataBefore["商品名"].str.contains("特选")==False]
    
    categoryPvtNow = pd.pivot_table(dataNow, index=["2级类别名"], values=targetColumns, aggfunc="sum")
    categoryPvtBefore = pd.pivot_table(dataBefore, index=["2级类别名"], values=targetColumns, aggfunc="sum")
    productsAList = pd.DataFrame(columns = ["商品编号","商品名","合并"])
    # storeBalanceData = storeBalanceData[storeBalanceData["库存合计"].isin(["--"])]
    for targetColumn in targetColumns:
        print("开始按"+targetColumn+"统计")
        resultNow = categoryLoop(dataNow,targetColumn,categoryPvtNow)
        resultBefore = categoryLoop(dataBefore,targetColumn,categoryPvtBefore)[["商品编号","销售数量","实际收入","实收毛利",targetColumn+"累计",targetColumn+"累计占比",targetColumn+"评级"]]
        resultBefore.columns = ["商品编号","去年销售数量","去年实际收入","去年实收毛利","去年"+targetColumn+"累计","去年"+targetColumn+"累计占比","去年"+targetColumn+"评级"]
        #.to_excel(writer, sheet_name=sheetname+targetColumn+"统计",index=False)
        
        resultNow = pd.merge(resultNow, resultBefore, how="left")
        resultNow["去年"+targetColumn+"评级"].fillna("",inplace = True)
        resultNow["合并"] = resultNow[targetColumn+"评级"] + resultNow["去年"+targetColumn+"评级"]
        resultNow = resultNow[["商品编号","商品名","2级类别名","销售数量","实际收入","实收毛利",targetColumn+"累计",targetColumn+"累计占比",targetColumn+"评级","去年"+targetColumn+"评级","合并","去年销售数量","去年实际收入","去年实收毛利","去年"+targetColumn+"累计","去年"+targetColumn+"累计占比"]]
        if targetColumn =="实收毛利":
            print("删除负毛利商品")
            resultNow = resultNow[resultNow["实收毛利累计占比"]>0]
            
        
        resultAList = resultNow[["商品编号","商品名","合并"]]
        
        resultAList = resultAList[resultAList["合并"].isin(finalType)]
        
        productsAList = productsAList.append(resultAList,ignore_index=True)
        resultNow.to_excel(writer, sheet_name="按"+targetColumn+"统计",index=False)
    productsAList = productsAList.drop_duplicates(["商品编号"])
    print("输出A类清单中")
    productsAList.to_excel(writer, sheet_name="A类清单列表",index=False)
    
print("开始定位文件")
path = os.getcwd()+"\\"
files = os.listdir(path)
filenameBefore = ""
filenameNow = ""
configFile = ""
targetColumns = ["销售数量","实际收入","实收毛利"]
finalType = ["AA","AB","AC"]
limits = [0.5,0.9]
for file in files:
    if file[0] != "~" and "今年" in file:
        filenameNow = path + file
    elif file[0] != "~" and "去年" in file:
        filenameBefore = path + file
    elif file[0] != "~" and "配置" in file:
        configFile = path + file
    else:
        pass
    
if filenameBefore == "" or filenameNow == "":
    input("最少需要2张表哦,一张带有'今年',一张带有'明年'，请检查后重试，按右上角X或输入任意内容回车结束。")
else:
    print("表名为"+filenameBefore+"和"+filenameNow)
    print("开始读取并处理数据")
    if configFile:
        limits = list(pd.read_excel(configFile, sheet_name = "配置").iloc[0,0:2])
        finalType = list(pd.read_excel(configFile, sheet_name = "配置").iloc[1,:])
        finalType = [i for i in finalType if i == i]
        print("ABC标准：",limits,"最终提取类别：",finalType)
# categoryReplace(dataNow,"2级类别名")
# categoryReplace(dataBefore,"2级类别名")
    writer = pd.ExcelWriter(path+"A类商品清单结果.xlsx")
    entryPoint(filenameBefore,filenameNow,targetColumns,finalType,writer)
    print("开始写入文件")
    writer.save()
print("结束")
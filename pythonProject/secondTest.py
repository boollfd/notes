# -*-coding = utf-8 -*-
import pandas as pd
import os
import re
import datetime
import numpy as np

# def sheetTransport(filename):
#     df = pd.DataFrame(columns=["活动类别", "编码", "活动区域", "日期"])
#     datas = pd.read_excel(filename, sheet_name=None)
#     for data in datas:
#         if len(datas[data].index)==0:
#             continue
#         print(datas[data].columns)
#         # print(datas[data].head())
#         datas[data].columns = datas[data].loc[0,:]
#         print(datas[data].columns)
#         if {"排期时间","上图商品名称","活动区域"} < set(datas[data].columns):
#             datas[data] = datas[data][["排期时间","上图商品名称","活动区域"]]
#             df = df.append(transferData(datas[data], df, data))
#     df.to_excel("促销排期整理.xlsx", index=False)
    
def sheetsTransport(filename, sheets,names):
    df = pd.DataFrame(columns=["活动类别", "编码", "活动区域", "日期"])
    for i in range(len(sheets)):
        data = pd.read_excel(filename, sheet_name=sheets[i], skiprows=1,usecols=["排期时间","上图商品名称","活动区域"],dtype = {"排期时间":np.str_})
        df = df.append(transferData(data, df, names[i]))
    df.drop_duplicates().to_excel(datetime.datetime.now().strftime('%Y%m%d')+"促销排期整理.xlsx", index=False)

def transferData(data, df, marketing):
    res = []
    for i in data.itertuples():
        print(i)
        ids = re.findall("\d{7}",str(i[2]))
        days=[]
        if "-" in str(i[1]):
            beginDate, endingDate = (i[1].split("-"))
            beginDate = beginDate.split(".")
            year = beginDate[-3] if len(beginDate) > 2 else datetime.datetime.now().year
            beginDate = datetime.date(int(year), int(beginDate[-2]), int(beginDate[-1]))
            endingDate = endingDate.split(".")
            year = endingDate[-3] if len(endingDate) > 2 else datetime.datetime.now().year
            month = endingDate[-2] if len(endingDate) > 1 else beginDate.month
            endingDate = datetime.date(int(year), int(month), int(endingDate[-1]))
            while beginDate <= endingDate:
                days.append(beginDate.strftime("%m-%d"))
                beginDate += datetime.timedelta(days=1)
        elif "." in str(i[1]):
            beginDate = i[1].split(".")
            year = beginDate[-3] if len(beginDate) > 2 else datetime.datetime.now().year
            beginDate = datetime.date(int(year), int(beginDate[-2]), int(beginDate[-1])).strftime("%m-%d")
            # days.append('{:.2f}'.format(float(i[1].replace("’",""))))
            days.append(beginDate)
        else:
            days.append(str(i[1]))
        areas = [j for j in str(i[3]) if j in "福厦广深武"]
        res += [[marketing, int(b), h, r] for b in ids for h in areas for r in days if b != None]
    return pd.DataFrame(res, columns=["活动类别", "编码", "活动区域", "日期"])
path = os.getcwd() + "\\"
files = os.listdir(path)
sheets = ["品牌墙", "分类页", "汇总（主题场景）","APP首页上图品汇总"]
names = ["品牌墙", "分类页", "主题场景", "banner"]
filename = ""
for i in files:
    if "促销排期" in i and i[0] != "~":
        filename = path+i
        break

# sheetTransport(path+filename)
sheetsTransport(filename, sheets,names)
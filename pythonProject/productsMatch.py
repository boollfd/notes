import pandas as pd
import numpy as np
import jieba
import re
from collections import Counter
from xlsxwriter.workbook import Workbook
import os
import datetime
def averageSize(s, li):
    print(s)
    averSize = ""
    if "±" in s:
        ss = s.split("±")
        if len(re.findall("\d+[\.]?\d*", ss[-2]))>0:
            size = re.findall("\d+[\.]?\d*", ss[-2])
            if len(ss) > 0:
                sizeList = ["±", size[-1]]
                if len(re.findall("\d+[\.]?\d*", ss[-1])) > 0:
                    sizeList.append(re.findall("\d+[\.]?\d*", ss[-1])[0])
            
                averSize = float(size[-1])
                li = deleteSize(li, sizeList)
    elif "*" in s:
        ss = s.split("*")
        le = re.findall("\d+[\.]?\d*", ss[-2])
        count = 0
        if len(le) > 0:
            le = le[-1]
            count += 1
        ri = re.findall("\d+[\.]?\d*", ss[-1])
        if len(ri) > 0:
            ri = ri[0]
            count += 1
        if count > 1:
            averSize = float(ri) * float(le)
            li = deleteSize(li, [ri, le, "*"])
        elif le:
            averSize = le
        else:
            pass
    elif "×" in s:
        ss = s.split("×")
        le = re.findall("\d+[\.]?\d*", ss[-2])
        count = 0
        if len(le) > 0:
            le = le[-1]
            count += 1
        ri = re.findall("\d+[\.]?\d*", ss[-1])
        if len(ri) > 0:
            ri = ri[0]
            count += 1
        if count > 1:
            averSize = float(ri) * float(le)
            li = deleteSize(li, [ri, le, "×"])
        elif le:
            averSize = le
        else:
            pass
    elif "+" in s:
        ss = s.split("+")
        size = re.findall("\d+[\.]?\d*", ss[-2])
        if len(size) > 0:
            sizeList = ["+", size[-1]]
            if len(re.findall("\d+[\.]?\d*", ss[-1]))>0:
                sizeList.append(re.findall("\d+[\.]?\d*", ss[-1])[0])
            averSize = float(size[-1])
            li = deleteSize(li, sizeList)
    elif "-" in s:
        ss = s.split("-")
        le = re.findall("\d+[\.]?\d*", ss[-2])
        count = 0
        if len(le) > 0:
            le = le[-1]
            count += 1
        ri = re.findall("\d+[\.]?\d*", ss[-1])
        if len(ri) > 0:
            ri = ri[0]
            count += 1
        if count > 1:
            averSize = (float(ri) + float(le)) / 2
            li = deleteSize(li, [ri, le, "-"])
    else:
        pass
    if not averSize:
        for i in ["g", "ml"]:
            le = re.findall("\d+[\.]?\d*"+i, s, flags=re.IGNORECASE)
            if len(le) > 0:
                averSize = float(re.findall("\d+[\.]?\d*", le[-1])[0])
                li = deleteSize(li, [le[-1]])
                break
        else:
            le = re.findall("\d+[\.]?\d*", s)
            if len(le) > 0:
                averSize = float(le[-1])
                li = deleteSize(li, [le[-1]])
            else:
                averSize = 0
    if "斤" in s:
        averSize = float(averSize) * 500
    elif re.findall('kg', s, flags=re.IGNORECASE):
        averSize = float(averSize) * 1000
    elif "两" in s:
        ss = re.findall("\d+[\.]?\d*两", s)
        if len(ss) > 0:
            averSize = float(re.findall("\d+[\.]?\d*", ss[-1])[0]) * 50
    else:
        pass
    deleteSize(li, ["斤", "kg", "两"])
    
    return float(averSize)


def deleteSize(li, sizeList):
    for i in sizeList:
        for j in li[::-1]:
            if i in j and j in li:
                ind = 50
                if j in li:
                    ind = li.index(j)
                li.remove(j)
                if ind < len(li):
                    if len(li[ind]) == 1:
                        li.pop(ind)
    return li


def deleteOthers(li):
    li = deleteSize(li, ['（', '）', '(', ')'])
    if "/" in li:
        ind = len(li) - 1 - li[::-1].index("/")
        if ind < len(li):
            li.pop(ind)
            if ind < len(li) and len(li[ind])<2:
                li.pop(ind)
    if "【" in li and "】" in li:
        for i in range(li.index("】"), li.index("【")-1,-1):
            li.pop(i)


def wordsReplace(s, words):
    for word in words:
        s = s.replace(str(word), "")
    return s


def managedata(data,config):
    data["分词包含单字"] = data["品名"].apply(lambda x: jieba.lcut(wordsReplace(x, config["删除词汇"]), cut_all=False))
    data["分词包含单字"].apply(lambda x: deleteOthers(x))
    data["平均规格"] = [averageSize(j, i) for i, j in zip(data["分词包含单字"], data["品名"])]
    # data["分词去除单字"] = [[j for j in i if len(j) > 1] for i in data["分词包含单字"]]
    # data["词数量"] = data["分词去除单字"].apply(lambda x: len(x))
    data["比对品名"] = ["".join(i) for i in data["分词包含单字"]]
    dic = {}
    for i in data.iterrows():
        for j in i[1]["分词包含单字"]:
            if ord(j[0]) < 256:
                i[1]["分词包含单字"].remove(j)
                continue
            if j in dic:
                dic[j].add(i[0])
            else:
                dic[j] = {i[0]}
    data["词数量"] = data["分词包含单字"].apply(lambda x: len(x))
    return dic


def running(filename):

    print("读取配置子表")
    config = pd.read_excel(filename, sheet_name="配置")
    ratio = [0.5,0.5]
    if "相同词汇占朴朴商品名的比例" in config.columns and config.loc[0,"相同词汇占朴朴商品名的比例"] == config.loc[0,"相同词汇占朴朴商品名的比例"]:
        ratio[0] = config.loc[0,"相同词汇占朴朴商品名的比例"]
    if "相同词汇占友商商品名的比例" in config.columns and config.loc[0,"相同词汇占友商商品名的比例"] == config.loc[0,"相同词汇占友商商品名的比例"]:
        ratio[0] = config.loc[0,"相同词汇占友商商品名的比例"]
    
    for word in config["往字典添加词汇"].tolist():
        if word == word:
            jieba.add_word(word,999999)
    print("读取朴朴子表并处理品名和平均规格")
    dataPUPU = pd.read_excel(filename, sheet_name="朴朴")#,dtype={"编码":"str"}
    dicPUPU = managedata(dataPUPU,config)
    lengthPUPU = dataPUPU.shape[0]
    listPUPU = [list() for i in range(lengthPUPU)]

    print("读取友商子表并处理品名和平均规格")
    dataYS = pd.read_excel(filename, sheet_name="友商")
    dicYS = managedata(dataYS,config)
    # dataYS = dataYS.append(dataYS.loc[[1,2,3,4],:])
    dataYS.columns = ["友商"+i for i in dataYS.columns]

    print("求朴朴和友商所有关键词的交集")
    keywords = set(dicPUPU).intersection(set(dicYS))
    for keyword in keywords:
        for j in dicPUPU[keyword]:
            listPUPU[j] += list(dicYS[keyword])
    dataPUPU["匹配序号"]=listPUPU
    print("根据有交集的关键词得出匹配列表，相同词汇占朴朴商品名的比例少于"+str(ratio[0])+"占友商商品名的比例少于"+str(ratio[1])+"的数据，并计算是否完全匹配。")
    dataCompair = pd.DataFrame(columns=list(dataPUPU.columns)+list(dataYS.columns))

    for i in range(lengthPUPU-1, -1, -1):
        count = Counter(listPUPU[i])
        most = 0
        for j in count:
            if count[j] > most:
                most = count[j]
            if count[j] < most:
                continue
            if most/dataPUPU.loc[i, "词数量"] < ratio[0] or most/dataYS.loc[j, "友商词数量"] < ratio[1]:
                continue
            dataCompair = dataCompair.append(dataPUPU.loc[i, :].append(dataYS.loc[j, :]), ignore_index=True)

    dataCompair["匹配词集合"] = [set(i).intersection(set(j)) for i, j in zip(dataCompair["分词包含单字"], dataCompair["友商分词包含单字"])]
    dataCompair["是否完全匹配"] = ["是"  if len(i)==j else "否" for i,j in zip(dataCompair["匹配词集合"],dataCompair["词数量"]) ]
    dataCompair["匹配词"] = [",".join(i) for i in dataCompair["匹配词集合"]]


    print("计算朴朴和友商平均规格的差距")
    # reshCategory = config["朴朴生鲜类别"].tolist()
    # relexSize = int(config["规格可放松多少"][0])
    # finalName = config["友商名称"][0]
    day = datetime.datetime.now().strftime('%Y%m%d')


    dataCompair["规格差距"] = 0
    dataCompair = dataCompair.reset_index()
    for i in range(dataCompair.shape[0]-1, -1, -1):
        if dataCompair.loc[i, "平均规格"] != 0 and dataCompair.loc[i, "友商平均规格"] != 0:
            dataCompair.loc[i, "规格差距"] = abs(dataCompair.loc[i, "平均规格"]-dataCompair.loc[i, "友商平均规格"])
            # if dataCompair.loc[i, "规格差距"] > 0 and dataCompair.loc[i, "品类"] not in reshCategory:
            #     dataCompair = dataCompair.drop(i)
            #     continue
        else:
            dataCompair.loc[i, "规格差距"] = -1
        # if len(dataCompair.loc[i, "匹配词集合"]) == 1 and len(list(dataCompair.loc[i, "匹配词集合"])[0]) == 1:
        #     dataCompair = dataCompair.drop(i)

    dataCompair = dataCompair.sort_values(by="规格差距")
    # dataCompair = dataCompair[dataCompair["规格差距"] < relexSize]

    print("删除重复项")
    dataCompair = dataCompair.drop_duplicates(["编码", "友商编码"])

    writer = pd.ExcelWriter("源数据-可不看.xlsx")
    dataCompair.to_excel(writer,  sheet_name="比对结果")
    dataPUPU.to_excel(writer,  sheet_name="朴朴")
    dataYS.to_excel(writer,  sheet_name="友商")
    writer.save()

    print("格式化输出结果中。")
    data_result = Workbook("比对结果-"+filename.replace("输入数据",""))
    worksheet = data_result.add_worksheet()
    red = data_result.add_format({
        'color': 'red',
    })
    form = data_result.add_format({'align': 'center',
                                'valign': 'vcenter',
                                'border': 1
                                })
    wordslist = ["比对品名", "友商比对品名", "是否完全匹配", "规格差距", "品名", "友商品名", "编码", "平均规格", "友商编码", "友商平均规格"]
    if "品类" in dataCompair:
        wordslist.append("品类")
    dataCompair = dataCompair.reset_index()[wordslist]
    cols = len(dataCompair.columns)
    worksheet.write_row("A1", list(dataCompair.columns), cell_format=form)
    for i in range(dataCompair.shape[0]):
        params = []
        for j in dataCompair.loc[i, "比对品名"]:
            if j in dataCompair.loc[i, "友商比对品名"]:
                params.extend((red, j))
            else:
                params.append(j)
        worksheet.write_rich_string(i+1, 0, *params, form)
        params = []
        for j in dataCompair.loc[i, "友商比对品名"]:
            if j in dataCompair.loc[i, "比对品名"]:
                params.extend((red, j))
            else:
                params.append(j)
        worksheet.write_rich_string(i+1, 1, *params, form)
        worksheet.write_row("C"+str(i+2), list(dataCompair.iloc[i, 2:]), cell_format=form)
        worksheet.set_row(i+1, 30)
    worksheet.set_column(0, 1, 25)
    worksheet.set_column(2,3, 14)
    worksheet.set_column(4, 5, 30)
    worksheet.set_column("G:K", 14)
    data_result.close()



# pd.DataFrame(columns=["朴朴编码","友商编码","匹配词","朴朴词数","友商词数","","","",""])
# print(data.loc[[1,2,3,4],:])
path = os.getcwd()+"\\"
files = os.listdir(path)
filename = []
print("开始加载字典")
jieba.set_dictionary("dict.txt")
jieba.initialize()

for file in files:
    if "输入数据" in file:
        filename.append(file)

for file in filename:
    print()
    print()
    print("开始读取"+file)
    running(file)
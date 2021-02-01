import pandas as pd
import os
import calendar

def mergefile(filepath,mergefile):
    iris_concat = pd.DataFrame()
    file = filepath
    iris1 = pd.ExcelFile(file)  # 读取Excel文件
    iris1Df = iris1.parse('172.16.X.X')
    iris2Df = iris1.parse('172.17.X.X')
    iris3Df = iris1.parse('(灾备）172.18.X.X')
    frames = [iris1Df,iris2Df,iris3Df]
    iris_concat = pd.concat(frames)#文件拼接
    iris_concat.to_excel(mergefile)  # 数据保存路径

def comparefile(filename,mergefilename):
    file = filename
    mergefile = mergefilename

    origin1 = pd.ExcelFile(mergefile)#读取IP源文件
    originDf = origin1.parse('Sheet1')
    origin1 = pd.read_excel(mergefile)
    searchfile1 = pd.read_excel(file)  # 读取要查询的Excel文件
    #searchfileDf = searchfile1.parse(sheet)
    #searchfile1 = pd.read_excel(file)
    rowsnumber1 = len(origin1)
    rowsnumber2 = len(searchfile1)
    #count = 0
    for k in range (rowsnumber1):#对IP原表中的每一个IP地址，遍历查询表中的IP，如果相同，就把主机名赋值
        for j in range (rowsnumber2):
            if "生产系统列表" in file:#如果IP地址相同，则把主机名称赋值给左侧列表
                #print (origin1.iloc[k,2])
                if origin1.iloc[k,2] == searchfile1.iloc[j,8] and searchfile1.iloc[j,8]!=None:
                    print(searchfile1.iloc[j,7])
                    origin1.iloc[k,3] = searchfile1.iloc[j,7]
                    continue
                else:
                    j = j+1
            elif "宁桥路灾备" in file:
                if origin1.iloc[k,2] == searchfile1.iloc[j,7] and searchfile1.iloc[j,7]!=None:
                    origin1.iloc[k,3] = searchfile1.iloc[j,6]
                    continue
                else:
                    j = j+1
    origin1.loc[:, ~origin1.columns.str.contains("^Unnamed")]  # 这两行代码主要作用是避免pandas在处理Excel时出现混乱的列
    origin1.to_excel(mergefile, index=False)

if __name__ == '__main__':
    if os.path.isfile('E:/python学习/存活IP合并.xlsx'):
        comparefile('E:/python学习/宁桥路灾备.xlsx', 'E:/python学习/存活IP合并.xlsx')
    else:
        mergefile('E:/python学习/存活IP-20201027.xlsx','E:/python学习/存活IP合并.xlsx')
        comparefile('E:/python学习/宁桥路灾备.xlsx','E:/python学习/存活IP合并.xlsx')

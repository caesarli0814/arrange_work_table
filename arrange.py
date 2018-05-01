# -*- coding: utf-8 -*-
"""
Created on Sun Apr 29 15:08:31 2018

@author: 55001
"""
import numpy as np
import pandas as pd
import tkinter
from tkinter import ttk
import cvxopt
import cvxopt.glpk

#read_xls函数用于读取excel文件，并进行预处理，将无空闲时间的人删掉
def read_xls(path):
    Excel_data = pd.read_excel(path, index = 0)
    table = Excel_data.iloc[0:, 1:]
    for i in range(table.shape[0]):
        if(np.sum(table.loc[i, :]) == 0):
            Excel_data.drop(Excel_data.index[i], inplace = True)
    table = Excel_data.iloc[0:, 1:]
    table = table.values
    name = Excel_data.name
    name = name.values
    return table, name

#creat_A函数用于构造约束矩阵
def create_A(data, max_work, max_people, min_people):
    size = data.shape
    n = size[0] * size[1]
    f = np.full((1, n), 1)
    A = np.zeros((size[1]*2 + 2*size[0] + 1, n))
    for i in range(size[1]):
        A[i, i*size[0] : (i+1)*size[0]] = -data[:,i]
        A[i + size[1], i*size[0] : (i+1)*size[0]] = 1
    for i in range(size[0]):
        for j in range(size[1]):
            A[i + size[1]*2, j * size[0]+i] = -1
            A[i + size[1]*2 + size[0], j * size[0] + i] = 1
    A[size[1]*2 + 2*size[0], :] = np.ones((1, n)) - np.reshape(data.T, (1, n))
    b = np.zeros((size[0] * 2 + size[1]*2 + 1, 1))
    b[0:size[1], 0] = -min_people
    b[size[1]:size[1]*2, 0] = max_people
    b[size[1]*2:size[0]+size[1]*2, 0] = -1
    b[size[0] + size[1]*2:size[0] * 2 + size[1]*2, 0] = max_work
    b[size[0] * 2 + size[1]*2, 0] = 1
    return f, A, b

#solve函数用于求解问题答案
def solve(f, A, b, n):
    cvxopt.glpk.options['msg_lev'] = 'GLP_MSG_OFF'
    f = cvxopt.matrix(f, tc = 'd')
    A = cvxopt.matrix(A, tc = 'd')
    b = cvxopt.matrix(b, tc = 'd')
    binVars = range(n)
    (status, x) = cvxopt.glpk.ilp(f.T, G = A, h = b, I=set(binVars), B=set(binVars))
    f = np.matrix(f)
    A = np.matrix(A)
    b = np.matrix(b)
    if(x is None):
        return None, 0
    else:
        x = np.matrix(x)
        return x, f.dot(x)

#change函数用于更改解，由于解不唯一，不满意可以更改
def change(f, A, b, X, fval, n):
    A = np.vstack((A, X.T))
    b = np.vstack((b, fval))
    s, fval = solve(f, A, b, n)
    return A, b, s, fval

#handle)list函数将结果转化为人名组合
def handle_list(x, name):
    p = [''] * (x.size // name.size)
    people = name.size
    for i in range(x.size):
        if(x[i] == 1):
            if(p[i//people] != ''):
                p[i//people] = p[i//people] + ' / '
            p[i//people] = p[i//people] + name[i % people]
    return p

#ask_choice函数用于弹出窗口，并可以更改结果，目前还未完成
def ask_choice(table, name):
    win = tkinter.Tk()
    tree = ttk.Treeview(win)
    tree["columns"] = ("a", "b", "c", 'd', 'e', 'f', 'g')
    tree.column("a", width = 150, anchor ="center")
    tree.column("b", width = 150, anchor ="center")
    tree.column("c", width = 150, anchor ="center")
    tree.column("d", width = 150, anchor ="center")
    tree.column("e", width = 150, anchor ="center")
    tree.column("f", width = 150, anchor ="center")
    tree.column("g", width = 150, anchor ="center")
    tree.heading("a",text="星期一")
    tree.heading("b",text="星期二")
    tree.heading("c",text="星期三")
    tree.heading("d",text="星期四")
    tree.heading("e",text="星期五")
    tree.heading("f",text="星期六")
    tree.heading("g",text="星期天")
    t = handle_list(table, name)
    tree.insert("",0,text="上午1-2节" ,values = [t[0], t[4], t[8], t[12], t[16], t[20], t[24]])  
    tree.insert("",1,text="上午3-4节" ,values = [t[1], t[5], t[9], t[13], t[17], t[21], t[25]])  
    tree.insert("",2,text="下午5-6节" ,values = [t[2], t[6], t[10], t[14], t[18], t[22], t[26]])  
    tree.insert("",3,text="下午7-8节" ,values = [t[3], t[7], t[11], t[15], t[19], t[23], t[27]]) 
    tree.pack()
 #   tkinter.Button(win, text = '满意', width = 200).pack(side = "bottom", expand = "yes")
 #   tkinter.Button(win, text = '不满意，重做', width = 200).pack(side = "bottom", expand = "yes")
    win.mainloop()
    return
def cal_work(table, p_num):
    ck = [0] * p_num
    for i in range(table.size):
        if (table[i] ==1):
            t = i % p_num
            ck[t] = ck[t] + 1
    return ck

#output函数用于输出结果到excel
def output(table, name, path):
    writer = pd.ExcelWriter(path)
    ck = cal_work(table, name.size)
    ck = np.matrix(ck)
    df0 = pd.DataFrame(ck.T, columns =['工作次数'], index = name)
    col = ['星期一', '星期二', '星期三', '星期四', '星期五', '星期六', '星期天']
    ind = ['上午1-2节', "上午3-4节", "下午5-6节", "下午7-8节"]
    df = pd.DataFrame(columns = col, index = ind)
    t = handle_list(table, name)
    for i in range(7):
        for j in range(4):
            df.loc[ind[j], col[i]] = t[i*4+j]
    df.to_excel(writer, sheet_name = '工作安排表')
    df0.to_excel(writer, sheet_name = '工作时间统计')
    writer.save()
    return

def main():
    path = input('请输入文件地址：（例如E://工作簿2.xlsx）')
    data, name = read_xls(path)
    size = data.shape
    max_work = int(input('每人最多几班：'))
    max_people = int(input('每班最多几人：'))
    min_people = int(input('每班至少几人：'))
    n = size[0] * size[1]
    f, A, b = create_A(data, max_work, max_people, min_people)
    cvxopt.glpk.options['msg_lev'] = 'GLP_MSG_OFF'
    X ,fval= solve(f, A, b, n)
    if(X is None):
        print('无解')
        return 
    output_path = input('请输入输出地址：（例如E://result1.xlsx）')
    ask_choice(X, name)
    output(X, name, output_path)
'''
    此段代码可用于寻找解空间方差最小的解，也就是排班比较均匀的，但解空间较大，算力有限，待完善
    X1 = X
    count = 0
    while(X1 is not None):
        count = count + 1
        A, b, X1, fval = change(f, A, b, X1, fval - 1, n)
        if(np.var(X1) < np.var(X)):
            X = X1
'''
'''
    此段代码用于更改选择，但解空间较大，且窗口函数还在编写中，暂时闲置
       ask_choice(X, name)
       if (X is None):
           print('没有其他方案了')
''' 
main()

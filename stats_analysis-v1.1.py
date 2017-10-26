from tkinter import *
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import * 
import xlrd
import pandas as pd
import numpy as np
from pyecharts import Radar
from pyecharts import Page
from pyecharts import Line
from pyecharts import WordCloud
from pyecharts import EffectScatter
from pyecharts import Scatter
from pyecharts import Overlap
import statsmodels.api as sm
import matplotlib.pyplot as plt
from PIL import Image, ImageTk 
from PIL import Image, ImageDraw, ImageFont
import pygame
import webbrowser

#选择文件函数
def selectPath():
    #申明全局变量
    global path1
    #获取路径
    path1 = askopenfilename(filetypes=[('All files','*.*'),('xlsx', '*.xlsx'),('xls','*.xls')])

    path.set(path1)
    return path1

#显示客户姓名
def Opendata():
    f = pd.read_excel(path1)
    print(f)
    Khxm = f.drop_duplicates(['客户姓名'])
    Khxm_khxm = Khxm['客户姓名'].tolist()
    Khxm1.set(Khxm_khxm)
#选择客户名称函数
def Chosekhxm():
    global a
    a = xm.get()
    getxm.set(xm.get())
    print(a)

def xianshixiangmu():
    m = pd.read_excel(path1)
    m1 = m[m['客户姓名'] == a]
    m2 = m1.drop_duplicates(['项目名称'])
    m3 = m2['项目名称'].tolist()
    m4.set(m3)

#选择客户项目函数
def Chosekhxxm():
    global b
    b = xxm.get()
    getxxm.set(xxm.get())
    print(b)


#站位函数
def xianxinghuigui():
    huigui_f = pd.read_excel(path1)
    print(huigui_f.head())
    Indv = huigui_f['状态打分']
    Depv = huigui_f['客户成绩']
    Indv=sm.add_constant(Indv) 
    huigui=sm.OLS(Depv,Indv)
    huigui=huigui.fit()

    EEE = str(huigui.summary())
    print(huigui.summary())
    print(EEE)

    space = 5

    # # PIL模块中，确定写入到图片中的文本字体
    # font = ImageFont.truetype('Arial.ttf', 15, encoding='utf-8')
    # # Image模块创建一个图片对象
    # im = Image.new('RGB',(800, 600),(255,255,255,255))
    # # ImageDraw向图片中进行操作，写入文字或者插入线条都可以
    # draw = ImageDraw.Draw(im)

    # draw.text((50,10),text = EEE,font=font, fill=(0,0,0))
    # im.show()
    # im.save('12345.PNG', "PNG")
    # PIL模块中，确定写入到图片中的文本字体
    font = ImageFont.truetype('hyqh.ttf', 18, encoding='utf-8')
    # Image模块创建一个图片对象
    im = Image.new('RGB',(10, 10),(255,255,255,255))
    # ImageDraw向图片中进行操作，写入文字或者插入线条都可以
    draw = ImageDraw.Draw(im, "RGB")
    # 根据插入图片中的文字内容和字体信息，来确定图片的最终大小
    img_size = draw.multiline_textsize(EEE, font=font)
    # 图片初始化的大小为10-10，现在根据图片内容要重新设置图片的大小
    im_new = im.resize((img_size[0]+space*2, img_size[1]+space*2))
    del draw
    del im
    draw = ImageDraw.Draw(im_new, 'RGB')
    # 批量写入到图片中，这里的multiline_text会自动识别换行符
    draw.multiline_text((space,space), text = EEE,  fill=(0,0,0), font=font)
    im_new.save('OLS.png', "png")

    page1 =Page(page_title='OLS')

    sc = Scatter('OLS Regression Result',width = 800,height = 600)
    sc_v1, sc_v2 = sc.draw("OLS.png")
    sc.add("summary", sc_v1, sc_v2, label_color=["#000"],symbol_size = 1, is_xaxis_show = False,is_yaxis_show = False,is_legend_show = 0)
    page1.add(sc)


    Indv_prime = np.linspace(Indv['状态打分'].min(),Indv['状态打分'].max(),100)[:,np.newaxis]

    BBB = []
    for BB in Indv_prime.tolist():
        BB = round(BB[0],2)
        BBB.append(BB)

    #print(BBB)

    Indv_prime=sm.add_constant(Indv_prime)

    Depv_hat=huigui.predict(Indv_prime)
    CCC = Depv_hat.tolist()


    line1 = Line('Linear regression curve')
    line1.add('',BBB,CCC,xaxis_name = '状态打分', yaxis_name = '客户预测值', yaxis_interval=20)

    page1.add(line1)
    page1.render('线性回归分析结果报告.html')



def open_statanalysis_resuilt():
    webbrowser.open_new_tab('简单统计分析结果报告.html') 


def open_ols_resuilt():
    webbrowser.open_new_tab('线性回归分析结果报告.html') 







    # print(Depv_hat)
    # plt.xlabel("zhuangtaidafen")
    # plt.ylabel("chengji")
    # plt.scatter(Indv['状态打分'], Depv, alpha=0.3)
    # plt.plot(Indv_prime[:,1], Depv_hat, 'r', alpha=0.9) 
    # plt.title('Linear regression curve')

    # plt.show()


#def pic_1():
#    pic_11 = Toplevel()
#    pic1 = Image.open(tui1)
#    pic11 = ImageTk.PhotoImage(pic1)
#    canvas1 = Canvas(top1, width = image.width*2 ,height = image.height*2, bg = 'white')
#    canvas1.create_image(0,0,image = img,anchor="nw")
#    canvas1.create_image(image.width,0,image = img,anchor="nw")
#    canvas1.pack()   
#    top1.mainloop()


def jinjiefenxi():
    pass

def yinzifenxi():
    pass

def statanalysis():
    ###----------------数据准备--------------------------
    #打开数据
    h = pd.read_excel(path1)
    #找出所选择的客户数据
    i = h[h['客户姓名'] == a]
    print(i)
    ####---------------------!!!!------------
    i0 = i['项目名称'].value_counts()
    print(i0)
    i00 = i0.to_frame()



    i00['项目名称'] = i00.index
    i00 = i00.reset_index(drop= True)
    i000 = i['项目名称'].value_counts().tolist()
    i00['训练次数'] = i000

    print(i0)
    print(i00)

    i2 = i['课程类型'].tolist()[0]
    page = Page(page_title=a)


    ##---------------radar图数据准备----------------------
    #对客户数据进行分组求和
    j = i.groupby(by=['项目名称'])['客户成绩'].mean()
    #查看j的类型
    #print(type(j))
    #把series转换为frame
    k = j.to_frame() 
    #取出K中索引，生成新一列，去掉k中索引
    k['项目名称'] = k.index
    k = k.reset_index(drop= True)
    print(k)

    l = i.groupby(by=['项目名称'])['客户成绩'].max()
    n = l.to_frame()
    n['项目名称'] = n.index
    n = n.reset_index(drop = True)

    ##--------------------radar图数据准备--------------------

    ##-------------------------合并dataframe-----------------

    diantu = pd.merge(i00,k,how = 'inner')
    print(diantu)
    ##-------------------------合并dataframe-----------------

    ##-------------------- line数据准备-----------------
    Dgkhxm = i.drop_duplicates(['项目名称'])
    Dgkhxm_list =Dgkhxm['项目名称'].tolist()
    #print(Dgkhxm)
    #print(Dgkhxm_list)
    Dgkhxm_len = len(Dgkhxm_list)
    #print(Dgkhxm_len)
    i1 = i[i['项目名称'] == b]
    #print(i1)
    ##-------------------line数据准------------------------

    ###----------------数据准备-----------------------------------------
    ##---------------------词云----------------------------------

    name = [i2,a]
    value = [55,50]
    wordcloud = WordCloud('客户基本信息',width=600, height=320)
    wordcloud.add("", name, value,word_size_range=[20, 25],shape='diamond')

    page.add(wordcloud)
    ##------------------------词云-----------------------------------

    #-----------------radar图--------------------------------
    schema = []
    for aa in k['项目名称'].tolist():
        aa = (aa,100)
        schema.append(aa)

    v_pingjun = []
    v_pingjun.append(k['客户成绩'].tolist())

    v_zuida =[]
    v_zuida.append(n['客户成绩'].tolist())

    #print(schema)
    #print(v_pingjun)
    #print(v_zuida)


    radar = Radar('客户成绩观星评测')
    radar.config(schema)
    radar.add("平均成绩",v_pingjun,is_splitline=True, is_axisline_show=True)
    radar.add("最好成绩", v_zuida, label_color=["#4e79a7"], is_area_show=False)
    page.add(radar)
    #------------------radar图-------------------------------------

    ##------------------------EffectScatter------------------------
    cishu = diantu['训练次数'].tolist()
    chengji = diantu['客户成绩'].tolist()
    es = EffectScatter("训练次数&成绩")
    es.add("effectScatter", cishu, chengji,xaxis_name = '训练次数',yaxis_name = '客户成绩')
    page.add(es)




    ##------------------------EffectScatter------------------------  

    ##---------------------line--------------------------------

    attr = i1['课程时间'].tolist()

    v_zoushi = i1['客户成绩'].tolist()
    line = Line(b)

    line.add(b, attr, v_zoushi, mark_line=["max", "average"])

    page.add(line)

    page.render('简单统计分析结果报告.html')
    ##----------------------line------------------------------------





#初始化窗口
root = Tk()
#设置窗口大小
root.geometry('620x600+300+100')
#程序名称
root.title('数据分析')
path = StringVar()

Khxm1 = StringVar()

xm = StringVar()

getxm = StringVar()

m4 =StringVar()

xxm = StringVar()

getxxm = StringVar()


Label(root,text = "                  ").grid(row = 0, column = 3)
Label(root,text = "   1:数据选择").grid(row = 1, column = 0)


Label(root,text = "数据路径:", fg = '#CD9B1D').grid(row = 2, column = 1)
Entry(root, textvariable = path, foreground = '#CD9B1D').grid(row = 2, column = 2)
Button(root, text = "选择数据", command = selectPath, fg = '#CD9B1D').grid(row = 2, column = 4)

Label(root, text = "训练客户:", fg = '#CD8C95').grid(row = 3, column = 1)
Button(root, text = "显示名称", command = Opendata, fg = '#CD8C95').grid(row = 3, column = 4)
Entry(root, textvariable = Khxm1, state = 'disabled',foreground = '#CD8C95').grid(row = 3, column = 2)

Label(root,text = "输入客户:", fg = '#CD7054').grid(row = 4, column = 1)
Entry(root, textvariable = xm, foreground = '#CD7054').grid(row = 4, column = 2)
Button(root, text = "确认姓名", command = Chosekhxm, fg = '#CD7054').grid(row = 4, column = 4)

Label(root,text = "您输入的客户是:",fg = '#CD69C9').grid(row = 5, column = 1)
Label(root,textvariable = getxm,fg = '#CD69C9', font=("黑体", 20, "bold")).grid(row = 5, column = 2)


Label(root,text = "客户项目:", fg = '#BDB76B').grid(row = 6, column = 1)
Button(root, text = "显示项目", command = xianshixiangmu, fg = '#BDB76B').grid(row = 6, column = 4)
Entry(root, textvariable = m4, state = 'disabled', foreground = '#BDB76B').grid(row = 6, column = 2)


Label(root,text = "输入项目:", fg = '#A4D3EE').grid(row = 7, column = 1)
Entry(root, textvariable = xxm, foreground = '#A4D3EE').grid(row = 7, column = 2)
Button(root, text = "确认项目", command = Chosekhxxm, fg = '#A4D3EE').grid(row = 7, column = 4)

Label(root,text = "您输入的项目是:", fg = '#9F79EE').grid(row = 8, column = 1)
Label(root,textvariable = getxxm, fg ='#9F79EE', font=("黑体", 20, "bold")).grid(row = 8, column = 2)


Label(root,text = "   2:数据分析").grid(row = 9, column = 0)

Button(root, text = "简单统计分析", command = statanalysis,height= 2,width =12, bg = '#FA8072',bd =4).grid(row = 10, column = 1,)
Button(root, text = "统计结果", command = open_statanalysis_resuilt,height= 2,width =12, bg = '#66FFFF',bd =4).grid(row = 10, column = 2)
Label(root,text = "").grid(row = 10, column = 0)

Button(root, text = "进阶分析", command = jinjiefenxi,height= 2,width =12, bg = '#FA8072',bd =4).grid(row = 10, column = 3,columnspan=2)
Button(root, text = "进阶结果", command = jinjiefenxi,height= 2,width =12, bg = '#66FFFF',bd =4).grid(row = 10, column = 5,columnspan=2)
Label(root,text = "").grid(row = 10, column = 0)

Label(root,text = "").grid(row = 11, column = 0)

Button(root, text = "线性回归", command = xianxinghuigui,height= 2,width =12, bg = '#FA8072',bd =4).grid(row = 12, column = 1)
Button(root, text = "线性结果", command = open_ols_resuilt,height= 2,width =12, bg = '#66FFFF',bd =4).grid(row = 12, column = 2)

Button(root, text = "因子分析", command = yinzifenxi,height= 2,width =12, bg = '#FA8072',bd =4).grid(row = 12, column = 3,columnspan=2)
Button(root, text = "因子结果", command = yinzifenxi,height= 2,width =12, bg = '#66FFFF',bd =4).grid(row = 12, column = 5,columnspan=2)


Label(root,text = "").grid(row = 13, column = 0)
Label(root,text = "----------------------------------我是分割线--------------------------------------").grid(row = 14, column = 0,columnspan=8)
Label(root,text = "欢迎您使用本软件，暂时只有简单统计分析和线性回归可用", fg ='#8B2500').grid(row = 15, column = 0,columnspan=8)
Label(root,text = "进阶分析、因子分析将在后续版本开放", fg ='#8B2500').grid(row = 16, column = 0,columnspan=8)

Label(root,text = "").grid(row = 17, column = 0)
Label(root,text = "").grid(row = 18, column = 0)
Button(root, text = "退出程序", command = root.quit, foreground = '#EE00EE').grid(row = 19, column = 6)

root.mainloop()

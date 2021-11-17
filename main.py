#!/usr/bin/env python
import PySimpleGUI as sg
import pandas as pd

# import os
# os.startfile('D:/yun/纸箱管理/3.doc')  #调用打印机打印文件
#sg.theme('Dark Red')
sg.theme('DarkGreen5')
#right_click_menu = ['', [ '复制订单','删除订单']]
filename='./订单.csv'
fahuodan='./发货单.csv'

header_list=['订单编号','订单日期','公司','规格','数量','状态'] #订单表格表头
fahuodanheader_list=['发货单号','发货日期','公司'] #订单表格表头

dingdanshujulx={'订单编号' : str,'订单日期':str} #订单数据类型

fahuodandanshujulx={'发货单号' : str,'发货日期':str} #订单数据类型
def getdata(filename,header_list,dingdanshujulx):
    #numcol=6#只显示到第8列
    #df = pd.read_csv(filename,header=None)    
    #header_list = df.iloc[0].tolist()[0:numcol]
    #print(header_list)
   # data = df.iloc[1:,0:numcol].values.tolist()
    #print(data)
    df = pd.read_csv(filename,dtype = dingdanshujulx)     
    data=df.loc[:,header_list].values.tolist()
    data.reverse()#数据倒叙，最新的在前
    return data
    
data = getdata(filename,header_list,dingdanshujulx)
fahuodandata = getdata(fahuodan,fahuodanheader_list,fahuodandanshujulx)
df = pd.read_csv(filename,dtype = dingdanshujulx) #打开数据文件
df.set_index('订单编号',inplace=True)    
dffahuodan = pd.read_csv(fahuodan,dtype = {'出货单号': str,'订单':str})
dffahuodan.set_index('发货单号',inplace=True)
dingdanbuttons=[sg.Button('新增订单',font=('MSYH',12),key = '-xinzdd-'), 
            sg.Button('修改订单',font=('MSYH', 12),key = '-xiugdd-'), 
            sg.Button('复制订单',font=('MSYH', 12),key = '-fuzhidd-') , 
            sg.Button('删除订单',font=('MSYH', 12),key = '-shanchudd-') ,           
            sg.Button('退出',font=('MSYH', 12),key = '-dingdantuichu-')
            ]
fahuodanbuttons=[
            sg.Button('新增发货单',font=('MSYH', 12),key = '-xinzfhd-'),
            sg.Button('修改发货单',font=('MSYH', 12),key = '-xiugfhd-'),
            sg.Button('查看发货单',font=('MSYH', 12),key = '-chakfhd-'),
            sg.Button('退出',font=('MSYH', 12),key = '-fahuodantuichu-')
            ]
duizhangdanbuttons=[
            sg.Button('新增对账单',font=('MSYH', 12),key = '-xinzdzd-'),
            sg.Button('修改对账单',font=('MSYH', 12),key = '-xiugdzd-'),
            sg.Button('查看对账单',font=('MSYH', 12),key = '-chakdzd-'),
            sg.Button('退出',font=('MSYH', 12),key = '-duizhangdantuichu-')
            ]
dingdankeys_list={'-dingdannumber-':'订单编号',
              '-date-':'订单日期',
              '-com-':'公司',
              '-sty-':'规格',
              '-value-':'单价',
              '-num-':'数量',
              '-valall-':'总价',
              '-zhuangtai-':'状态',
              '-length-':'长',
              '-width-':'宽',
              '-height-':'高'
              }
dingdantable= sg.Frame('', [[sg.Table(values=data,
                    headings=header_list,
                    max_col_width=500,                    
                    col_widths=[10,7,14,8,4,8,8],
                    auto_size_columns=False,
                    justification='center',
                    #display_row_numbers=True,
                    # alternating_row_color='lightblue',
                    num_rows=min(len(data)+50, 100),
                    key='-TABLE-',
                    selected_row_colors='red on yellow',                    
                    #right_click_menu=right_click_menu,
                    font=('微软雅黑', 12),
                    #enable_events = True,
                    enable_click_events = True,
                    vertical_scroll_only =False,
                    display_row_numbers=True
                    )]],size=(650,750)
                  )	
#sg.InputCombo(['Combobox 1','Combobox 2'],size=(20,3))


dingdancolumn = sg.Column([ [dingdantable,
                   sg.Frame('订单详情:', 
                    [[sg.Text('公司：',size=(6, 1),font=('MSYH', 12)),
                     sg.Combo(['宝马','大众'],key='-com-',
                        size=(6, 1),font=('MSYH', 12), disabled = True,text_color = 'red',background_color='white'),
                     sg.Text('日期：',size=(6, 1),font=('MSYH', 12)),
                     sg.Input(key='-date-', size=(8, 1),disabled = True,font=('MSYH', 12)),
                     sg.Text('订单号：',size=(6, 1),font=('MSYH', 12)),
                     sg.Input( key='-dingdannumber-', 
                        size=(16, 1),disabled = True,font=('MSYH', 12))],
                    [sg.Text('规格：',size=(6, 1),font=('MSYH', 12)),
                     sg.Input(key='-sty-', size=(8, 1),disabled = True,font=('MSYH', 12)),
                     sg.Text('单价：',size=(6, 1),font=('MSYH', 12)),
                     sg.Input(key='-value-', size=(8, 1),disabled = True,font=('MSYH', 12)),
                     sg.Text('数量：',size=(6, 1),font=('MSYH', 12)),
                     sg.Input(key='-num-', size=(8, 1),disabled = True,font=('MSYH', 12)),
                     sg.Text('总价：',size=(6, 1),font=('MSYH', 12)),
                     sg.Input(key='-valall-', size=(8, 1),disabled = True,font=('MSYH', 12))],
                   [sg.Text('长：',size=(6, 1),font=('MSYH', 12)),
                    sg.Input(key='-length-', size=(8, 1),disabled = True,font=('MSYH', 12)), 
                    sg.Text('宽：',size=(6, 1),font=('MSYH', 12)),
                    sg.Input(key='-width-', size=(8, 1),disabled = True,font=('MSYH', 12)), 
                    sg.Text('高：',size=(6, 1),font=('MSYH', 12)),
                    sg.Input(key='-height-', size=(8, 1),disabled = True,font=('MSYH', 12))], 
                   [sg.Text('状态：',size=(6, 1),font=('MSYH', 12)),
                    sg.Input(key='-zhuangtai-', size=(8, 1),disabled = True,font=('MSYH', 12))],# note must create a layout from scratch every time. No reuse
                   [sg.Button('保存',size=(6, 1),font=('MSYH', 12),key='-dingdanbaocun-'),
                    sg.Button('取消',font=('MSYH', 12),key='-dingdanquxiao-')]],
                   size=(600,750),
                   font=('MSYH', 12),
                   key='-frame2-')
                ], ], 
             pad=(0,0),
             vertical_alignment='bottom' 
             )
fahuodantable= sg.Frame('', [[sg.Table(values=fahuodandata,
                    headings=fahuodanheader_list,
                    max_col_width=500,                    
                    col_widths=[10,7,14,8,4,8,8],
                    auto_size_columns=False,
                    justification='center',
                    #display_row_numbers=True,
                    # alternating_row_color='lightblue',
                    num_rows=min(len(data)+50, 100),
                    key='-TABLE-',
                    selected_row_colors='red on yellow',                    
                    #right_click_menu=right_click_menu,
                    font=('微软雅黑', 12),
                    #enable_events = True,
                    enable_click_events = True,
                    vertical_scroll_only =False,
                    display_row_numbers=True
                    )]],size=(650,750)
                  ) 
fahuodancolumn = sg.Column([ [fahuodantable,
                   sg.Frame('发货详情:', 
                    [],
                   size=(600,750),
                   font=('MSYH', 12),
                   key='-framefahuo-')
                ], ], 
             pad=(0,0),
             vertical_alignment='bottom' 
             )
            
dingdanlayout =[dingdanbuttons,[dingdancolumn]   ]#订单
fahuodanlayout = [fahuodanbuttons,
               [fahuodancolumn]]
duizhangdanlayout = [duizhangdanbuttons,
               [sg.Input(key='-in2-')]]
layout = [
           [sg.TabGroup([[sg.Tab('订单管理', dingdanlayout,  key='-mykey-'),
                         sg.Tab('发货单管理', fahuodanlayout),
                         sg.Tab('对账单管理',duizhangdanlayout)]], key='-group1-',
                                tab_location='topleft',font=('MSYH', 12))]
          ]
dingdanwin = sg.Window('订单管理', layout,grab_anywhere=False,resizable=True,finalize=True) #订单

dingdanwin.Maximize()


def fahuodanddh(fahuodanhao,df):# 提取发货单中的订单编号
    a=df.loc[fahuodanhao,'订单']
    print('a',a)
    #c=a[0]
    #print('c',c)
    d=a.split(",")#把字符串c按逗号分裂为列表
    return d
def fahuodanbaocundingdan(fahuodanhao,dingdanhao,df): #向发货单中保存订单编号    
    if fahuodanhao in df.index.values.tolist(): #如果出货单号已经存在
        d=fahuodanddh(fahuodanhao,df) #订单信息列表
        print('d', d)
        if dingdanhao in d:
            sg.popup('订单编号重复，不能保存请重新修改',font=('MSYH', 12))
        else:            
            string2=','.join(d)#把列表插入逗号后连接为字符串
            print('string211', string2)
            string2=string2+','+ dingdanhao#追加订单
            print('string21', string2)
            df.loc[fahuodanhao,'订单'] = string2#追加订单到数据
            df.to_csv('./出货单.csv')#保存数据到文件
            print('string2', string2)
    else:#如果出货单号不存在直接保存
       print('订单号', dingdanhao)
       df.loc[fahuodanhao,'订单'] = dingdanhao#添加加订单到数据 
       df.to_csv('./出货单.csv')#保存数据到文件

def shuruzhuangt(a,b):#关闭且清空输入
    if a==1:
        c=True
    else:
        c=False
    if b==1:
        for keys in dingdankeys_list.keys():
            dingdanwin[keys].update('')#刷新窗口名称
    for keys in dingdankeys_list.keys():
        dingdanwin[keys].update(disabled = c)#刷新窗口名称
   
def xianshidingdanxinxi(dingdanbianhao): #显示订单信息
    df = pd.read_csv(filename,dtype = {'订单编号': str,'订单日期':str}) #打开数据文件
    #print('数据：',df)
    #dingdanshuju = df.loc[dingdanbianhao].values.tolist() #根据新订单编号确定数据 
    #print('订单数据为：',dingdanshuju)   
    df.set_index('订单编号',inplace=True)
    dingdanwin['-dingdannumber-'].update(dingdanbianhao,disabled = True)#刷新窗口名称

    # print(type(dingdanbianhao))

    # print(df.index)
    for keys in dingdankeys_list.keys():
        if keys !='-dingdannumber-' :
            valu=df.loc[dingdanbianhao,dingdankeys_list[keys]]      
            dingdanwin[keys].update(valu,disabled = True)#刷新窗口名称



def baocundingdan(dingdanbianhao,values): #保存订单  
    print('values:',values )
    print('dingdanbianhao:',dingdanbianhao,type(dingdanbianhao)) 
    df = pd.read_csv(filename,dtype = {'订单编号': str,'订单日期':str}) #打开数据文件
    df.set_index('订单编号',inplace=True)
    for key in dingdankeys_list.keys():
            if key !='-dingdannumber-' :
                df.loc[dingdanbianhao,dingdankeys_list[key]] = values[key]#修改数据 
    df.to_csv(filename)#保存数据到文件
    # header_list, data = getdata(filename,header_list,dingdanshujulx)#重新从文件读取数据
    # dingdanwin['-TABLE-'].update(data)#刷新数据
    # shuruzhuangt(1,0)#更新输入界面状态
def shanchudingdan(dingdanbianhao): #删除订单
    df = pd.read_csv(filename,dtype = {'订单编号' : str,'订单日期':str}) #打开数据文件
    df.set_index('订单编号',inplace=True)
    df=df.drop(dingdanbianhao) #根据新订单编号删除数据
    df.to_csv(filename) #保存数据到文件
    


 
while True: 
    
    ev1, values = dingdanwin.read()
    # print('事件为：',ev1)
    # print('值为：',values)
    if ev1!=None and ev1[0]=='-TABLE-' and ev1[1] == '+CICKED+' and ev1[2][0]!=None:         
        if ev1[2][0]>-1: #根据鼠标的点击实时显示订单信息
            hangn=ev1[2][0]#获取行号
            print('hangn:',hangn)
            dingdanbianhao=data[hangn][0]#根据行号获取订单编号
            #print('订单编号为：',dingdanbianhao)
            print('dingdanbianhao:',dingdanbianhao)
            dingdanwin['-frame2-'].update('订单信息')#刷新窗口名称
            xianshidingdanxinxi(dingdanbianhao)
            
    if ev1=='-fuzhidd-' and values['-TABLE-']!=[]: #复制数据
        #print('行号为：',values['-TABLE-'][0])
        hangn=values['-TABLE-'][0]#获取行号
        dingdanbianhao=data[hangn][0]#根据行号获取订单编号
        print('订单编号为：',dingdanbianhao)
        df = pd.read_csv(filename,index_col=0,header=None) #打开数据文件
        dingdanbianhaonew = dingdanbianhao+'复制'   
        print('订单编号为：',dingdanbianhaonew)          
        df.loc[dingdanbianhaonew] = df.loc[dingdanbianhao] #根据新订单编号复制数据        
        df.to_csv(filename,header=None)#保存数据到文件
        data = getdata(filename,header_list,dingdanshujulx)#重新从文件读取数据
        dingdanwin['-TABLE-'].update(data)#刷新数据
    if ev1=='-xiugdd-' and values['-TABLE-']!=[]: #修改订单
        dingdanwin['-frame2-'].update('修改订单')#刷新窗口名称
        shuruzhuangt(0,0)#更新输入界面状态        
        hangn=values['-TABLE-'][0]#获取行号
        print('行号为：',hangn)
        dingdanbianhao=data[hangn][0]#根据行号获取订单编号
        #print('订单编号为：',dingdanbianhao)
        ev1, values = dingdanwin.read()
        # print('事件为：',ev1)
        # print('值为：',values)
        # if ev2=='-dingdanbaocun-':
        if ev1=='-dingdanbaocun-':
             df = pd.read_csv(filename,index_col=0,header=None) #打开数据文件
             dingdanbianhaonew = values['-dingdannumber-']#确定新的订单编号
             print('dingdanbianhaonew:',dingdanbianhaonew,type(dingdanbianhaonew))
             print('dingdanbianhao:',dingdanbianhao,type(dingdanbianhao))  
             if type(dingdanbianhao)== type(dingdanbianhaonew):
                 if dingdanbianhao != dingdanbianhaonew :#如果订单编号做出了修改，下面修改订单编号
                     if dingdanbianhaonew in df.index: #如果修改后的订单编号和已经存在的订单编号重合
                        sg.popup('订单编号重复，不能保存请重新修改',font=('MSYH', 12))
                     else: #如果不重合
                        df.rename(index={dingdanbianhao:dingdanbianhaonew},inplace=True)#修改原订单编号
                        print('订单编号已修改') 
                        df.to_csv(filename,header=None)#保存数据到文件
                        baocundingdan(dingdanbianhaonew,values)
                 else: #如果订单编号没有修改则直接保存数据
                    baocundingdan(dingdanbianhaonew,values)
             else:
                sg.popup('新订单编号和旧订单编号数据类型不一致',font=('MSYH', 12))
                 
             data = getdata(filename,header_list,dingdanshujulx) #重新从文件读取数据
             dingdanwin['-TABLE-'].update(data) #刷新数据
            #shanchudingdan(dingdanbianhao) #删除旧订单  
        shuruzhuangt(1,0)#更新输入界面状态
        #ev1, values = dingdanwin.read()
    if ev1=='-shanchudd-' and values['-TABLE-']!=[]: #复制数据
        #print('行号为：',values['-TABLE-'][0])
        hangn=values['-TABLE-'][0]#获取行号
        dingdanbianhao=data[hangn][0]#根据行号获取订单编号
        shanchudingdan(dingdanbianhao) 
        data = getdata(filename,header_list,dingdanshujulx) #重新从文件读取数据
        dingdanwin['-TABLE-'].update(data) #刷新数据
        shuruzhuangt(1,1)#更新输入界面状态
    if ev1 == '-xinzdd-':
        shuruzhuangt(0,1)#更新输入界面状态
        
        dingdanwin['-frame2-'].update('新增订单')#刷新窗口名称
        df = pd.read_csv(filename,index_col=0,header=None) #打开数据文件
        
        ev1, values = dingdanwin.read()
        # print('事件为：',ev2)
        # print('值为：',values2)
        
        if ev1=='-dingdanbaocun-':
            dingdanbianhao = values['-dingdannumber-']
            if dingdanbianhao in df.index:
                sg.popup('订单编号重复，不能保存,请重新修改或者请点击取消',font=('MSYH', 12)) 
                while dingdanbianhao in df.index:                    
                    ev1, values = dingdanwin.read()
                    dingdanbianhao = values['-dingdannumber-']
                    if dingdanbianhao in df.index:
                       if ev1=='-dingdanquxiao-':
                           break
                       else:
                           sg.popup('订单编号重复，不能保存,请重新修改或者请点击取消',font=('MSYH', 12)) 
                    else:
                        baocundingdan(dingdanbianhao,values)
                     
            else:        
                baocundingdan(dingdanbianhao,values)
        data = getdata(filename,header_list,dingdanshujulx) #重新从文件读取数据
        dingdanwin['-TABLE-'].update(data) #刷新数据
        shuruzhuangt(1,0)#更新输入界面状态
    if ev1 in (None, '-dingdantuichu-') : 
        dingdanwin.close()
        break


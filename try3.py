
import configparser
from os import truncate
from socket import *
import time
import threading 
import os

import xlrd
import math
import numpy as np 
import xlsxwriter
from  openpyxl import  Workbook 


from mpl_toolkits.mplot3d import Axes3D
from matplotlib.pyplot import MultipleLocator
from datetime import datetime
from xlrd import xldate_as_tuple

from matplotlib import pyplot as plt
#解决 plt 中文显示的问题
plt.rcParams['font.sans-serif'] = ['SimHei']
plt.rcParams['axes.unicode_minus'] = False

#import pymysql




#【网络数据传输类定义】

class inf():
    def __init__(self,ip,sock,id):
        self.ip=ip
        self.sock=sock
        self.id=id


#【全局信息定义】

#本地文件夹路径设置    [注意route_0是 \ 形式]  [使用replace进行代换]
route_0=os.path.dirname(os.path.realpath(__file__))
route=route_0.replace("\\","\\\\")
#ini文件读取
inipath = route+"\\\\org.ini"           # org.ini的路径               
conf = configparser.ConfigParser()      # 创建管理对象
conf.read(inipath, encoding="utf-8")    # 读取ini文件
items = conf.items('ip_socket_set')     # 读入对应section
#写入表格数据标记量
write_number=0
#本地数据表patient_list=[]
patient_list=[]
#自身信息：
inf_self=inf(items[0][1],int(items[1][1]),items[2][1])
#组建LAN动态IP表                         # 用于储存本地inf数据 与被链接inf数据
inf_list=[]
inf_list.append(inf_self)               # 加入自身信息
#组建临时LAN接收IP表                     用于储存负反馈广播所接收到的inf数据
inf_temp_list=[]
#组建洪泛信息列表                        # 用于维护洪泛网络动态链接inf数据
inf_flooding_list=[]
items = conf.items('inf_flooding_list_set')     # 读入对应section

for i in range(0,int(len(items)/2)):            # 读取原始配置文件加入洪泛列表
    ip_temp=items[2*i][1]
    socket_temp=int(items[2*i+1][1])
    id_temp="unknown"
    inf_temp=inf(ip_temp,socket_temp,id_temp)
    inf_flooding_list.append(inf_temp)


#【网络交互函数】

#本地信息展示函数
def show_all_inf():
    print("the range is ",len(inf_list))
    for i in range(0,len(inf_list)):
        print("IP:",inf_list[i].ip,"socket:",inf_list[i].sock,"ID",inf_list[i].id)
    return 

#暂存展示函数
def show_all_inf_temp():
    print("the range is ",len(inf_temp_list))
    for i in range(0,len(inf_temp_list)):
        print("IP:",inf_temp_list[i].ip,"socket:",inf_temp_list[i].sock,"ID",inf_temp_list[i].id)
    return 

#接收线程
def Sever(): 
    s = socket(AF_INET, SOCK_STREAM)
    s.bind((inf_self.ip,inf_self.sock)) 
    s.listen(5) 

    while True: 
        #信息标识码 1 : 添加LAN身份请求
        #信息标识码 2 : 组网传输请求
        #信息标识码 3 : database数据共享请求
        #print("waitting for a new connection")
        conn, addr = s.accept()                                     #等待链接 阻塞本线程
        #print("Accept new connection from %s:%s" % addr) 
        sentence=conn.recv(1024).decode()                           #信息标识码读取
        temp_list=sentence.split()
        if temp_list[0]=="1":         #信息标识码 1 : 添加LAN身份请求
            #print("添加LAN身份请求")
            inf_temp=inf(temp_list[1],temp_list[2],temp_list[3])    #创建信息实例
            inf_list.append(inf_temp)                               #追加该身份信息
            show_all_inf()
            #print("更新完成")
        elif temp_list[0]=='2':       #信息标识码 2 : 组网传输请求
            #print("组网传输请求")
            inf_temp_list.clear()   #清空原有temp_ip表
            for i in range(0,int((len(temp_list)-1)/3)):
                inf_temp=inf(temp_list[3*i+1],temp_list[3*i+2],temp_list[3*i+3])    #创建信息实例
                inf_temp_list.append(inf_temp)                                      #追加该身份信息
            show_all_inf_temp()
            #print("更新完成")
        elif temp_list[0]=='3':       #信息标识码 3 : database数据共享请求
            #print("数据共享请求")
            filesize = str(os.path.getsize(route+"\\database.xlsx"))         #获取本地database文件大小
            conn.send(filesize.encode())                                     #传输本地database文件大小
            f = open(route+"\\database.xlsx",'rb')                           #打开本地database文件
            for line in f:                                                   #传输本地database文件
                conn.send(line)
            f.close()                                                        #进行文件关闭
        elif temp_list[0]=='4':       #信息标识码 4 : 洪泛搜索请求
            patient_id=temp_list[-1]                #取出目标病人id
            org_ip=temp_list[-3]                    #取出原始主机的ip
            org_socket=int(temp_list[-2])           #取出原始主机的端口号
            have_patient_id=False                   #查询本地病人信息id标识
            TTL=int(temp_list[1])                   #取出TTL值
            inf_ed_list=[]                          #取出已访问inf列表
            for i in range(0,int((len(temp_list)-3)/2)):
                inf_temp=inf(temp_list[2+2*i],temp_list[2+2*i+1],"unknown")
                inf_ed_list.append(inf_temp)
            patient_temp=''                         #暂存病人数据               
            for i in range(0,len(patient_list)):
                if patient_list[i].ID == patient_id:        #若发现所需id
                    patient_temp=patient_list[i]            #病例赋值
                    have_patient_id=True                    #更改标识
                    break
            if have_patient_id==True:                       #尝试写入temp.xlsx并传回数据 之后删除本地temp文件
                #组建请求报文
                data0="5"                                   #信息标识码 5 : 洪泛返回链接
                data0=data0+" "+inf_self.ip+" "+str(inf_self.sock)+" "+inf_self.id+" "+patient_id      #data0组建
                target_ip=org_ip                            #获取目标ip
                target_socket=org_socket                    #获取目标端口号
                s_temp = socket(AF_INET, SOCK_STREAM)
                flag=s_temp.connect_ex((target_ip, int(target_socket)))       #发起TCP链接
                jishu=0
                name="flooding_temp.xlsx"                                #本地中间传输文件命名
                patient_temp.write(route,name)                           #向本地文件夹写入中间传输文件
                filesize = str(os.path.getsize(route+"\\"+name))         #获取本地中间传输文件大小
                while True:
                    jishu=jishu+1
                    if flag==0:
                        s_temp.send(data0.encode())                              #传输指令标识信息
                        s_temp.send(filesize.encode())                           #传输中间传输文件大小
                        f = open(route+"\\"+name,'rb')                           #打开中间传输文件
                        for line in f:                                           #传输本地database文件
                            s_temp.send(line)
                        f.close()                   #关闭文件
                        s_temp.close()              #发送端成功传输信息，关闭TCP链接
                        os.remove(route+"\\"+name)  #删除中间文件
                        break
                    else:
                        if jishu>3:
                            break
                        flag=s_temp.connect_ex((target_ip, int(target_socket)))      #重新发起TCP链接
            else:                                   #判断TTL-1情况 检索发送序列 增添自己标识后将其发送
                TTL=TTL-1
                if TTL>0:                           #TTL仍然存活，检索发送序列进行发送
                    data0="4"                           #信息标识码 4 : 洪泛请求连接
                    data0=data0+" "+str(TTL)+" "+inf_self.ip+" "+str(inf_self.id)   #data0加入自身信息
                    for i in range(2,len(temp_list)):
                        data0=data0+" "+temp_list[i]                                #data0加入原始信息
                    for i in range(0,len(inf_flooding_list)):                            #该节点遍历洪泛peer列表
                        target_ip=inf_flooding_list[i].ip                                #获取目标ip
                        target_socket=inf_flooding_list[i].sock                          #获取目标端口号
                        same_flag=False                                                  #重复性检查标志位
                        for i1 in range(0,len(inf_ed_list)):                             #遍历已经过节点列表进行重复性检测
                            if target_ip==inf_ed_list[i1].ip and target_socket==inf_ed_list[i1].sock:
                                same_flag==True                                          #该节点已经历过查找
                                break
                        if same_flag==False:                                             #若该节点从未经历过查找
                            s_temp = socket(AF_INET, SOCK_STREAM)
                            flag=s_temp.connect_ex((target_ip, int(target_socket)))      #发起TCP链接
                            jishu=0
                            while True:
                                jishu=jishu+1
                                if flag==0:
                                    s_temp.send(data0.encode())                          #传输指令标识信息
                                    s_temp.close()              #发送端成功传输信息，关闭TCP链接
                                    break
                                else:
                                    if jishu>3:
                                        break
                                    flag=s.connect_ex((target_ip, int(target_socket)))   #重新发起TCP链接
        elif temp_list[0]=='5':       #信息标识码 5 : 洪泛文件传回请求
            from_ip=temp_list[-4]                           #来源ip储存
            from_socket=temp_list[-3]                       #来源sock储存
            from_id=temp_list[-2]                           #来源id储存
            patient_id=temp_list[-1]                        #传回patient_id储存
            data_range = conn.recv(1024)                    #首次获取传输文件长度信息
            file_total_size = int(data_range.decode())      #获取传输文件长度
            received_size = 0                               #记录已获取长度
            name="flooding_temp.xlsx"                       #中间传输文件名
            f = open(route+"\\"+name, 'wb')                 #打开临时储存文件
            while received_size < file_total_size:          
                data = conn.recv(1024)
                f.write(data)
                received_size += len(data)  
            f.close()                                       #关闭文件
            read(route+"\\"+name)                           #进行文件读取 和 校验
            os.remove(route+"\\"+name)                      #删除临时文件
            print("")
            print("recieve:",patient_id,"from",from_ip,from_socket,from_id)     #回执信息打印

#分布式网络共享模块
thread_sever=threading.Thread(target=Sever, args=())
thread_sever.start()




#【控制函数】

#ini_list_default
def ini_list_default():
    patient_list.clear()                   #清空原有数据库
    try:
        read(route+"\\"+"database.xlsx")   #读取database_temp 
        arrangement()                      #进行整理
        remake_database()                  #重新进行计算
    except IOError:
        print("can't find this file")
    else:
        print("init finished")




#发送线程 [同时负责程序控制]
ini_list_default()                      #初始化本地信息




while True:
    # 发送端设置:
    # show_database                     显示本地database
    # show_3D                           显示本地3D文件
    # write_zhou_angle                  写出支架轴向夹角文件 zhou_angle
    # write_statistics_zhou_angle       写出支架轴向夹角文件 statistics_zhou_angle
    # write_huan_shape                  写出支架首尾环文件   huan_shape
    # -1   - 查看本地ip表 ip-temp表
    # 0    - 发送添加身份请求
    # 1    - LAN发送组网传输回馈 
    # 2.1  - database模块数据传输事务添加请求 使用ip-list
    # 2.2  - database模块数据传输事务添加请求 使用ip-temp-list
    # 3    - 洪泛搜索特定ID患者信息请求       使用inf_flooding_list

    a = input("please input oder:")
    #局域网控制命令
    if a == "-1":
        print("ip-list")
        show_all_inf()
        print("ip-temp-list")
        show_all_inf_temp()
    elif a == "0":        #添加身份请求 [输入目标主机IP与SOCKET进行身份信息添加]
        target_ip=input("please input target_ip:")                  #输入目标ip
        target_socket=input("please input target_socket:")          #输入目标端口号
        s = socket(AF_INET, SOCK_STREAM)
        data="1"+" "+inf_self.ip+" "+str(inf_self.sock)+" "+inf_self.id     #组建传输信息

        flag=s.connect_ex((target_ip, int(target_socket)))                          #发起TCP链接
        jishu=0
        while True:
            jishu=jishu+1
            if flag==0:
                s.send(data.encode())
                s.close()                   #发送端成功传输信息，关闭TCP链接
                jishu=0                     #jishu=0表明正常发送
                break
            else:
                if jishu>3:
                    jishu=-1                #jishu=-1表明发送3次失败，放弃本次传输
                    break
                time.sleep(1)
                flag=s.connect_ex((target_ip, int(target_socket)))                  #重新发起TCP链接
        if jishu==0:
            print(target_ip,target_socket,"successfully sent")
        elif jishu==-1:
            print(target_ip,target_socket,"fail in send")
    elif a == "1":        #LAN发送组网传输回馈 [将本地inf_list表广播至本网络节点]
        #组建data报文
        data='2'
        for i in range(0,len(inf_list)):
            data=data+" "+inf_list[i].ip+" "+str(inf_list[i].sock)+" "+inf_list[i].id
        #进行TCP文件传输
        for i in range(0,len(inf_list)):
            if inf_list[i].ip==inf_self.ip and inf_list[i].sock==inf_self.sock:
                continue
            else:
                target_ip=inf_list[i].ip
                target_socket=inf_list[i].sock
                s = socket(AF_INET, SOCK_STREAM)

                flag=s.connect_ex((target_ip, int(target_socket)))                          #发起TCP链接
                jishu=0
                while True:
                    jishu=jishu+1
                    if flag==0:
                        s.send(data.encode())
                        s.close()                   #发送端成功传输信息，关闭TCP链接
                        jishu=0                     #jishu=0表明正常发送
                        break
                    else:
                        if jishu>3:
                            jishu=-1                #jishu=-1表明发送3次失败，放弃本次传输
                            break
                        time.sleep(0.1)
                        flag=s.connect_ex((target_ip, int(target_socket)))                  #重新发起TCP链接
                if jishu==0:
                    print(target_ip,target_socket,"successfully sent")
                elif jishu==-1:
                    print(target_ip,target_socket,"fail in send")
    elif a == '2.1':      #请求进行database共享 按ip-list寻址
        #组建请求报文
        data0='3' 
        str0=" temp temp"
        data0=data0+" temp temp"
        data0=str(data0)
        #进行TCP文件传输
        for i in range(0,len(inf_list)):
            if inf_list[i].ip==inf_self.ip and inf_list[i].sock==inf_self.sock:
                continue
            else:
                target_ip=inf_list[i].ip
                target_socket=inf_list[i].sock
                s = socket(AF_INET, SOCK_STREAM)
                flag=s.connect_ex((target_ip, int(target_socket)))                          #发起TCP链接
                jishu=0
                while True:
                    jishu=jishu+1
                    if flag==0:                         #连接正常 开始进行database传输 获得的database存放于database-temp中 用完即删
                        s.send(data0.encode())          #发送标识信息 信息标识码 - 3
                        data_range = s.recv(1024)       #首次获取传输文件长度信息
                        file_total_size = int(data_range.decode())       #获取传输文件长度
                        received_size = 0                                #记录已获取长度
                        f = open(route+"\\database_temp.xlsx", 'wb')              #打开临时储存文件
                        while received_size < file_total_size:           #未接收完时：
                            data = s.recv(1024)
                            f.write(data)
                            received_size += len(data)
                        #接收完成 关闭TCP 与文件
                        s.close()
                        f.close()
                        jishu=0
                        #进行文件读取 
                        read(route+"\\database_temp.xlsx")
                        #进行文件校验
                        arrangement()
                        #进删除临时文件
                        os.remove(route+"\\database_temp.xlsx")
                        break
                    else:
                        if jishu>3:
                            jishu=-1
                            break
                        time.sleep(0.1)
                        flag=s.connect_ex((target_ip, int(target_socket)))                   #重新发起TCP链接
                if jishu==0:
                    print(target_ip,target_socket,"successfully sent")
                elif jishu==-1:
                    print(target_ip,target_socket,"fail in send")
    elif a == '2.2':      #请求进行database共享 按ip-temp-list寻址
        #组建请求报文
        data0='3' 
        str0=" temp temp"
        data0=data0+" temp temp"
        data=str(data0)
        #进行TCP文件传输
        for i in range(0,len(inf_temp_list)):
            if inf_temp_list[i].ip==inf_self.ip and inf_temp_list[i].sock==inf_self.sock:
                continue
            else:
                target_ip=inf_temp_list[i].ip
                target_socket=inf_temp_list[i].sock
                s = socket(AF_INET, SOCK_STREAM)
                flag=s.connect_ex((target_ip, int(target_socket)))                          #发起TCP链接
                jishu=0
                while True:
                    jishu=jishu+1
                    if flag==0:                         #连接正常 开始进行database传输 获得的database存放于database-temp中 用完即删
                        s.send(data0.encode())          #发送标识信息 信息标识码 - 3
                        data_range = s.recv(1024)       #首次获取传输文件长度信息
                        file_total_size = int(data_range.decode())       #获取传输文件长度
                        received_size = 0                                #记录已获取长度
                        f = open(route+"\\database_temp.xlsx", 'wb')              #打开临时储存文件
                        while received_size < file_total_size:           #未接收完时：
                            data = s.recv(1024)
                            f.write(data)
                            received_size += len(data)
                        #接收完成 关闭TCP 关闭文件
                        s.close()
                        f.close()
                        jishu=0
                        #进行文件读取 和 校验
                        read(route+"\\database_temp.xlsx")
                        #删除临时文件
                        os.remove(route+"\\database_temp.xlsx")
                        break
                    else:
                        if jishu>3:
                            jishu=-1
                            break
                        time.sleep(0.1)
                        flag=s.connect_ex((target_ip, int(target_socket)))                   #重新发起TCP链接
                if jishu==0:
                    print(target_ip,target_socket,"successfully sent")
                elif jishu==-1:
                    print(target_ip,target_socket,"fail in send")
    elif a == "3":        #洪泛网络搜寻特定ID数据
        #组建请求报文
        data0="4"                                               #信息标识码4 ：洪泛数据请求
        TTL_temp=input("please define the TTL:")              #定义TTL
        patient_id=input("please input the patient's ID:")    #输入病人ID
        data0=data0+" "+str(TTL_temp)+" "+inf_self.ip+" "+str(inf_self.sock)+" "+patient_id
        for i in range(0,len(inf_flooding_list)):
            target_ip=inf_flooding_list[i].ip                   #获取目标ip
            target_socket=inf_flooding_list[i].sock             #获取目标端口号
            s = socket(AF_INET, SOCK_STREAM)
            flag=s.connect_ex((target_ip, int(target_socket)))  #发起TCP链接
            jishu=0
            while True:
                jishu=jishu+1
                if flag==0:
                    s.send(data0.encode())
                    s.close()                   #发送端成功传输信息，关闭TCP链接
                    jishu=0                     #jishu=0表明正常发送
                    break
                else:
                    if jishu>3:
                        jishu=-1                #jishu=-1表明发送3次失败，放弃本次传输
                        break
                    time.sleep(1)
                    flag=s.connect_ex((target_ip, int(target_socket)))                  #重新发起TCP链接
            if jishu==0:
                print(target_ip,target_socket,"successfully sent")
            elif jishu==-1:
                print(target_ip,target_socket,"fail in send")
    #洪泛查找命令
    elif a == "show_flooding_list":
        for i in range(0,len(inf_flooding_list)):
            print("IP:",inf_flooding_list[i].ip,"socket:",inf_flooding_list[i].sock,"ID",inf_flooding_list[i].id)
    elif a == "clear_flooding_list":
        inf_flooding_list.clear()
    elif a == "append_flooding_list":
        peer_ip=input("please input peer_ip:")                  #输入目标ip
        peer_socket=input("please input peer_socket:")          #输入目标端口号
        inf_temp=inf(peer_ip,int(peer_socket),"unknown")
        inf_flooding_list.append(inf_temp)                      #追加peer_temp
    #内存数据list操作命令
    elif a == "clear_list":             #清空当前list
        patient_list.clear()            
    elif a == "ini_list":               #用户输入初始化数据库名 进行初始化
        patient_list.clear()                   #清空原有数据库
        database_temp = input("please input filename:")
        try:
            read(route+"\\"+database_temp)     #读取database_temp 
            arrangement()                      #进行整理
            remake_database()                  #重新进行计算
        except IOError:
            print("can't find this file")
        else:
            print("finished")     
    elif a == "ini_list_default":       #使用本地的database文件完成初始化
        ini_list_default()
    elif a == "appened_list":           #用户输入数据库名进行追加
        database_temp = input("please input filename:")
        try:
            read(route+"\\"+database_temp)     #读取database_temp 追加
            arrangement()                      #进行整理
        except IOError:
            print("can't find this file")
        else:
            print("finished")    
    elif a == "write_list":             #用户输入文件名 进行文件覆盖输出
        file_temp = input("please input filename:")
        #文件追加
        workbook = xlsxwriter.Workbook(route+"\\"+file_temp)
        worksheet = workbook.add_worksheet('Sheet1')
        number_x=0                                      #记录x写入位置 从第一列开始读入        
        number_y=1                                      #记录y写入位置 从第二行开始读入
        for j in range(0,len(patient_list)):
            patient_temp=patient_list[j]        #取出patient_list中待写入病例

            for i in range(0,patient_temp.period_length):
                number_x=0                                      #x归位
                worksheet.write(number_y,number_x,patient_temp.name)            #写入姓名
                worksheet.write(number_y,number_x+1,patient_temp.ID)            #写入ID
                worksheet.write(number_y,number_x+2,patient_temp.sex)           #写入性别
                worksheet.write(number_y,number_x+3,patient_temp.birthday)      #写入出生年月
                number_x=4                                      #x归位
                date_time_temp = datetime.strptime(patient_temp.period_list[i].data, '%Y/%m/%d')
                worksheet.write_datetime(number_y+i,number_x,date_time_temp)                            #写入日期
                worksheet.write(number_y+i,number_x+1,patient_temp.period_list[i].top12_x)              #写入x
                worksheet.write(number_y+i,number_x+2,patient_temp.period_list[i].top12_y)              #写入y
                worksheet.write(number_y+i,number_x+3,patient_temp.period_list[i].top12_z)              #写入z
                number_x=number_x+4                             #x移位
                #支架写入部分
                for i1 in range(0,len(patient_temp.period_list[i].stent_list)):                         #遍历所有支架
                    worksheet.write(number_y+i,number_x,"#")    #写入#
                    worksheet.write(number_y+i,number_x+1,patient_temp.period_list[i].stent_list[i1].stent_type)
                    worksheet.write(number_y+i,number_x+2,patient_temp.period_list[i].stent_list[i1].stent_shape)
                    number_x=number_x+3                         #x移位
                    for i2 in range(0,len(patient_temp.period_list[i].stent_list[i1].huan_list)):
                        huan_temp=patient_temp.period_list[i].stent_list[i1].huan_list[i2]
                        for i3 in range(0,len(huan_temp.point_list)):
                            point_temp=huan_temp.point_list[i3]
                            worksheet.write(number_y+i,number_x,point_temp.x) 
                            worksheet.write(number_y+i,number_x+1,point_temp.y)
                            worksheet.write(number_y+i,number_x+2,point_temp.z)
                            number_x=number_x+3                 #x移位
                #血管数据写入部分
                #x,y,z,Tx,Ty,Tz,Nx,Ny,Nz,BNx,BNy,BNz,Dfit,Dmin,Dmax,C,Dh,Xh,Scf,Area,E
                if len(patient_temp.period_list[i].see_point_list)!=0:  #若该时期存在有血管数据
                    worksheet.write(number_y+i,number_x,"**")
                    number_x=number_x+1                         #x移位
                    for i1 in range(0,len(patient_temp.period_list[i].see_point_list)):
                        see_point_temp=patient_temp.period_list[i].see_point_list[i1]
                        for i2 in range(0,len(see_point_temp.information_list)):
                            worksheet.write(number_y+i,number_x+i2,see_point_temp.information_list[i2])
                        number_x=number_x+len(see_point_temp.information_list)                  #x移位
                #末位标识写入
                worksheet.write(number_y+i,number_x,"##")
            number_x=4                                                                          #x归位
            number_y=number_y+patient_temp.period_length                                        #y移位
            #第零时期写入
            if patient_temp.calculation==True:                                                  #若存在第零时期
                date_time_temp = datetime.strptime(patient_temp.period_list[i].data, '%Y/%m/%d')
                worksheet.write_datetime(number_y,number_x,date_time_temp)                      #写入日期
                worksheet.write(number_y,number_x+1,patient_temp.period_list[-1].top12_x)    #写入x
                worksheet.write(number_y,number_x+2,patient_temp.period_list[-1].top12_y)    #写入y
                worksheet.write(number_y,number_x+3,patient_temp.period_list[-1].top12_z)    #写入z
                number_x=number_x+4                                      #x移位
                worksheet.write(number_y,number_x,"**")                  #写入标识符
                number_x=number_x+1                                      #x移位
                for i1 in range(0,len(patient_temp.period_list[-1].see_point_list)):
                        see_point_temp=patient_temp.period_list[-1].see_point_list[i1]
                        for i2 in range(0,len(see_point_temp.information_list)):
                            worksheet.write(number_y,number_x+i2,see_point_temp.information_list[i2])
                        number_x=number_x+len(see_point_temp.information_list)                 #x移位
                #末位标识写入
                worksheet.write(number_y,number_x,"##")
                number_y=number_y+1
        workbook.close()                #写入完成后关闭文件并保存
        print("finished")
   
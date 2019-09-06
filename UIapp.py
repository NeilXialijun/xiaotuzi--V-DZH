# -*- coding:utf-8 -*-
from mousePrint import *
from opencvCC import *
from xlutils.copy import copy
from tkinter import Tk, Button, Canvas, ttk, Frame, Label
from PIL import Image, ImageTk, ImageFont
import re

from tkinter import Text, IntVar, Entry, LEFT, StringVar, BOTH, END, LEFT, DISABLED
import threading
import string
import time
#import cv2 as cv
import numpy as np
#import requests
import xlwt
import xlrd
import datetime
import logging
import unittest
# import uuid
#import multiprocessing

logging.basicConfig(filename='E:/xiaotuzi/%s-JCWY'% time.strftime('%Y-%m-%d', time.localtime(time.time())), 
    format='%(asctime)s:%(message)s', 
    level = logging.DEBUG,filemode='a',
    datefmt='%Y-%m-%d%I:%M:%S %p')

class Application(object):

    def __init__(self, master=None):
        self.root = master    #    定义内部变量root
        self.root.geometry('655x418')
        self.root.resizable(0,0)
        logging.info("start ：...")
        
        # # 创建一个容器,
        self.frm = ttk.LabelFrame(self.root, text="picture processing")     # 创建一个容器，其父容器为win
        self.frm.place(x=0, y=0, anchor="nw", width=650, height=230)
        self.frm_p = Frame(self.frm)

        self.num_f = ttk.LabelFrame(self.root, text="data processing")     # 创建一个容器，其父容器为win
        self.num_f.place(x=0, y=235, anchor="nw", width=650, height=180)
        self.frm_num = Frame(self.num_f)

        self.clean_coordinate_str = ""
        self.final_data = ""
        self.results_not_count = 0
        self.bet_count = 0
        self.bet_flag = 0

        self.bet_count2 = 0
        self.bet_flag2 = 0
        self.results_not_count2 = 0

        self.bet_count3 = 0
        self.bet_flag3 = 0
        self.results_not_count3 = 0

        self.bet_count4 = 0
        self.bet_flag4 = 0
        self.results_not_count4 = 0   

        self.bet_count5 = 0
        self.bet_flag5 = 0
        self.results_not_count5 = 0   

        self.bet_count6 = 0
        self.bet_flag6 = 0
        self.results_not_count6 = 0   

        self.bet_count7 = 0
        self.bet_flag7 = 0
        self.results_not_count7 = 0           
 
        self.bet_count8 = 0
        self.bet_flag8 = 0
        self.results_not_count8 = 0   
        
        self.not_neet_bet1 = 0
        self.not_neet_bet2 = 0
        self.not_neet_bet3 = 0
        self.not_neet_bet4 = 0
        self.not_neet_bet5 = 0
        self.not_neet_bet6 = 0
        self.not_neet_bet7 = 0
        self.not_neet_bet8 = 0        

        self.The_lottery_results_old = []
        self.The_lottery_results = []
        self.periods = 0
        self.event = threading.Event();
        self.auto_thread = None
        # self.pack(expand=YES, fill=BOTH)
        self.excel_init()
        self.excel_test_init()
        self.excel_auto_click_init()

        self.createWidgets()
        self.start_display_xy_thread()


    def get_clean_coordinate(self):
        self.clean_coordinate_str = self.clean_pixel.get()
        list1 = []
        if self.clean_coordinate_str == '' or self.clean_coordinate_str == "请输入坐标":
            self.e6.delete(0, END)
            self.e6.insert(0, "请输入坐标")
            return [50, 100, 150, 200, 250, 300, 350, 400, 450, 500]
        try:
            list1 = self.clean_coordinate_str.split('-', 9)
            list1 = list(map(int, list1))
        except ValueError:
            pass
        return list1

    def excel_init(self):
        self.workbook = xlwt.Workbook()
        self.data = time.strftime('%Y-%m-%d', time.localtime(time.time()))
        self.worksheet = self.workbook.add_sheet(self.data)
        self.pattern = xlwt.Pattern()  # Create the Pattern
        self.pattern.pattern = xlwt.Pattern.SOLID_PATTERN  # May be: NO_PATTERN, SOLID_PATTERN, or 0x00 through 0x12
        self.pattern.pattern_fore_colour = 1
        #  May be: 8 through 63. 0 = Black, 1 = White, 2 = Red, 3 = Green, 4 = Blue, 5 = Yellow, 6 = Magenta, 7 = Cyan, 16 = Maroon, 17 = Dark Green, 18 = Dark Blue, 19 = Dark Yellow , almost brown), 20 = Dark Magenta, 21 = Teal, 22 = Light Gray, 23 = Dark Gray, the list goes on...
        self.style = xlwt.XFStyle()  # Create the Pattern
        self.styleBlueBkg = xlwt.easyxf('pattern: pattern solid, fore_colour red;')  # 红色


        # style.pattern = pattern # Add Pattern to Style
        self.worksheet.write(0, 0, '期数', self.style)
        self.worksheet.write(0, 1, '号码', self.style)

    def createWidgets(self):
        self.frm_x = Frame(self.frm)
        self.Label1 = Label(self.frm_x, text='X start:').grid(row=0, column=0)
        self.frm_x.place(x=0, y=50)

        self.frm_y = Frame(self.frm)
        self.Label2 = Label(self.frm_y, text='Y start:').grid(row=0, column=2)
        self.frm_y.place(x=100, y=50)

        self.X_coordinate = IntVar()
        self.e1 = Entry(self.frm_x, width=5, textvariable=self.X_coordinate)    # Entry 是 Tkinter 用来接收字符串等输入的控件. state='disabled'
        self.e1.grid(row=0, column=1, padx=1, pady=1)  # 设置输入框显示的位置，以及长和宽属性

        self.Y_coordinate = IntVar()
        self.e2 = Entry(self.frm_y, width=5, textvariable=self.Y_coordinate, justify=LEFT)
        self.e2.grid(row=0, column=3, padx=1, pady=1)

# # ====================================================================================
        self.frm_xx = Frame(self.frm)
        self.frm_yy = Frame(self.frm)
        self.Labe1 = Label(self.frm_xx, text='X end:', justify=LEFT).grid(row=0, column=4)
        self.Labe2 = Label(self.frm_yy, text='Y end:', justify=LEFT).grid(row=0, column=6)
        self.frm_xx.place(x=300, y=50)
        self.frm_yy.place(x=400, y=50)

        self.Xx_coordinate = IntVar()
        self.e3 = Entry(self.frm_xx, width=5, textvariable=self.Xx_coordinate, justify=LEFT)    # Entry 是 Tkinter 用来接收字符串等输入的控件.
        self.e3.grid(row=0, column=5, padx=1, pady=1)  # 设置输入框显示的位置，以及长和宽属性

        self.Yy_coordinate = IntVar()
        self.e4 = Entry(self.frm_yy, width=5, textvariable=self.Yy_coordinate, justify=LEFT)
        self.e4.grid(row=0, column=7, padx=1, pady=1)

        self.test_BT_f = Frame(self.frm)
        self.test_BT = Button(self.test_BT_f, text='startTest', width=10, command=self.TestStart)
        self.test_BT.grid(row=0, column=4, padx=1, pady=1)
        self.test_BT_f.place(x=550, y=25)

        self.stop_BT_F = Frame(self.frm)
        self.stop_BT = Button(self.stop_BT_F, text='pause', width=10, command=self.stop_thread)
        self.stop_BT .grid(row=0, column=4, padx=1, pady=1)
        self.stop_BT_F.place(x=550, y=70)


        self.start_BT_F = Frame(self.frm)
        self.start_BT = Button(self.start_BT_F, text='startCatch', width=10, command=self.start_auto_thread)
        self.start_BT .grid(row=0, column=4, padx=1, pady=1)
        self.start_BT_F.place(x=550, y=120)


# # -------------清理条 输入 ---------------------------------------------------------
        self.clean_pixel_f = Frame(self.frm)
        self.clean_pixel_l = Label(self.clean_pixel_f, text='xy-clean:', justify=LEFT).grid(row=0, column=11)

        self.clean_pixel = StringVar()
        self.e6 = Entry(self.clean_pixel_f, width=50, textvariable=self.clean_pixel, justify=LEFT)
        self.e6.grid(row=0, column=12, padx=1, pady=1)
        self.clean_pixel_f.place(x=50, y=150)

# # -------------鼠标 坐标------------------------------------------------------------------
        self.mouse_f = Frame(self.frm)
        self.mouse_l = Label(self.mouse_f, text='coordinate:', justify=LEFT).grid(row=0, column=11)

        self.set_xy = StringVar()
        self.Temp = ("%s.%s") % get_mouse_point()
        self.e7 = Entry(self.mouse_f, width=10, textvariable=self.Temp)
        self.e7.grid(row=0, column=12, padx=1, pady=1)
        self.mouse_f.place(x=520, y=170)


# # -------------识别结果------------------------------------------------------------------
        self.recognition_f = Frame(self.frm)
        self.recognition_l = Label(self.recognition_f, text='result :', justify=LEFT).grid(row=0, column=11)

        self.e8 = Entry(self.recognition_f, width=50, textvariable=self.final_data)
        self.e8 .grid(row=0, column=12, padx=1, pady=1)
        self.recognition_f.place(x=50, y=180)
# # -----------------------------------------------------------------------------------------------------
# # ****************************     数据处理layout    **************************************************
# # -----------------------------------------------------------------------------------------------------
        self.M_enter_f = Frame(self.num_f)
        self.M_enter_l = Label(self.M_enter_f, text='M-number :', justify=LEFT).grid(row=0, column=0)

        self.M_enter_s = StringVar()
        self.e9 = Entry(self.M_enter_f, width=50, textvariable=self.M_enter_s)
        self.e9.grid(row=0, column=12, padx=1, pady=1)
        self.M_enter_f.place(x=10, y=10)

        self.num_enter_f = Frame(self.num_f)
        self.num_enter_l = Label(self.num_enter_f, text='NO1:', justify=LEFT).grid(row=0, column=0)

        self.num_enter_s = StringVar()
        self.ea = Entry(self.num_enter_f, width=20, textvariable=self.num_enter_s)
        self.ea.grid(row=0, column=12, padx=1, pady=1)
        self.num_enter_f.place(x=10, y=40)
        
        self.num_enter2_f = Frame(self.num_f)
        self.num_enter2_l = Label(self.num_enter2_f, text='NO2:', justify=LEFT).grid(row=0, column=0)

        self.num_enter2_s = StringVar()
        self.ea2 = Entry(self.num_enter2_f, width=20, textvariable=self.num_enter2_s)
        self.ea2.grid(row=0, column=12, padx=1, pady=1)
        self.num_enter2_f.place(x=10, y=65)

        self.num_enter3_f = Frame(self.num_f)
        self.num_enter3_l = Label(self.num_enter3_f, text='NO3:', justify=LEFT).grid(row=0, column=0)

        self.num_enter3_s = StringVar()
        self.ea3 = Entry(self.num_enter3_f, width=20, textvariable=self.num_enter3_s)
        self.ea3.grid(row=0, column=12, padx=1, pady=1)
        self.num_enter3_f.place(x=10, y=90)
        

        self.num_enter4_f = Frame(self.num_f)
        self.num_enter4_l = Label(self.num_enter4_f, text='NO4:', justify=LEFT).grid(row=0, column=0)

        self.num_enter4_s = StringVar()
        self.ea4 = Entry(self.num_enter4_f, width=20, textvariable=self.num_enter4_s)
        self.ea4.grid(row=0, column=12, padx=1, pady=1)
        self.num_enter4_f.place(x=10, y=115)


        self.num_enter5_f = Frame(self.num_f)
        self.num_enter5_l = Label(self.num_enter5_f, text='NO5:', justify=LEFT).grid(row=0, column=0)

        self.num_enter5_s = StringVar()
        self.ea5 = Entry(self.num_enter5_f, width=20, textvariable=self.num_enter5_s)
        self.ea5.grid(row=0, column=12, padx=1, pady=1)
        self.num_enter5_f.place(x=250, y=90)
        
        
        self.num_enter6_f = Frame(self.num_f)
        self.num_enter6_l = Label(self.num_enter6_f, text='NO6:', justify=LEFT).grid(row=0, column=0)

        self.num_enter6_s = StringVar()
        self.ea6 = Entry(self.num_enter6_f, width=20, textvariable=self.num_enter6_s)
        self.ea6.grid(row=0, column=12, padx=1, pady=1)
        self.num_enter6_f.place(x=250, y=65)

        self.num_enter7_f = Frame(self.num_f)
        self.num_enter7_l = Label(self.num_enter7_f, text='NO7:', justify=LEFT).grid(row=0, column=0)

        self.num_enter7_s = StringVar()
        self.ea7 = Entry(self.num_enter7_f, width=20, textvariable=self.num_enter7_s)
        self.ea7.grid(row=0, column=12, padx=1, pady=1)
        self.num_enter7_f.place(x=250, y=40)

        self.num_enter8_f = Frame(self.num_f)
        self.num_enter8_l = Label(self.num_enter8_f, text='NO8:', justify=LEFT).grid(row=0, column=0)

        self.num_enter8_s = StringVar()
        self.ea8 = Entry(self.num_enter8_f, width=20, textvariable=self.num_enter8_s)
        self.ea8.grid(row=0, column=12, padx=1, pady=1)
        self.num_enter8_f.place(x=250, y=115)        
        

        self.num_NO_f  = Frame(self.num_f)
        self.num_NO_BT = Button(self.num_NO_f, text='R_M_test', width=10, command=lambda: auto_click.NO_Num_BT(auto_click, self.aotu_click_sheet))
        self.num_NO_BT.grid(row=0, column=4, padx=1, pady=1)
        # Button(self.num_NO_f, text='名号测试', width=10, command='') \
        #     .grid(row=0, column=4, sticky=E, padx=1, pady=1)
        self.num_NO_f.place(x=550, y=40)

        self.M_f = Frame(self.num_f)
        self.M_BT = Button(self.M_f, text='M tese', width=10, command=lambda: auto_click.M_BT(auto_click, self.aotu_click_sheet))
        self.M_BT.grid(row=0, column=4, padx=1, pady=1)
        self.M_f.place(x=550, y=80)

        self.M_list = self.M_enter_s.get()
        self.reality_f = Frame(self.num_f)
        self.reality_BT = Button(self.reality_f, text='real test', width=10, command=lambda: auto_click.reality_BT(auto_click, self.aotu_click_sheet))
        self.reality_BT.grid(row=0, column=4, padx=1, pady=1)
        self.reality_f.place(x=550, y=120)

        # self.periods_select = Frame(self.num_f)
        # self.num_enter_l = Label(self.periods_select, text='周期间隔 :').grid(row=0, column=0)

        # self.comvalue = StringVar()  # 窗体自带的文本，新建一个值
        # self.comboxlist = ttk.Combobox(self.periods_select, width=6, textvariable=self.comvalue)  # 初始化
        # self.comboxlist["values"] = ('每一期',0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10)
        # self.comboxlist.current(4)  # 选择第一个
        # self.comboxlist.grid(row=0, column=6)
        # self.periods_select.place(x=450, y=8)

        # temp2 = [5,2,6,3,10,4,1,9,8,7]
        # self.number_of_periods = 140
        self.test_f = Frame(self.num_f)
        self.test_BT = Button(self.test_f, text='ZStest', width=10, command=lambda: self.test_BT_fun())
        self.test_BT.grid(row=0, column=4, padx=1, pady=1)
        self.test_f.place(x=550, y=0)

    def stop_thread(self):
        print("stop thread")
        self.event.clear()
        self.start_BT['state'] = "normal"
        self.e1["state"] = "normal"
        self.e2["state"] = "normal"
        self.e3["state"] = "normal"
        self.e4["state"] = "normal"
        self.e6["state"] = "normal"
        self.e9["state"] = "normal"
        self.ea["state"] = "normal"
        self.ea2["state"] = "normal"
        self.ea3["state"] = "normal"
        self.ea4["state"] = "normal"
        self.ea5["state"] = "normal"
        self.ea6["state"] = "normal"
        self.ea7["state"] = "normal"
        self.ea8["state"] = "normal"
        #self.comboxlist["state"] = "normal"
        self.num_NO_BT["state"] = "normal"
        self.M_BT["state"] = "normal"
        self.reality_BT["state"] = "normal"
        self.test_BT["state"] = "normal"

    def get_M_num_data(self):
        self.num_list = self.num_enter_s.get()
        self.num_list2 = self.num_enter2_s.get()
        self.num_list3 = self.num_enter3_s.get()
        self.num_list4 = self.num_enter4_s.get()

        self.num_list5 = self.num_enter5_s.get()
        self.num_list6 = self.num_enter6_s.get()
        self.num_list7 = self.num_enter7_s.get()
        self.num_list8 = self.num_enter8_s.get()        
        
        self.M_list = self.M_enter_s.get()
        if self.num_list is '':
            self.ea.insert(0, "请输入号码，格式 1-2-3...")
            return None

        if self.num_list2 is '':
            print("第二组没数据")
        else:
            self.num_list2 = self.num_list2.split('-', 5)
            self.num_list2 = list(map(int, self.num_list2))

        if self.num_list3 is '':
            print("第三组没数据")
        else:
            self.num_list3 = self.num_list3.split('-', 5)
            self.num_list3 = list(map(int, self.num_list3))
            
        if self.num_list4 is'':
            print("第四组没数据")
        else:
            self.num_list4 = self.num_list4.split('-', 5)
            self.num_list4 = list(map(int, self.num_list4))

        if self.num_list5 is'':
            print("第五组没数据")
        else:
            self.num_list5 = self.num_list5.split('-', 5)
            self.num_list5 = list(map(int, self.num_list5))

        if self.num_list6 is'':
            print("第六组没数据")
        else:
            self.num_list6 = self.num_list6.split('-', 5)
            self.num_list6 = list(map(int, self.num_list6))

        if self.num_list7 is'':
            print("第七组没数据")
        else:
            self.num_list7 = self.num_list7.split('-', 5)
            self.num_list7 = list(map(int, self.num_list7))

        if self.num_list8 is'':
            print("第八组没数据")
        else:
            self.num_list8 = self.num_list8.split('-', 5)
            self.num_list8 = list(map(int, self.num_list8))


        if self.M_list is '':
            self.e9.insert(0, "请输入M值，格式 1-2-3...")
            print("获取M 值错误")
            return None
        self.num_list = self.num_list.split('-', 5)
        try:
            self.num_list = list(map(int, self.num_list))
        except ValueError:
            self.num_list = []

        self.M_list = self.M_list.split('-', 20)
        try:
            self.M_list = list(map(int, self.M_list))
        except ValueError:
            self.M_list = []
        return True
    
    def test_BT_fun(self):
        print("test begin:")
        self.periods = 1
        for number in range(2, 177):

            number_value1 = self.test_sheet.cell_value(number, 0)
            number_value2 = self.test_sheet.cell_value(number, 1)
            # number = number + 1
            print(number_value1)
            print(number_value2)

            temp = list(map(int, re.compile(r'(10|[1-9])').findall(number_value2)))

            self.number_of_periods = int(number_value1)

            # 比价数据  是否  z  j
            self.The_winning_recognition(temp)

            # 保存 跟下一次比较
            self.The_lottery_results_old = temp
            
            self.periods = self.periods + 1   # 测试BT 专用

        self.ExcelFile_test1.save('test_BT_fun.xls')
# -----------------------------------------------------------------------------------------------------
# *****************************       数据处理layout END       ****************************************
# -----------------------------------------------------------------------------------------------------
    # 开始捕捉图 返回 识别结果
    def TestStart(self):
        img = None
        img = screen_capture(self.X_coordinate.get(), self.Y_coordinate.get(), self.Xx_coordinate.get(), self.Yy_coordinate.get())
        if self.display_img() == None :
            self.e6.delete(0, END)
            self.e6.insert(0, "请正确输入")
            logging.info("输入 坐标 错误！！")
            print("输入坐标错误")
            pass
        else:
            test_list = self.get_clean_coordinate()
            # print(test_list)
            test_list_len = len(test_list)
            if test_list_len != 10:
                print("识别错误")
                return False

            img = processor_discern(img, test_list)
            self.final_data = pytesseract.image_to_string(img)
            # print(self.final_data)

            # 在 结果 上显示
            self.e8.delete(0, END)
            self.e8.insert(0, self.final_data)

            # 显示 处理后的图片
            self.display_Endimg()

            temp = []
            temp = list(map(int, re.compile(r'(10|[1-9])').findall(self.final_data)))
            # temp 就是最后的数组
            if len(temp) < 10:
                logging.info(" 识别 数据 小于  10 ！！")
                print("识别数据小于10")
                return False
            return temp

    def start_auto_thread(self):
        if not self.auto_thread:
            self.auto_thread = threading.Thread(target=self.auto_start, args=())
            # 住线程推出的时候， 子线程也要退出。
            self.auto_thread.setDaemon(True)
            self.auto_thread.start()
        self.event.set()
        self.start_BT["state"] = DISABLED
        self.e1["state"] = DISABLED
        self.e2["state"] = DISABLED
        self.e3["state"] = DISABLED
        self.e4["state"] = DISABLED
        self.e6["state"] = DISABLED
        self.e9["state"] = DISABLED
        self.ea["state"] = DISABLED
        self.ea2["state"] = DISABLED
        self.ea3["state"] = DISABLED
        self.ea4["state"] = DISABLED
        self.ea5["state"] = DISABLED
        self.ea6["state"] = DISABLED
        self.ea7["state"] = DISABLED
        self.ea8["state"] = DISABLED
        #self.comboxlist["state"] = DISABLED
        self.num_NO_BT["state"] = DISABLED
        self.M_BT["state"] = DISABLED
        self.reality_BT["state"] = DISABLED
        self.test_BT["state"] = DISABLED

    def auto_start(self):
        logging.info("auto start run ：")
        print("auto_start run:")
        while True:
            self.event.wait()
            d_time = datetime.datetime.strptime(str(datetime.datetime.now().date())+'13:10', '%Y-%m-%d%H:%M')
            d_time1 = datetime.datetime.strptime(str(datetime.datetime.now().date())+'4:09', '%Y-%m-%d%H:%M')
            endtime = datetime.datetime.now()
            if endtime > d_time or endtime<d_time1:
                print("时间范围内：....")
                logging.info("时间 范围内 ：...")
                break
            else:
                logging.info("时间没到 ：...")
                print("时间没到 ==：....")
                time.sleep(5)

        print("开始：....")
        logging.info("start  开始：...")
        now_year = datetime.datetime.now().year
        now_hour = datetime.datetime.now().hour
        now_month = datetime.datetime.now().month
        now_days = datetime.datetime.now().day

        #  保留 确保中间 0~4点的 启动计算时间
        if (now_hour <= 4):
            if now_days == 1 and now_month in (2, 4, 6, 8,  9, 11):
                starttime = datetime.datetime.now().replace(month=(now_month - 1), day=31, hour=13, minute=4)

            elif now_days == 1 and now_month in (5, 7, 10, 12):
                starttime = datetime.datetime.now().replace(month=(now_month - 1), day=30, hour=13, minute=4)

            if now_days == 1 and now_month == 3:
                if now_year/4 == 0:
                    starttime = datetime.datetime.now().replace(month=(now_month - 1), day=29, hour=13, minute=4)
                else:
                    starttime = datetime.datetime.now().replace(month=(now_month - 1), day=28, hour=13, minute=4)

            elif now_days == 1 and now_month == 1:
                starttime = datetime.datetime.now().replace(year=(now_year - 1), month=12, day=31, hour=13, minute=4)
            else:
                starttime = datetime.datetime.now().replace(day=(now_days - 1), hour=13, minute=4)
        else:
            starttime = datetime.datetime.now().replace(hour=13, minute=4)
            
        endtime = datetime.datetime.now()
        
        count = (endtime - starttime).total_seconds() / 60
        if count > 5:
            # 保存一个 用作比较  如果第一次开奖已经到了 先获取一次
            self.The_lottery_results_old = self.TestStart()

        print("auto  start :")
        logging.info("开始 222：...")
        while True:
            now_year = datetime.datetime.now().year
            now_hour = datetime.datetime.now().hour
            now_month = datetime.datetime.now().month
            now_days = datetime.datetime.now().day
            self.event.wait()
            if (now_hour <= 4):
                if now_days == 1 and now_month in (2, 4, 6, 8, 9, 11):
                    starttime = datetime.datetime.now().replace(month=(now_month - 1), day=31, hour=13, minute=4)

                elif now_days == 1 and now_month in (5, 7, 10, 12):
                    starttime = datetime.datetime.now().replace(month=(now_month - 1), day=30, hour=13, minute=4)

                if now_days == 1 and now_month == 3:
                    if now_year / 4 == 0:
                        starttime = datetime.datetime.now().replace(month=(now_month - 1), day=29, hour=13, minute=4)
                    else:
                        starttime = datetime.datetime.now().replace(month=(now_month - 1), day=28, hour=13, minute=4)

                elif now_days == 1 and now_month == 1:
                    starttime = datetime.datetime.now().replace(year=(now_year - 1), month=12, day=31, hour=13, minute=4)

                else:
                    starttime = datetime.datetime.now().replace(day=(now_days - 1), hour=13, minute=4)

            else:
                starttime = datetime.datetime.now().replace(hour=13, minute=4)
                
            endtime = datetime.datetime.now()

            count = (endtime - starttime).total_seconds() / 60
            
            if count < 5:  # 如果时间没到 第一个， 继续等待
                time.sleep(5)
                continue

            self.The_lottery_results = self.TestStart()
            if self.The_lottery_results == False:
                time.sleep(5)
                continue

            # 不等于 表 已经更新， 可以走一下流程
            if self.The_lottery_results_old != self.The_lottery_results:
            
                self.number_of_periods = count / 5
                self.number_of_periods = int(self.number_of_periods)

                self.The_winning_recognition(self.The_lottery_results)

                # 保存 跟下一次比较
                self.The_lottery_results_old = self.The_lottery_results
                
            else:   # 还没有更新
                time.sleep(5)
# @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    def The_winning_recognition(self, results):
        print(self.number_of_periods)
        ret_val = self.get_M_num_data()
        if ret_val is None:
            print("get_M_num_data error!!")
            logging.info("get M num date error ：...")
            return

        # 11 期  下   1 号  这样的方式   0  下 号
        temp1 = self.number_of_periods % 10

        if temp1 == 0:
            temp1 = 10

        temp1 = temp1 - 1
        print("下4  名次  结果  名次上的数据  M")
        logging.info("名次  名次上的数据  名次上的数据  M")
        print(temp1+1)
        print(results)
        print(results[temp1])

        logging.info(temp1+1)
        logging.info(results)
        logging.info(results[temp1])
        
        # 保存数据上色标志  1-2-4-6-9-15-21-31-45-65-97-139-210-300-450
        marked = 0
        # print(self.comboxlist.get())
        
        if self.bet_count == len(self.M_list):
            self.bet_count = 0
            self.bet_flag = 0
            self.results_not_count = 0
        
        if self.bet_count2 == len(self.M_list):
            self.bet_count2 = 0
            self.bet_flag2 = 0
            self.results_not_count2 = 0
            
        if self.bet_count3 == len(self.M_list):
            self.bet_count3 = 0
            self.bet_flag3 = 0
            self.results_not_count3 = 0
        
        if self.bet_count4 == len(self.M_list):
            self.bet_count4 = 0
            self.bet_flag4 = 0
            self.results_not_count4 = 0

        if self.bet_count5 == len(self.M_list):
            self.bet_count5 = 0
            self.bet_flag5 = 0
            self.results_not_count5 = 0

        if self.bet_count6 == len(self.M_list):
            self.bet_count6 = 0
            self.bet_flag6 = 0
            self.results_not_count6 = 0
            
        if self.bet_count7 == len(self.M_list):
            self.bet_count7 = 0
            self.bet_flag7 = 0
            self.results_not_count7 = 0

        if self.bet_count8 == len(self.M_list):
            self.bet_count8 = 0
            self.bet_flag8 = 0
            self.results_not_count8 = 0            
        # ZJ 0 判断  =============================0000000000000000000===================
        if results[temp1] in self.num_list:
            print("Z J....！！！！！")
            logging.info("Z j  1111  1")
            marked = 1
            now_hour = datetime.datetime.now().hour
            now_minute = datetime.datetime.now().minute

            if now_hour == 3 and now_minute > 14:
               print("最后几期  不下了1， 睡觉了。。。。请关闭程序。。")
               logging.info("第一组  最后几 期  不压了 ！！ 睡觉了")
               self.not_neet_bet1 = 1
               #self.stop_thread()

            self.bet_count = 0
            self.results_not_count = 0
        else:
            self.results_not_count = self.results_not_count + 1
           
            
        # ZJ 2 判断 ===============22222222222222222222222222222222==============================================
        if isinstance(self.num_list2, str) is True:
            logging.info("第2组  为null")
            print("第2组  为null")
        elif results[temp1] in self.num_list2:
            print("Z J 2....！！！！！")
            logging.info("Z J 2....！！！！！")
            now_hour = datetime.datetime.now().hour
            now_minute = datetime.datetime.now().minute            
            if now_hour == 3 and now_minute > 14:
               print("最后几期  不下了2， 睡觉了。。。。请关闭程序。。")
               logging.info("最后几期  不下了2， 睡觉了。。。。请关闭程序。。")
               self.not_neet_bet2 = 1
               #self.stop_thread()
               
            self.bet_count2 = 0
            self.results_not_count2 = 0    
        else:
            self.results_not_count2 = self.results_not_count2 + 1

        # ZJ 3 判断      ===========================333333333333333333===============================================
        if isinstance(self.num_list3, str) is True:
            print("第3组  为null")
            logging.info("第3组  为null")
        elif results[temp1] in self.num_list3:
            print("Z J 3....！！！！！")
            logging.info("Z J 3....！！！！！")
            now_hour = datetime.datetime.now().hour
            now_minute = datetime.datetime.now().minute            
            if now_hour == 3 and now_minute > 14:
               print("最后几期  不下了3， 睡觉了。。。。请关闭程序。。")
               logging.info("最后几期  不下了3， 睡觉了。。。。请关闭程序。。")
               self.not_neet_bet3 = 1
               #self.stop_thread()
               
            self.bet_count3 = 0     
            self.results_not_count3 = 0
        else:
            self.results_not_count3 = self.results_not_count3 + 1
            
        # ZJ 4 判断      ===========================44444444444444444===============================================
        if isinstance(self.num_list4, str) is True:
            print("第4组  为null")
            logging.info("第4组  为null")
        elif results[temp1] in self.num_list4:
            print("Z J 4....！！！！！")
            logging.info("Z J 4....！！！！！")
            now_hour = datetime.datetime.now().hour
            now_minute = datetime.datetime.now().minute            
            if now_hour == 3 and now_minute > 14:
               print("最后几期  4  不下了， 睡觉了。。。。请关闭程序。。")
               logging.info("最后几期  4  不下了， 睡觉了。。。。请关闭程序。。")
               self.not_neet_bet4 = 1
               #self.stop_thread()
               
            self.bet_count4 = 0     
            self.results_not_count4 = 0           
        else:
            self.results_not_count4 = self.results_not_count4 + 1
 
        # ZJ 5 判断      ===========================555555555555555===============================================
        if isinstance(self.num_list5, str) is True:
            print("第5组  为null")
            logging.info("第5组  为null")
        elif results[temp1] in self.num_list5:
            print("Z J 5....！！！！！")
            logging.info("Z J 5....！！！！！")
            now_hour = datetime.datetime.now().hour
            now_minute = datetime.datetime.now().minute            
            if now_hour == 3 and now_minute > 14:
               print("最后几期  不下了5， 睡觉了。。。。请关闭程序。。")
               logging.info("最后几期  不下了3， 睡觉了。。。。请关闭程序。。")
               self.not_neet_bet5 = 1
               #self.stop_thread()
               
            self.bet_count5 = 0     
            self.results_not_count5 = 0
        else:
            self.results_not_count5 = self.results_not_count5 + 1
            
        # ZJ 6 判断      ===========================666666666666666666===============================================
        if isinstance(self.num_list6, str) is True:
            print("第6组  为null")
            logging.info("第6组  为null")
        elif results[temp1] in self.num_list6:
            print("Z J 6....！！！！！")
            logging.info("Z J 6....！！！！！")
            now_hour = datetime.datetime.now().hour
            now_minute = datetime.datetime.now().minute            
            if now_hour == 3 and now_minute > 14:
               print("最后几期  不下了6， 睡觉了。。。。请关闭程序。。")
               logging.info("最后几期  不下了6， 睡觉了。。。。请关闭程序。。")
               self.not_neet_bet6 = 1
               
            self.bet_count6 = 0     
            self.results_not_count6 = 0
        else:
            self.results_not_count6 = self.results_not_count6 + 1

        # ZJ 7 判断      ===========================77777777777777777===============================================
        if isinstance(self.num_list7, str) is True:
            print("第7组  为null")
            logging.info("第7组  为null")
        elif results[temp1] in self.num_list7:
            print("Z J 7....！！！！！")
            logging.info("Z J 7....！！！！！")
            now_hour = datetime.datetime.now().hour
            now_minute = datetime.datetime.now().minute            
            if now_hour == 3 and now_minute > 14:
               print("最后几期  不下了7， 睡觉了。。。。请关闭程序。。")
               logging.info("最后几期  不下了7， 睡觉了。。。。请关闭程序。。")
               self.not_neet_bet7 = 1
               #self.stop_thread()
               
            self.bet_count7 = 0     
            self.results_not_count7 = 0
        else:
            self.results_not_count7 = self.results_not_count7 + 1

        # ZJ 8 判断      ===========================888888888888888888===============================================
        if isinstance(self.num_list8, str) is True:
            print("第8组  为null")
            logging.info("第8组  为null")
        elif results[temp1] in self.num_list8:
            print("Z J 8....！！！！！")
            logging.info("Z J 8....！！！！！")
            now_hour = datetime.datetime.now().hour
            now_minute = datetime.datetime.now().minute            
            if now_hour == 3 and now_minute > 14:
               print("最后几期  不下了8， 睡觉了。。。。请关闭程序。。")
               logging.info("最后几期  不下了8， 睡觉了。。。。请关闭程序。。")
               self.not_neet_bet8 = 1
               #self.stop_thread()
               
            self.bet_count8 = 0     
            self.results_not_count8 = 0
        else:
            self.results_not_count8 = self.results_not_count8 + 1            
            
 #================================================================================           
        temp1 = (self.number_of_periods % 10) + 1   # 因为是压下 一期的

        if self.M_list[self.bet_count] != 0 and self.not_neet_bet1 != 1:
            auto_click.reality_bet(auto_click, temp1, self.num_list, self.M_list[self.bet_count], self.aotu_click_sheet)
            self.worksheet.write((self.periods+1), 3, ("%s") % (self.M_list[self.bet_count]), self.style)  # 期数
        else:
            print("这期不用压！！！！1")
            logging.info("这期不用压！！！！1")
        

        if self.M_list[self.bet_count2] != 0 and self.not_neet_bet2 != 1:
            auto_click.reality_bet(auto_click, temp1, self.num_list2, self.M_list[self.bet_count2], self.aotu_click_sheet)
            self.worksheet.write((self.periods+1), 4, ("%s") % (self.M_list[self.bet_count2]), self.style)  # 期数
        else:
            print("这期不用压2！！！！1")
            logging.info("这期不用压2！！！！1")

        if self.M_list[self.bet_count3] != 0 and self.not_neet_bet3 != 1:
            auto_click.reality_bet(auto_click, temp1, self.num_list3, self.M_list[self.bet_count3], self.aotu_click_sheet)
            self.worksheet.write((self.periods+1), 5, ("%s") % (self.M_list[self.bet_count3]), self.style)  # 期数
        else:
            print("这期不用压3！！！！1")
            logging.info("这期不用压3！！！！1")

        if self.M_list[self.bet_count4] != 0 and self.not_neet_bet4 != 1:
            auto_click.reality_bet(auto_click, temp1, self.num_list4, self.M_list[self.bet_count4], self.aotu_click_sheet)
            self.worksheet.write((self.periods+1), 6, ("%s") % (self.M_list[self.bet_count4]), self.style)  # 期数
        else:
            print("这期不用压4！！！！1")
            logging.info("这期不用压4！！！！1")

        if self.M_list[self.bet_count5] != 0 and self.not_neet_bet5 != 1:
            auto_click.reality_bet(auto_click, temp1, self.num_list5, self.M_list[self.bet_count5], self.aotu_click_sheet)
            self.worksheet.write((self.periods+1), 7, ("%s") % (self.M_list[self.bet_count5]), self.style)  # 期数
        else:
            print("这期不用压5！！！！1")
            logging.info("这期不用压5！！！！1")
          
        if self.M_list[self.bet_count6] != 0 and self.not_neet_bet6 != 1:
            auto_click.reality_bet(auto_click, temp1, self.num_list6, self.M_list[self.bet_count6], self.aotu_click_sheet)
            self.worksheet.write((self.periods+1), 8, ("%s") % (self.M_list[self.bet_count6]), self.style)  # 期数
        else:
            print("这期不用压6！！！！1")
            logging.info("这期不用压6！！！！1")
            
        if self.M_list[self.bet_count7] != 0 and self.not_neet_bet7 != 1:
            auto_click.reality_bet(auto_click, temp1, self.num_list7, self.M_list[self.bet_count7], self.aotu_click_sheet)
            self.worksheet.write((self.periods+1), 9, ("%s") % (self.M_list[self.bet_count7]), self.style)  # 期数
        else:
            print("这期不用压7！！！！1")
            logging.info("这期不用压7！！！！1")
            
        if self.M_list[self.bet_count8] != 0 and self.not_neet_bet8 != 1:
            auto_click.reality_bet(auto_click, temp1, self.num_list8, self.M_list[self.bet_count8], self.aotu_click_sheet)
            self.worksheet.write((self.periods+1), 10, ("%s") % (self.M_list[self.bet_count8]), self.style)  # 期数
        else:
            print("这期不用压8！！！！1")
            logging.info("这期不用压8！！！！1")            
        # 保存原始数据
        self.save_excel(marked)

        print("now00 没z 计数 %d 压入%d=====" % (self.results_not_count, self.M_list[self.bet_count]))
        print("now2 没z 计数 %d  压入%d=======" % (self.results_not_count2, self.M_list[self.bet_count2]))
        print("now3 没z 计数 %d  压入%d=======" % (self.results_not_count3, self.M_list[self.bet_count3]))
        print("now4 没z 计数 %d  压入%d=======" % (self.results_not_count4, self.M_list[self.bet_count4]))
        print("now5 没z 计数 %d 压入%d=====" % (self.results_not_count5, self.M_list[self.bet_count5]))
        print("now6 没z 计数 %d  压入%d=======" % (self.results_not_count6, self.M_list[self.bet_count6]))
        print("now7 没z 计数 %d  压入%d=======" % (self.results_not_count7, self.M_list[self.bet_count7]))
        print("now8 没z 计数 %d  压入%d=======" % (self.results_not_count8, self.M_list[self.bet_count8]))
        logging.info("now00 没z 计数 %d 压入%d=====" % (self.results_not_count, self.M_list[self.bet_count]))
        logging.info("now2 没z 计数 %d  压入%d=======" % (self.results_not_count2, self.M_list[self.bet_count2]))
        logging.info("now3 没z 计数 %d  压入%d=======" % (self.results_not_count3, self.M_list[self.bet_count3]))
        logging.info("now4 没z 计数 %d  压入%d=======" % (self.results_not_count4, self.M_list[self.bet_count4]))
        logging.info("now5 没z 计数 %d 压入%d=====" % (self.results_not_count5, self.M_list[self.bet_count5]))
        logging.info("now6 没z 计数 %d  压入%d=======" % (self.results_not_count6, self.M_list[self.bet_count6]))
        logging.info("now7 没z 计数 %d  压入%d=======" % (self.results_not_count7, self.M_list[self.bet_count7]))
        logging.info("now8 没z 计数 %d  压入%d=======" % (self.results_not_count8, self.M_list[self.bet_count8]))        


        self.bet_count = self.bet_count + 1
        self.bet_count2 = self.bet_count2 + 1
        self.bet_count3 = self.bet_count3 + 1
        self.bet_count4 = self.bet_count4 + 1
        self.bet_count5 = self.bet_count5 + 1
        self.bet_count6 = self.bet_count6 + 1
        self.bet_count7 = self.bet_count7 + 1
        self.bet_count8 = self.bet_count8 + 1        
# @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

    def display_img(self):
        self.frm_p = Frame(self.frm)
        try:
            self.im = Image.open(r"E:\xiaotuzi\5.jpg")
        except OSError:
            print("读取截图，图片错误")
            logging.info("读取截图，图片错误")
            pass
            return None
        else:
            self.img = ImageTk.PhotoImage(self.im)
            self.imLabel = Label(self.frm_p, image=self.img).grid(row=0, column=0)
            self.frm_p.place(x=0, y=0)
            return True


    def display_Endimg(self):
        self.frm_p1 = Frame(self.frm)
        try:
            self.im1 = Image.open(r"E:\xiaotuzi\66.jpg")
        except :
            print("读取图片错误")
            pass
            return None

        self.img1 = ImageTk.PhotoImage(self.im1)
        self.imLabel = Label(self.frm_p1, image=self.img1).grid(row=0, column=0)
        self.frm_p1.place(x=0, y=100)

    def start_display_xy_thread(self):
        self.thread = threading.Thread(target=self.play_coordinate, args=())
        # 住线程推出的时候， 子线程也要退出。
        self.thread.setDaemon(True)
        self.thread.start()


    def close(self):
        print("all close ...")
        logging.info("all close ...")
        self.device.close()

    def play_coordinate(self):
        while True:
            x, y = get_mouse_point()
            self.Temp = ("%s .%s") % (x, y)
            self.e7.delete(0, END)
            self.e7.insert(0, self.Temp)
            time.sleep(0.1)


    def save_excel(self, marked):
        self.periods = self.periods+1

        now_hour = datetime.datetime.now().hour
        now_days = datetime.datetime.now().day

        if (now_hour <= 4):
            starttime = datetime.datetime.now().replace(day=(now_days - 1), hour=13, minute=4)
        else:
            starttime = datetime.datetime.now().replace(hour=13, minute=4)

        endtime = datetime.datetime.now()
        count = (endtime - starttime).total_seconds() / 60
        if count < 1:
            print("Not billing time")
            return

        count = count / 5
        count = int(count)
        print(count)
        if marked ==1:
            self.worksheet.write(self.periods, 0,("%s")%count, self.styleBlueBkg)   # 期数
            self.worksheet.write(self.periods, 1, self.final_data, self.styleBlueBkg)  # 号码
        else:
            self.worksheet.write(self.periods, 0,("%s")%count, self.style)   # 期数
            self.worksheet.write(self.periods, 1, self.final_data, self.style)  # 号码

        self.workbook.save(("%s.xls") % time.strftime('%Y-%m-%d', time.localtime(time.time())))
        # print("save excel  ok   !!!!")

    def excel_test_init(self):
        self.ExcelFile_test = xlrd.open_workbook(r'E:\xiaotuzi\test_BT_fun.xls')
        # sheet = ExcelFile_test.add_sheet('2019-07-13',cell_overwrite_ok=True)
        self.Excel_test = self.ExcelFile_test.sheet_names()
        self.test_sheet = self.ExcelFile_test.sheet_by_name(self.Excel_test[0])
        self.test_sheet_data = self.ExcelFile_test.sheet_by_name(self.Excel_test[0])
        
        self.ExcelFile_test1 = copy(self.ExcelFile_test)
        self.test_sheet1 = self.ExcelFile_test1.get_sheet(0)

    def excel_auto_click_init(self):
        self.ExcelFile_auto_click = xlrd.open_workbook(r'E:\xiaotuzi\auto_click.xls')
        self.aotu_click = self.ExcelFile_auto_click.sheet_names()
        self.aotu_click_sheet = self.ExcelFile_auto_click.sheet_by_name(self.aotu_click[0])
        self.aotu_click_sheet_data = self.ExcelFile_auto_click.sheet_by_name(self.aotu_click[0])


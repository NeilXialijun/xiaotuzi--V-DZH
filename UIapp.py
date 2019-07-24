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
# import uuid
#import multiprocessing

class Application(object):

    def __init__(self, master=None):
        self.root = master    #    定义内部变量root
        self.root.geometry('660x410')
        self.root.resizable(0,0)

        # # 创建一个容器,
        self.frm = ttk.LabelFrame(self.root, text="图片处理")     # 创建一个容器，其父容器为win
        self.frm.place(x=0, y=0, anchor="nw", width=650, height=250)
        self.frm_p = Frame(self.frm)

        self.num_f = ttk.LabelFrame(self.root, text="数据处理")     # 创建一个容器，其父容器为win
        self.num_f.place(x=0, y=255, anchor="nw", width=650, height=150)
        self.frm_num = Frame(self.num_f)

        self.clean_coordinate_str = ""
        self.final_data = ""
        self.results_not_count = 0
        self.bet_count = 0
        self.bet_flag = 0
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
        self.Label1 = Label(self.frm_x, text='X坐标开始:').grid(row=0, column=0)
        self.frm_x.place(x=0, y=50)

        self.frm_y = Frame(self.frm)
        self.Label2 = Label(self.frm_y, text='Y坐标开始:').grid(row=0, column=2)
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
        self.Labe1 = Label(self.frm_xx, text='X坐标终点:', justify=LEFT).grid(row=0, column=4)
        self.Labe2 = Label(self.frm_yy, text='Y坐标终点:', justify=LEFT).grid(row=0, column=6)
        self.frm_xx.place(x=300, y=50)
        self.frm_yy.place(x=400, y=50)

        self.Xx_coordinate = IntVar()
        self.e3 = Entry(self.frm_xx, width=5, textvariable=self.Xx_coordinate, justify=LEFT)    # Entry 是 Tkinter 用来接收字符串等输入的控件.
        self.e3.grid(row=0, column=5, padx=1, pady=1)  # 设置输入框显示的位置，以及长和宽属性

        self.Yy_coordinate = IntVar()
        self.e4 = Entry(self.frm_yy, width=5, textvariable=self.Yy_coordinate, justify=LEFT)
        self.e4.grid(row=0, column=7, padx=1, pady=1)

        self.test_BT = Frame(self.frm)
        Button(self.test_BT, text='开始测试', width=10, command=self.TestStart) \
            .grid(row=0, column=4, padx=1, pady=1)
        self.test_BT.place(x=550, y=25)

        self.start_BT_F = Frame(self.frm)
        self.start_BT = Button(self.start_BT_F, text='停止铺抓', width=10, command=self.stop_thread)
        self.start_BT .grid(row=0, column=4, padx=1, pady=1)
        self.start_BT_F.place(x=550, y=80)


        self.start_BT_F = Frame(self.frm)
        self.start_BT = Button(self.start_BT_F, text='开始铺抓', width=10, command=self.start_auto_thread)
        self.start_BT .grid(row=0, column=4, padx=1, pady=1)
        self.start_BT_F.place(x=550, y=150)


# # -------------清理条 输入 ---------------------------------------------------------
        self.clean_pixel_f = Frame(self.frm)
        self.clean_pixel_l = Label(self.clean_pixel_f, text='清理坐标:', justify=LEFT).grid(row=0, column=11)

        self.clean_pixel = StringVar()
        self.e6 = Entry(self.clean_pixel_f, width=50, textvariable=self.clean_pixel, justify=LEFT)
        self.e6.grid(row=0, column=12, padx=1, pady=1)
        self.clean_pixel_f.place(x=50, y=150)

# # -------------鼠标 坐标------------------------------------------------------------------
        self.mouse_f = Frame(self.frm)
        self.mouse_l = Label(self.mouse_f, text='坐标:', justify=LEFT).grid(row=0, column=11)

        self.set_xy = StringVar()
        self.Temp = ("%s.%s") % get_mouse_point()
        self.e7 = Entry(self.mouse_f, width=10, textvariable=self.Temp)
        self.e7.grid(row=0, column=12, padx=1, pady=1)
        self.mouse_f.place(x=520, y=200)


# # -------------识别结果------------------------------------------------------------------
        self.recognition_f = Frame(self.frm)
        self.recognition_l = Label(self.recognition_f, text='识别结果:', justify=LEFT).grid(row=0, column=11)

        self.e8 = Entry(self.recognition_f, width=50, textvariable=self.final_data)
        self.e8 .grid(row=0, column=12, padx=1, pady=1)
        self.recognition_f.place(x=50, y=180)
# # -----------------------------------------------------------------------------------------------------
# # ****************************     数据处理layout    **************************************************
# # -----------------------------------------------------------------------------------------------------
        self.M_enter_f = Frame(self.num_f)
        self.M_enter_l = Label(self.M_enter_f, text='M 输 入 :', justify=LEFT).grid(row=0, column=0)

        self.M_enter_s = StringVar()
        self.e9 = Entry(self.M_enter_f, width=50, textvariable=self.M_enter_s)
        self.e9.grid(row=0, column=12, padx=1, pady=1)
        self.M_enter_f.place(x=10, y=10)

        self.num_enter_f = Frame(self.num_f)
        self.num_enter_l = Label(self.num_enter_f, text='号码输入:', justify=LEFT).grid(row=0, column=0)

        self.num_enter_s = StringVar()
        self.ea = Entry(self.num_enter_f, width=50, textvariable=self.num_enter_s)
        self.ea.grid(row=0, column=12, padx=1, pady=1)
        self.num_enter_f.place(x=10, y=40)

        self.num_NO_f  = Frame(self.num_f)
        Button(self.num_NO_f, text='名号测试', width=10, command=lambda: auto_click.NO_Num_BT(auto_click, self.aotu_click_sheet)) \
            .grid(row=0, column=4, padx=1, pady=1)
        # Button(self.num_NO_f, text='名号测试', width=10, command='') \
        #     .grid(row=0, column=4, sticky=E, padx=1, pady=1)
        self.num_NO_f.place(x=50, y=85)

        self.M_f = Frame(self.num_f)
        Button(self.M_f, text='M 测试', width=10, command=lambda: auto_click.M_BT(auto_click, self.aotu_click_sheet)) \
            .grid(row=0, column=4, padx=1, pady=1)
        self.M_f.place(x=150, y=85)

        self.M_list = self.M_enter_s.get()
        self.reality_f = Frame(self.num_f)
        Button(self.reality_f, text='真实模拟测试', width=10, command=lambda: auto_click.reality_BT(auto_click, self.aotu_click_sheet)) \
            .grid(row=0, column=4, padx=1, pady=1)
        self.reality_f.place(x=250, y=85)

        self.periods_select = Frame(self.num_f)
        self.num_enter_l = Label(self.periods_select, text='周期间隔 :').grid(row=0, column=0)

        self.comvalue = StringVar()  # 窗体自带的文本，新建一个值
        self.comboxlist = ttk.Combobox(self.periods_select, width=3, textvariable=self.comvalue)  # 初始化
        self.comboxlist["values"] = (0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10)
        self.comboxlist.current(6)  # 选择第一个
        self.comboxlist.grid(row=0, column=6)
        self.periods_select.place(x=450, y=8)

        # temp2 = [5,2,6,3,10,4,1,9,8,7]
        # self.number_of_periods = 140
        self.test_f = Frame(self.num_f)
        Button(self.test_f, text='暂时测试', width=10, command=lambda: self.test_BT_fun()) \
            .grid(row=0, column=4, padx=1, pady=1)
        self.test_f.place(x=550, y=85)

    def stop_thread(self):
        # _async_raise(self.auto_thread.ident, SystemExit)
        # ._Thread__stop()
        # pinger_instance.self.auto_thread.is_set()
        # self.auto_thread.join()
        #top_thread(self.auto_thread)
        print("stop thread")
        self.event.clear()
        self.start_BT['state'] = "normal"


    def get_M_num_data(self):
        self.num_list = self.num_enter_s.get()
        self.M_list = self.M_enter_s.get()
        if self.num_list is '':
            self.ea.insert(0, "请输入号码，格式 1-2-3...")
            return None

        if self.M_list is '':
            self.e9.insert(0, "请输入M值，格式 1-2-3...")
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
        # print(self.M_list)
        # print(self.num_list)
    
    def test_BT_fun(self):
        print("test begin:")
        self.periods = 13
        for number in range(14, 88):

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
            return temp

    def start_auto_thread(self):
        if not self.auto_thread:
            self.auto_thread = threading.Thread(target=self.auto_start, args=())
            # 住线程推出的时候， 子线程也要退出。
            self.auto_thread.setDaemon(True)
            self.auto_thread.start()
        self.event.set()
        self.start_BT["state"] = DISABLED

    def auto_start(self):

        print("auto_start run:")
        while True:
            self.event.wait()
            d_time = datetime.datetime.strptime(str(datetime.datetime.now().date())+'13:10', '%Y-%m-%d%H:%M')
            d_time1 = datetime.datetime.strptime(str(datetime.datetime.now().date())+'4:09', '%Y-%m-%d%H:%M')
            endtime = datetime.datetime.now()
            if endtime > d_time or endtime<d_time1:
                print("时间范围内：....")
                break
            else:
                print("时间没到 ==：....")
                time.sleep(5)

        print("开始：....")
        now_hour = datetime.datetime.now().hour
        now_days = datetime.datetime.now().day

        #  保留 确保中间 0~4点的 启动 计算时间
        if (now_hour <= 4):
            starttime = datetime.datetime.now().replace(day=(now_days - 1), hour=13, minute=4)
        else:
            starttime = datetime.datetime.now().replace(hour=13, minute=4)
            
        endtime = datetime.datetime.now()
        
        count = (endtime - starttime).total_seconds() / 60
        if count > 5:
            # 保存一个 用作比较  如果第一次开奖已经到了 先获取一次
            self.The_lottery_results_old = self.TestStart()

        print("auto  start :")
        while True:
            now_hour = datetime.datetime.now().hour
            now_days = datetime.datetime.now().day
            self.event.wait()
            if (now_hour <= 4):
                starttime = datetime.datetime.now().replace(day=(now_days - 1), hour=13, minute=4)
            else:
                starttime = datetime.datetime.now().replace(hour=13, minute=4)
                
            endtime = datetime.datetime.now()

            count = (endtime - starttime).total_seconds() / 60
            
            if count < 5:  # 如果时间没到 第一个， 继续等待
                time.sleep(5)
                continue

            self.The_lottery_results = self.TestStart()
            
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
            return

        # 11 期  下   1 号  这样的方式   0  下 号
        temp1 = self.number_of_periods % 10

        if temp1 == 0:
            temp1 = 10

        temp1 = temp1 - 1
        print("下4  名次  结果  名次上的数据  M")
        print(temp1)
        print(results)
        print(results[temp1])
        print(self.num_list)
        
        # 保存数据上色标志
        marked = 0

         # ZJ 判断
        if results[temp1] in self.num_list:
            print("Z J....！！！！！")
            marked = 1
            
            if self.bet_flag == 1:
                self.bet_count = 0
                self.bet_flag = 0
            self.results_not_count = 0
        else:
            self.results_not_count = self.results_not_count + 1
            if self.results_not_count >= int(self.comboxlist.get()) or self.bet_flag == 1:
                self.bet_flag = 1
                temp1 = (self.number_of_periods % 10) + 1   # 因为是压下 一期的
                
                auto_click.reality_bet(auto_click, temp1, self.num_list, self.M_list[self.bet_count], self.aotu_click_sheet)
                self.worksheet.write((self.periods+1), 3, ("%s")%(self.M_list[self.bet_count]), self.style)  # 期数

                # self.test_sheet1.write(self.periods, 3, ("%s") % (self.M_list[self.bet_count]), self.style)  # 测试BT专用

                print("投入M %d  。。。" %self.M_list[self.bet_count])
                self.bet_count = self.bet_count + 1
                print("bet  on    ....！！！！！")

        if self.bet_count > len(self.M_list):
            self.bet_count = 0
            self.bet_flag = 0

        temp1 = len(self.M_list) + int(self.comboxlist.get())
        if self.results_not_count == temp1:
            self.results_not_count = 0

        # 保存原始数据，
        self.save_excel(marked)

        print("now 没z 计数 %d   ====================="%self.results_not_count)
# @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

    def display_img(self):
        self.frm_p = Frame(self.frm)
        try:
            self.im = Image.open(r"E:\xiaotuzi\5.jpg")
        except OSError:
            print("读取截图，图片错误")
            pass
            return None
        else:
            self.img = ImageTk.PhotoImage(self.im)
            self.imLabel = Label(self.frm_p, image=self.img).grid(row=0, column=0)
            self.frm_p.place(x=0, y=0)
            return True


    def display_Endimg(self):
        self.frm_p1 = Frame(self.frm)
        self.im1 = Image.open(r"E:\xiaotuzi\66.jpg")
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


import win32api
import win32con
import win32gui
from ctypes import *
import time

class POINT(Structure):
    _fields_ = [("x", c_ulong),("y", c_ulong)]

def get_mouse_pinrt():
    po = POINT()
    windll.user32.GetCursorPos(byref(po))
    print(int(po.x), int(po.y))
    time.sleep(0.2)
    return int(po.x), int(po.y)

def get_mouse_point():
    po = POINT()
    windll.user32.GetCursorPos(byref(po))
    return int(po.x), int(po.y)


def mouse_click(x=None,y=None):
    if not x is None and not y is None:
        mouse_move(x,y)
        time.sleep(0.05)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)

def mouse_dclick(x=None,y=None):
    if not x is None and not y is None:
        mouse_move(x,y)
        time.sleep(0.05)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)

def mouse_move(x,y):
    windll.user32.SetCursorPos(x, y)
    
def key_input(str=''):
    for c in str:
        win32api.keybd_event(VK_CODE[c],0,0,0)
        win32api.keybd_event(VK_CODE[c],0,win32con.KEYEVENTF_KEYUP,0)
        time.sleep(0.01)
             
class auto_click():
    def excel_init(self):
        #文件打开
        try :
            ExcelFile=xlrd.open_workbook(r'E:\xiaotuzi\auto_click.xls')
        except:
            print("no more aotu_click date1")
        print(ExcelFile.sheet_names())
        # 点击 坐标打开
        try :
            self.aotu_click1 = ExcelFile.sheet_by_name(r'E:\xiaotuzi\auto_click.xls')
        except :
            print("no more aotu_click date1")
        
        try :
            self.aotu_click2 = ExcelFile.sheet_by_name(r'E:\xiaotuzi\auto_click.xls')
            
        except :
            print("no more aotu_click date2")

        # 数据 记录
        self.data_logging = ExcelFile.sheet_by_name('data_logging')

    def M_number_Enter_all(self, M_Enter, aotu_click):
        if M_Enter >= 1000:
            M_temp_q = int(M_Enter/1000)
            M_temp_b = int((M_Enter%1000)/100)
            M_temp_s = int((M_Enter%100)/10)
            M_temp_g = int(M_Enter%10)
            time.sleep(0.6)
            self.M_number_Enter(self, M_temp_q, aotu_click)
            time.sleep(0.6)
            self.M_number_Enter(self, M_temp_b, aotu_click)
            time.sleep(0.6)
            self.M_number_Enter(self, M_temp_s, aotu_click)
            time.sleep(0.6)
            self.M_number_Enter(self, M_temp_g, aotu_click)

        elif M_Enter >= 100:
            M_temp_b = int((M_Enter%1000)/100)
            M_temp_s = int((M_Enter%100)/10)
            M_temp_g = int(M_Enter%10)
            time.sleep(0.6)
            self.M_number_Enter(self, M_temp_b, aotu_click)
            time.sleep(0.6)
            self.M_number_Enter(self, M_temp_s, aotu_click)
            time.sleep(0.6)
            self.M_number_Enter(self, M_temp_g, aotu_click)

        elif M_Enter >= 10:
            M_temp_s = int((M_Enter%100)/10)
            M_temp_g = int(M_Enter%10)

            self.M_number_Enter(self, M_temp_s, aotu_click)
            time.sleep(0.6)
            self.M_number_Enter(self, M_temp_g, aotu_click)
            time.sleep(0.6)

        else:
            time.sleep(0.6)
            self.M_number_Enter(self, M_Enter, aotu_click)


    def M_number_Enter(self, M_Enter, aotu_click):
        if M_Enter >= 0 and M_Enter < 11:
            if M_Enter == 0:
                M_Enter = 10  # 0坐标 的数据在 最后一个
            # 获取名次坐标
            ranking_data = aotu_click.cell_value(3, M_Enter)

            # 转换成坐标数据
            list1 = ranking_data.split(' ', 1)
            list1 = list(map(int, list1))
            #print(list1)
            mouse_click(list1[0], list1[1])

        # 选 号 1~ 10
    def number_enter_for_list(self, Num_Enter_List, aotu_click):

        for i in Num_Enter_List:
            # 获取名次坐标
            ranking_data = aotu_click.cell_value(1, i)
            # print(ranking_data)

            # 转换成坐标数据
            list1 = ranking_data.split(' ', 1)
            list1 = list(map(int, list1))
            #print(list1)
            mouse_click(list1[0], list1[1])
            time.sleep(0.8)

    # 选 号 1~ 10
    def number_enter(self, Num_Enter, aotu_click):
        if Num_Enter > 0 and Num_Enter < 11:
           
            # 获取名次坐标
            ranking_data = aotu_click.cell_value(1, Num_Enter)
            # print(ranking_data)

            # 转换成坐标数据
            list1 = ranking_data.split(' ', 1)
            list1 = list(map(int, list1))
            #print(list1)
            mouse_click(list1[0], list1[1])

        # 名次 选中  冠  ~   10


    def ranking_selection(self, ranking, aotu_click_sheet):
        if ranking > 0 and ranking < 11:
            # 获取名次坐标
            ranking_data = aotu_click_sheet.cell_value(0, ranking)
            # print(ranking_data)

            # 转换成坐标数据
            list1 = ranking_data.split(' ', 1)
            list1 = list(map(int, list1))
            #print(list1)
            mouse_click(list1[0], list1[1])

        # M 输入 选中


    def clean_all(self, aotu_click_sheet):
        ranking_data = aotu_click_sheet.cell_value(6, 1)
        # 转换成坐标数据
        list1 = ranking_data.split(' ', 1)
        list1 = list(map(int, list1))
        #print(list1)
        mouse_click(list1[0], list1[1])


    # M 下zhu
    def Down_enterM(self, aotu_click_sheet):
        ranking_data = aotu_click_sheet.cell_value(4, 1)
        # 转换成坐标数据
        list1 = ranking_data.split(' ', 1)
        list1 = list(map(int, list1))
        #print(list1)
        mouse_click(list1[0], list1[1])


    # M 下  最后的确认
    def Down_confirmM(self, aotu_click_sheet):
        ranking_data = aotu_click_sheet.cell_value(5, 1)
        # 转换成坐标数据
        list1 = ranking_data.split(' ', 1)
        list1 = list(map(int, list1))
        #print(list1)
        mouse_click(list1[0], list1[1])


    # M 输入 选中
    def select_enterM(self, aotu_click_sheet):
        ranking_data = aotu_click_sheet.cell_value(2, 1)
        # 转换成坐标数据
        list1 = ranking_data.split(' ', 1)
        list1 = list(map(int, list1))
        #print(list1)
        mouse_click(list1[0], list1[1])


    # 名次 测试
    def select_enterM_test(self, aotu_click_sheet):
        number = 1
        for i in range(10):
            number_selection_data = aotu_click_sheet.cell_value(0, number)
            number = number + 1
            if number_selection_data != '':
                list1 = number_selection_data.split(' ', 1)
                list1 = list(map(int, list1))
                print(list1)
                mouse_click(list1[0], list1[1])
                time.sleep(0.3)
            else:
                print("%d is None" % i)
                return


    # 选 号 测试
    def number_selection_test(self, aotu_click):
        number = 1
        for i in range(10):
            number_selection_data = aotu_click.cell_value(1, number)
            number = number + 1
            if number_selection_data != '':
                list1 = number_selection_data.split(' ', 1)
                list1 = list(map(int, list1))
                print(list1)
                mouse_click(list1[0], list1[1])
                time.sleep(0.3)

            else:
                print("%d is None" % i)
                return

    # M 号码 测试
    def M_number_test(self, aotu_click):
        number = 1
        for i in range(10):
            number_selection_data = aotu_click.cell_value(3, number)
            number = number + 1
            if number_selection_data != '':
                list1 = number_selection_data.split(' ', 1)
                list1 = list(map(int, list1))
                print(list1)
                mouse_click(list1[0], list1[1])
                time.sleep(0.3)

            else:
                print("%d is None" % i)
                return

    def NO_Num_BT(self, aotu_click):
        self.number_selection_test(self, aotu_click)
        self.select_enterM_test(self, aotu_click)
        
    def M_BT(self, aotu_click):
        self.select_enterM(self,aotu_click)
        time.sleep(0.5)
        self.M_number_test(self, aotu_click)

    # 入 M  整套 测试    
    def reality_BT(self, aotu_click):
        self.ranking_selection(self, 1, aotu_click)
        time.sleep(0.5)
        self.number_enter(self, 2, aotu_click)
        time.sleep(0.5)
        self.select_enterM(self, aotu_click)
        time.sleep(0.8)
        # self.M_number_Enter(self, 10, aotu_click)
        self.M_number_Enter_all(self, 1, aotu_click)
        time.sleep(0.5)
        self.Down_enterM(self, aotu_click)
        time.sleep(0.5)
        self.Down_confirmM(self, aotu_click)
        
    def reality_bet(self, ranking, number, M, aotu_click):
        self.ranking_selection(self, ranking, aotu_click)
        time.sleep(1)
        self.number_enter_for_list(self, number, aotu_click)
        time.sleep(1)
        self.select_enterM(self, aotu_click)
        time.sleep(1)

        self.M_number_Enter_all(self, M, aotu_click)
        time.sleep(1)
        self.Down_enterM(self, aotu_click)
        time.sleep(1)
        self.Down_confirmM(self, aotu_click)
        time.sleep(5)
        self.clean_all(self, aotu_click)

from tkinter import *
from tkinter.messagebox import *
from UIapp import *
# import xlwt
import xlrd
import uuid
# from MainPage import *
  
class LoginPage(object): 
	def __init__(self, master=None): 
		self.root = master #定义内部变量root 
		self.root.geometry('%dx%d' % (300, 180)) #设置窗口大小 
		self.username = StringVar() 
		self.password = StringVar()
		self.createPage() 
		
		# 10 进制  显示 mac
		self.username.set(self.get_mac_address())
		
		return_value = self.aotu_verify()
		
		if return_value == True:
			print("123")
			self.ExcelFile_auto_click.release_resources()
			Application(self.root)
			# MainPage(self.root) 
			self.page.destroy()
			
	def createPage(self): 
		self.page = Frame(self.root) #创建Frame 
		self.page.pack() 
		Label(self.page).grid(row=0, stick=W) 
		Label(self.page, text = '机械编码: ').grid(row=1, stick=W, pady=10) 
		Entry(self.page, textvariable=self.username).grid(row=1, column=1, stick=E) 
		
		Label(self.page, text = '秘钥输入: ').grid(row=2, stick=W, pady=10) 
		Entry(self.page, textvariable=self.password).grid(row=2, column=1, stick=E) 
		Button(self.page, text='确认', command=self.loginCheck).grid(row=3, stick=W, pady=10) 
		Button(self.page, text='退出', command=self.page.quit).grid(row=3, column=1, stick=E) 


	def loginCheck(self): 
		secret = self.password.get() 
		print(secret)
		return_value = self.Encryption_algorithms(secret)
		if return_value == True:
			self.ExcelFile_auto_click.release_resources()
			Application(self.root)
			# MainPage(self.root) 
			self.page.destroy()  
		else: 
			showinfo(title='错误', message='输入错误！  请从新输入！！') 
			
	# 获取mic 
	def get_mac_address(self): 
		mac=uuid.UUID(int = uuid.getnode()).hex[-12:]
		mac = int(mac, 16)
		return mac

	def Encryption_algorithms(self, enter):
		my_mac = self.get_mac_address()
		
		my_mac = my_mac / 9527
		my_mac = int(my_mac)
		my_mac = my_mac * 2828
		my_mac = my_mac - 123456789
		my_mac = my_mac + 987654321

		enter1 = int(enter)
		print("mac = %d"%my_mac)
		print("enter = %d "%enter1)
		
		if my_mac == enter:
			
			print(" 校验正确")
			#self.excel_sheet.write(8, 0, enter1)
			#self.ExcelFile_auto_click.save('auto_click.xls')
			return True
			
		else:
		
			print(" 校验错误")
			return False
			
	# 获取 本地秘钥
	def keys_confirm(self):
		self.ExcelFile_auto_click  = xlrd.open_workbook(r'E:\xiaotuzi\auto_click.xls')
		#tem1  = xlrd.open_workbook(r'E:\PYthon\test\test.xls')

		aotu_click = self.ExcelFile_auto_click.sheet_names()
		self.key_confirm_sheet = self.ExcelFile_auto_click.sheet_by_name(aotu_click[0])
		key_data = self.key_confirm_sheet.cell_value(8, 0)

		
		key_data = int(key_data)
		# key_data = "12321"
		
		return key_data
	
	def aotu_verify(self):
		temp = self.keys_confirm()
		
		return self.Encryption_algorithms(temp)
		













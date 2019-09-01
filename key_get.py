from tkinter import *
from tkinter import *
from tkinter.messagebox import *
import datetime

class LoginPage(object): 
	def __init__(self, master=None): 
		self.root = master #定义内部变量root 
		self.root.geometry('%dx%d' % (300, 180)) #设置窗口大小 
		self.username = StringVar() 
		self.password = StringVar() 
		self.createPage() 

	def createPage(self): 
		self.page = Frame(self.root) #创建Frame 
		self.page.pack() 
		Label(self.page).grid(row=0, stick=W) 
		Label(self.page, text = '机械编码: ').grid(row=1, stick=W, pady=10) 
		Entry(self.page, textvariable=self.username).grid(row=1, column=1, stick=E) 
		Label(self.page, text = '秘钥生成: ').grid(row=2, stick=W, pady=10) 
		Entry(self.page, textvariable=self.password).grid(row=2, column=1, stick=E) 
		Button(self.page, text='确定', command=self.loginCheck).grid(row=3, stick=W, pady=10) 
		Button(self.page, text='退出', command=self.page.quit).grid(row=3, column=1, stick=E) 

	def loginCheck(self): 
		name = self.username.get() 
        
		nowTime = datetime.datetime.now().strftime('%Y%m%d')
		nowTime = int(nowTime)
        
		my_mac = int(name)
		my_mac = my_mac / 9527
		my_mac = int(my_mac)
		my_mac = my_mac * nowTime
		my_mac = my_mac - 123456789
		my_mac = my_mac + 987654321
		self.password.set(my_mac)

root = Tk() 
root.title('小兔子乖乖秘钥生产器') 
LoginPage(root) 
root.mainloop()


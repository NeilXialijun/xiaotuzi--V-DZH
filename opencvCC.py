#coding=utf-8
# import win32api
# import win32con
# import win32gui
# from ctypes import *
import time
import requests
# from tkinter import *
import cv2
from PIL import Image
import pytesseract
# import io
# import sys
import numpy as np
import matplotlib.pyplot as plt
import time
from PIL import ImageGrab
# import mousePrint
# from mousePrint import *

# 清楚黑杆的
# test_list = []  #= [50, 100, 150, 200, 250, 300, 350, 400, 450, 500]

# sys.stdout = io.TextIOWrapper(sys.stdout.buffer,encoding='gb18030')
# pytesseract.pytesseract.tesseract_cmd = 'C:/Program Files (x86)/Tesseract-OCR/tesseract.exe'

# 每抓取一次屏幕需要的时间约为1s,如果图像尺寸小一些效率就会高一些
def screen_capture(x, y, x1, y1):
    img = None
    beg = time.time()
    debug = False
    for i in range(10):
        img = ImageGrab.grab(bbox=(x, y, x1, y1))
        img = np.array(img.getdata(), np.uint8).reshape(img.size[1], img.size[0], 3)
    end = time.time()
    # print(end - beg)
    # print(":%d %d %d %d \n"%(x, y, x1, y1))
    cv2.imwrite(r'E:\xiaotuzi\5.jpg', img)
    return img

def processor_discern(img, test_list):
    # 1、读取图像，并把图像转换为灰度图像并显示
    # 转换了灰度化
    img_gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    # 将图片做二值化处理
    ret, im_inv = cv2.threshold(img_gray,220,255,cv2.THRESH_BINARY_INV)

    # 应用高斯模糊 降噪
    kernel = 1/16*np.array([[1,2,1], [2,4,2], [1,2,1]])
    im_blur = cv2.filter2D(im_inv,-1,kernel)

    # 可以看到一些颗粒化的噪声被平滑掉了
    # 降噪后，我们对图片再做一轮二值化处理
    ret, im_res = cv2.threshold(im_blur,230,255,cv2.THRESH_BINARY)

    # 保存黑白图片
    cv2.imwrite(r'E:\xiaotuzi\55.jpg', im_res)

    img_thre = im_res
    white = (255, 255, 255)
    n = 0
    # 填充黑 条
    while n < 9:
        cv2.line(img_thre, (test_list[n], 0), (test_list[n], 50), white, 3)
        n += 1

    #cv2.imshow("image", img_thre)
    #cv2.waitKey()

    cv2.imwrite(r'E:\xiaotuzi\66.jpg', img_thre)
    return img_thre
    # 识别文字
    # text = pytesseract.image_to_string(Image.open('66.jpg'))
    # print(text)

if __name__ == "__main__":
    while True:
        get_mouse_point()

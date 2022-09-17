# -*- coding: utf-8 -*-
# python3.7.5 
# 安装cv2    pip install opencv-contrib-python安装cv2  
# cd /Users/wuchunlong/local/upgit/cv2/cv2-qr-xlsx
# 使用电脑自带摄像头，扫描二维码，数据添加到电子表格。苹果电脑测试成功！
# 1、如果电子表格不存在，则创建一个空的电子表格；
# 2、如果有声音提示，说明扫描数据写入电子表格成功；
# 3、支持扫描的二维码，中文、英文、数字
# 4、在视频图像上按Q键退出

import os
import cv2
import pyzbar.pyzbar as pyzbar
from playsound import playsound
from openpyxl import Workbook
from xlutils.copy import copy
from xlrd import open_workbook 

XLSX_PATH = 'qr.xlsx'   #保存电子表格文件名
def write_excel(excel_name, data):
    """    
    数据data写入excel文件   2020.03.28
    excel_name = 'test.xlsx' # 文件名   
    data = [['姓名','职务','年龄'],['刘备','董事长','50'],['诸葛亮','军事','46'],['赵云','将军','39']]     
    write_excel(excel_name, data)
    """
    ret = True
    try:
        sheet_name = "sheet1"
        wb = Workbook()
        ws = wb.active
        ws.title = sheet_name  # sheet名称
        for row in data:
            ws.append(row)
        wb.save(excel_name)              
    except Exception as ex:
        print('excel err: %s' %ex) 
        ret =  False   
    return ret

def add_data_excel(excel_name, one_data):
    """
    往excel最后追加一行内容。如果电子表格不存在，则创建一个空的电子表格 2022.09.16
    excel_name = 'test.xlsx' # 文件名 
    one_data = ['李四','上尉','19']
    """
    ret = True
    try:
        if not os.path.isfile(excel_name):
            write_excel(excel_name, [])  #如果电子表格不存在，则创建一个空的电子表格      
        r_xls = open_workbook(excel_name)  # 读取excel文件
        row = r_xls.sheets()[0].nrows  # 获取已有的行数
        excel = copy(r_xls)  # 将xlrd的对象转化为xlwt的对象
        worksheet = excel.get_sheet(0)  # 获取要操作的sheet    
        for (index,value) in enumerate(one_data):
            # 对excel表追加一行内容
            worksheet.write(row, index, value)  # 括号内分别为行数、列数、内容
        excel.save(excel_name)  # 保存并覆盖文件
    except Exception as ex:
        print('excel err: %s' %ex) 
        ret =  False   
    return ret

def detect():
    camera = cv2.VideoCapture(0)
    #print('开始的扫描...')
    while True:
        data = ''
        ret, frame = camera.read()
 
        # 转为灰度图像
        frame = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        barcodes = pyzbar.decode(frame)  # 解析摄像头捕获到的所有二维码
        
        # 遍历所有的二维码
        for barcode in barcodes:
            data = barcode.data.decode('utf-8')  # 对数据进行转码
        
        if data:
            data = data.split(',')
            data = [i.strip() for i in data]
            #print('data:', data)            
            add_data_excel(XLSX_PATH, data) #电子文件名XLSX_PATH
            playsound('5012.wav') 
               
        frame = cv2.rectangle(frame,  # 图片
                     (300, 200),  # (xmin, ymin)左上角坐标
                     (600, 500),  # (xmax, ymax)右下角坐标
                     (0, 0, 0), 1)  # 颜色，线条宽度

        # 在图像上显示, 注意：在视频图像上按Q键退出
        cv2.putText(frame,
                "Press Q To Quit. Audio prompt for successful QR code scanning",
                (10, 30), cv2.FONT_HERSHEY_SIMPLEX, 0.8,
                (0, 0, 0), 2)        
        cv2.imshow('QR-CODE', frame)
        if(cv2.waitKey(1) & 0xFF) == ord('q'):
            break
               
    camera.release()
    cv2.destroyAllWindows()
 
 
if __name__ == '__main__':

    detect()
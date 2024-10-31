import ElectricityInquiry
import openpyxl as op
import schedule
import time
import subprocess
import apscheduler
from apscheduler.schedulers.blocking import BlockingScheduler
from datetime import datetime

if __name__ == '__main__':
    #创建工作簿,子表
    workbook = op.Workbook()
    worksheet = workbook.active
    #调整表头
    worksheet.column_dimensions['A'].width = 20
    worksheet.column_dimensions['B'].width = 20
    worksheet.append(['BuildingId','Building'])##表头
    workbook.save("Buildings.xlsx")
    #数据少不做格式化
    acc=str(input("校园卡卡号(6位数字):\n"))
    bui=str(input("宿舍楼(字母(大写)+数字):\n"))
    roo=str(input("房间号(大小写均可):\n"))
    ElectricityInquiry.empty_query(acc)
    print("该房间剩余电量："+ ElectricityInquiry.query(acc, bui, roo)+"元.")
    print("需要定时查询和邮件提醒请输入1:")
    x = (int)(input())##是否开启自动查询
    if x == 1:
        H = (int)(input("输入你期望的查询间隔(整数,小时):\n"))
        add = (str)(input("输入你的邮箱:\n"))
        ElectricityInquiry.auto_check(H,add,acc,bui,roo)
    else :
        print("已停止.")


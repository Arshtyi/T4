import json
import requests
import openpyxl as op
from openpyxl import load_workbook
import threading
import smtplib
from smtplib import SMTP_SSL
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header
import time

#楼信息
BUILDINGS = eval("""
[
    {
         "buildingid":"1574231830",
         "building":"T1"
    },
    {
         "buildingid":"1574231833",
         "building":"T2"
    },
    {
         "buildingid":"1574231835",
         "building":"T3"
    },
    {
         "buildingid":"1503975832",
         "building":"S1"
    },
    {
         "buildingid":"1503975890",
         "building":"S2"
    },
    {
         "buildingid":"1503975967",
         "building":"S5"
    },
    {
         "buildingid":"1503975980",
         "building":"S6"
    },
    {
         "buildingid":"1503975988",
         "building":"S7"
    },
    {
         "buildingid":"1503975995",
         "building":"S8"
    },
    {
         "buildingid":"1503976004",
         "building":"S9"
    },
    {
         "buildingid":"1503976037",
         "building":"S10"
    },
    {
         "buildingid":"1599193777",
         "building":"S11-13"
    },
    {
         "buildingid":"1661835249",
         "building":"B1"
    },
    {
         "buildingid":"1661835256",
         "building":"B2"
    },
     {
         "buildingid":"1661835273",
         "building":"B5"
    },
    {
         "buildingid":"1693031698",
         "building":"B9"
    },
    {
         "buildingid":"1693031710",
         "building":"B10"
    }
]
""")
###将输入的楼名字转为对应id
def building_to_id(building):
    for _building in BUILDINGS:
        if _building['building'] == building:
            return _building['buildingid']#如果输入的楼名字在这里,输出id
    print('ERROR: Wrong building number')
    input("任意键退出")
    exit(-1)

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Linux; Android 10; SM-G9600 Build/QP1A.190711.020; wv) AppleWebKit/537.36 (KHTML, like Gecko) Version/4.0 Chrome/86.0.4240.198 Mobile Safari/537.36",
    "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
}

def query(account, building, room):
    """
    :正式使用的data,操作
    :param account: 6位校园卡账号
    :param building: 宿舍楼名称, ['T1', 'T2', 'T3', 'S1', 'S2', 'S5', 'S6', 'S7', 'S8', 'S9', 'S10', 'S11-13', 'B1', 'B2', 'B5', 'B9', 'B10']
    :param room: 宿舍号
    :return: 电余量
    """
    data = {
        "jsondata": json.dumps(
            {
            #转成Python对象
            "query_elec_roominfo": {
                "aid": "0030000000002505",
                "account": str(account),#校园卡
                 "room": {
                    "roomid": room,
                    "room": room
                 },#房间
                "floor": {
                    "floorid": "",
                    "floor": ""
                },#层,空
                "area": {
                    "area": "青岛校区",
                    "areaname": "青岛校区"
                },#区
                "building": {
                    "buildingid": building_to_id(building),
                    "building": building
                }#楼
            }
            ######不用ascii编码
        }, ensure_ascii=False),
        "funname": "synjones.onecard.query.elec.roominfo",
        "json": "true"
    }
    try:
        response = requests.post('http://10.100.1.24:8988/web/Common/Tsm.html', headers=HEADERS, data=data, timeout=3)
        #print(response.text)############
        ########3返回的信息
        electricity = json.loads(response.text)['query_elec_roominfo']['errmsg']
        return electricity[8:]
    except Exception as e:
        print(e)
        exit(-1)
        
def empty_query(account):
    """
    :
    :用于爬取buildingid的empty_data
    :
    """
    empty_data = {
        "jsondata": json.dumps(
            {
            #转成Python对象
            "query_elec_building": {
                "aid": "0030000000002505",
                "account": str(account),#校园卡
                "area": {
                    "area": "青岛校区",
                    "areaname": "青岛校区"
                },#区
                "building": {
                    "buildingid": "",
                    "building": ""
                }#楼
            }
            ######不用ascii编码
        }, ensure_ascii=False),
        "funname": "synjones.onecard.query.elec.building",
        "json": "true"
    }
    
    response = requests.post('http://10.100.1.24:8988/web/Common/Tsm.html', headers=HEADERS, data=empty_data, timeout=3)
    buildingtab = json.loads(response.text)['query_elec_building']['buildingtab']#获得list-buildingtab
    workbook = load_workbook(filename = r'./Buildings.xlsx')
    worksheet = workbook.active#取表
    #print(buildingtab)
    for a_building in buildingtab:
        a_building_id = a_building["buildingid"]
        a_building_name = a_building["building"]
        content = [a_building_id , a_building_name]
        worksheet.append(content)######写入信息
    workbook.save("Buildings.xlsx")###存入


def email_query(account,building,room,add):
    host_server = 'smtp.qq.com'####qq邮箱smtp服务器
    sender_qq = "3842004484@qq.com"##sender_qq
    password = "nrwfxmrzirtxcbhb"#授权码
    receiver_qq = [add]###收件人
    mail_title = "寝室电量自动查询结果"##头
    ##########内容
    email_content = "同学,你上次查询的"+str(building)+"宿舍楼"+str(room)+"房间目前剩余电量:"+ query(account,building,room) + "元."
    if (float)(query(account,building,room)) < 10:
        email_content = email_content + "\n寝室剩余电量较少,注意及时充值."
        ####电量低时提醒
    mail_content = email_content
    #########编辑邮件
    msg = MIMEMultipart()
    msg["Subject"] = Header(mail_title,'utf-8')
    msg["From"] = sender_qq
    msg["To"] = ";".join(receiver_qq)
    msg.attach(MIMEText(mail_content,'plain','utf-8'))
    #接到服务器
    smtp = SMTP_SSL(host_server)

    smtp.login(sender_qq,password)
    #发送
    smtp.sendmail(sender_qq,receiver_qq,msg.as_string())
    #结束
    smtp.quit()

    print("邮件发送成功.")


    

def auto_check(H,add,account,building,room):###自动查询功能
    while True:
        try :
            email_query(account,building,room,add)
            time.sleep(H*60*60)
        except KeyboardInterrupt:#停止信号
            print("程序已终止.")#####Ctrl+C停止
            break





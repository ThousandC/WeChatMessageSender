from wxpy import *
import xlrd
from datetime import datetime
from apscheduler.schedulers.background import BackgroundScheduler


#扫码登陆微信网页版
bot=Bot()

#加载待发送信息
def info():
    Name=[]
    Message=[]
    Time=[]
    book = xlrd.open_workbook(input("输入信息表格路径："))
    Sheet = book.sheets()[0]
    for i in range(1,Sheet.nrows):
        Name.append(Sheet.row_values(i)[0])
        Message.append(Sheet.row_values(i)[1])
        Time.append(Sheet.row_values(i)[2])
    return [Name,Message,Time]
    
#发送信息
def Send(name,message):
    my_friend = bot.friends().search(name)
    if len(my_friend) == 1:
        my_friend[0].send(message)
        print("给{}的内容已发送".format(my_friend))
    else:
        if len(my_friend) == 0:
            print("找不到",name)
        else:
            print("存在同名好友")
            print(my_friend)

info = info()
name = info[0]
message = info[1]
time = info[2]
sc = BackgroundScheduler()
for i in range(len(name)):
    if time[i] == '':
        Send(name[i],message[i])
    else:
        sc.add_job(Send,'date',run_date=time[i],args=[name[i],message[i]])
    print(name,"信息发送成功")
sc.start()

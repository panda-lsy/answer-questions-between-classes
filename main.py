# -*- coding: utf-8 -*- 
from ast import Pass
import os
import shutil
import easygui as g
import time
import webbrowser                                       #自动打开源码网站
import win32file
import xlwt
import random

nowhour=time.strftime("%H", time.localtime()) 
datetime=time.strftime("%b %d", time.localtime()) 
weekday=time.strftime("%a",time.localtime())            #检测今天是星期几
weekdayDisplay=time.strftime("%A",time.localtime()) 

moveDir=r'older_versions/'
sourceDir=r'older_versions/master/'
sourceDirtwo=r'older_versions/'
targetDir=r'/'
listDir=os.getcwd()
listDirtwo=os.path.join(listDir,"older_versions")
imageDir=r'older_versions/images/'
writetime=0                                             #填写次数

Names=[]
Goreasons=[]
GuessBackTime=[]
Gotime=[]
Psws=[]

listDir=listDir+'/'

mylink="https://github.com/panda-lsy/answer-questions-between-classes"

datanamesfinally=[]

def is_used(file_name):                                 #检测文件占用
    try:
        v_handle = win32file.CreateFile(file_name, win32file.GENERIC_READ, 0, None, win32file.OPEN_EXISTING, 
                                        win32file.FILE_ATTRIBUTE_NORMAL, None)
        result = bool(int(v_handle) == win32file.INVALID_HANDLE_VALUE)
        win32file.CloseHandle(v_handle)
    except Exception:
        return True
    return result

'''def backup():                                           #回档
    global sourceDirtwo
    global datanamesfinally
    Xlscount = 0                                       #检测文件数
    success_state = ""
    text="作者:LSY\n未经作者授权随意转载\n开源是一种美德。"
    imagename = "successful"
    def successGUI():                                   #回档成功或失败界面
        choice2=g.indexbox(text,image=imageDir+imagename+".png",title="HoMo答疑人口管理系统1.0:回档"+success_state,choices=("好的","查看使用说明"))
        if choice2 == 0:
            os._exit(0)
        if choice2 == 1:
            choice3 = g.indexbox(msg="使用说明：双击本应用即可立刻复制表格母本并设置好今日日期\n复制完毕后你可以再次开启应用回档之前的表格清单。\n待更新内容:\n最终目标:通过Python的QQAPI来进行各个学科群表格关键字自动更新表格。",title="HoMo答疑人口管理系统1.0:使用说明",choices=("好的","打开本项目的Github网址(你可能需要科学上网)"))
            if choice3 == 0:
                os._exit(0)
            if choice3 == 1:
                webbrowser.open(mylink, new=0, autoraise=True)
                os._exit(0)
    datanamestwo = os.listdir(listDirtwo)
    for datanametwo in datanamestwo:
        if os.path.splitext(datanametwo)[1] == '.xls':                         #目录下包含.xls的文件
            datanamesfinally.append(datanametwo)
            Xlscount = Xlscount + 1 
    if Xlscount >=2:  
        back=g.choicebox("请选择需要回档的文件", "文件回档", datanamesfinally)
        if back == None:
            os._exit(0)
        sourceDirtwo = os.path.join(sourceDirtwo,back)
        if is_used(sourceDirtwo) == True:
            success_state = "失败"
            text = "移动回档库文件失败!文件"+back+"被占用,请尝试关闭Word里面你要回档的文件名."
            imagename = "error"
            successGUI()
        else:
            shutil.move(sourceDirtwo,listDir)
            success_state = "成功"
            successGUI()   
    else:
        success_state = "失败"
        text = "移动回档文件失败!你需要检测older_versions这个库文件夹里是否拥有两个以上旧文件."
        imagename = "error"
        successGUI()'''
        
'''def day_check(weekdayDisplay ,nowhour):                 #检测今天日期，防误删

    if not nowhour=="16":
            choice1=g.indexbox("今天是"+weekdayDisplay+",现在还不到使用时间(16:00-17:00)",title="HoMo答疑人口管理系统1.0:防误触",choices=("好的,退出","不好,退出"))
            if choice1==0:
                os._exit(0)

            if choice1==1:
                os._exit(0)
            
            if choice1==None:
                os._exit(0)

Checkdate = day_check(weekdayDisplay, nowhour)'''

def writemessage(datetime,writetime,Names,Goreasons,GuessBackTime,Gotime,wb,sh1,Psws):
    msg = "请填写一下信息(其中带*号的项为必填项)"
    title = "HoMo答疑人口管理系统1.0:答疑信息填写"
    fieldNames = ["*姓名","*出去原因","预计何时回来"]
    fieldValues = []
    fieldValues = g.multenterbox(msg,title,fieldNames)
    #print(fieldValues)
    while True:
        if fieldValues == None :
            break
        errmsg = ""
        for i in range(len(fieldNames)):
            option = fieldNames[i].strip()
            if fieldValues[i].strip() == "" and option[0] == "*":
                errmsg += ("【%s】为必填项   " %fieldNames[i])
        if errmsg == "":
            #g.textbox(msg='请填写你的必填项', title='HoMo答疑人口管理系统1.0:填写错误', text='', codebox=0) 
            break
        fieldValues = g.multenterbox(errmsg,title,fieldNames,fieldValues)
    a='还有这些小伙伴也参加了答疑(｡･∀･)ﾉﾞ\n'
    for goname in Names:
        num=Names.index(goname)
        a += str(goname) + '\t' + str(Gotime[num]) + '\n'
    psw=random.randrange(0,1000)
    pswmi=False
    while True: #查重
        for password in Psws:
            if psw == password: #如果密码重复
                psw=random.randrange(0,1000)
                pswmi=True
        if pswmi == False:
            break                   
    g.textbox(msg="您填写的资料如下\n姓名:"+fieldValues[0]+"\n出去原因:"+fieldValues[1]+"\n预计返回时间:"+fieldValues[2]+'\n你的验证密码是:'+str(psw)+'\n注:验证密码是答疑完成后返回验证用的', title='HoMo答疑人口管理系统1.0:录入成功',text=a , codebox=0)
    writetime=writetime+1
    Psws.append(psw)
    Names.append(fieldValues[0]) 
    sh1.write(writetime, 0, fieldValues[0])
    Goreasons.append(fieldValues[1])
    sh1.write(writetime, 1, fieldValues[1])
    Gotime.append(time.strftime("%H:%M:%S", time.localtime()))
    sh1.write(writetime, 2, time.strftime("%H:%M:%S", time.localtime()) )
    GuessBackTime.append(fieldValues[2])
    sh1.write(writetime, 3, fieldValues[2])
    wb.save(str(datetime)+".xls")

def verify(Names,sh1,Psws): 
    msg = "请输入姓名和密码"
    title = "HoMo答疑人口管理系统1.0:用户登录接口"
    user_info = []
    user_info = g.multpasswordbox(msg,title,("姓名","密码"))
    for name in Names:
        if user_info[0] == name:
            num=Names.index(name)
            if user_info[1] == Psws[num]:
                sh1.write(num, 4, 'Yes')
                break
            else:
                imagename='error'
                text='密码错误'
                g.indexbox(text,image=imageDir+imagename+".png",title="HoMo答疑人口管理系统1.0:密码错误",choices=("返回"))
                break
    imagename='error'
    text='不存在用户名'
    g.indexbox(text,image=imageDir+imagename+".png",title="HoMo答疑人口管理系统1.0:不存在用户名",choices=("返回"))
        

def main():
    '''Checkdate'''
    goname=''
    wb = xlwt.Workbook()
    sh1 = wb.add_sheet('外出记录')
    sh1.write(0, 0, '姓名')
    sh1.write(0, 1, '出去原因')
    sh1.write(0, 2, '出去时间')
    sh1.write(0, 3, '预计返回时间')
    sh1.write(0, 4, '是否返回')
    while True:
        text='HoMo答疑人口管理系统1.0:\n如果你需要出去答疑,请点击[申请答疑]按钮申请答疑。\n如果答疑完成,请点击[答疑完成]按钮结束外出答疑。'
        imagename='logo'
        choice=g.indexbox(text,image=imageDir+imagename+".png",title="HoMo答疑人口管理系统1.0:主界面",choices=("申请答疑","答疑完成",'控制台'))
        if choice == 0:
            writemessage(datetime,writetime,Names,Goreasons,GuessBackTime,Gotime,wb,sh1,Psws)
        if choice == 1:
            verify(Names,sh1,Psws)
    os._exit(0)

if __name__ == '__main__':
    main()


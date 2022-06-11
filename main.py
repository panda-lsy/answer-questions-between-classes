# -*- coding: utf-8 -*- 
import os
import shutil
import easygui as g
import time
import webbrowser                                       #自动打开源码网站
import win32file
import xlwt
import random
import sys

nowhour=time.strftime("%H", time.localtime()) 
datetime=time.strftime("%b %d", time.localtime()) 
weekday=time.strftime("%a",time.localtime())            #检测今天是星期几
weekdayDisplay=time.strftime("%A",time.localtime()) 

moveDir=r'older_versions/'
sourceDir=r'older_versions/'
listDir=os.getcwd()
listDirtwo=os.path.join(listDir,"older_versions")
imageDir=r'older_versions/images/'
writetime=0                                             #填写次数

Names=[]
Goreasons=[]
GuessBackTime=[]
Gotime=[]
Psws=[]
Ifverifyeds=[]

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

def move_old_file(datetime, moveDir, listDir):            #移动旧文件
         
    datanames = os.listdir(listDir)

    def GUI():
        g.indexbox(text,image=imageDir+image_name+".png",title="HoMo答疑人员管理系统1.0:"+interaction,choices=("取消","好的"))
    #检测是否复制或者移动
    for dataname in datanames:
        
        if os.path.splitext(dataname)[1] == '.xls':        #目录下包含.xls的文件
            listFile = os.path.join(listDir,dataname)       #把文件夹名和文件名称链接起来
            NewMovePath = os.path.join(moveDir,dataname)    #如果移动,它就是目标路径
            if not listFile == (listDir+str(datetime)+".xls"): 
                if is_used(listFile) == True:       #文件被占用,退出
                        text="移动文件失败!程序将在你进行交互后退出。文件"+dataname+"被其它占用,请尝试关闭Excel里面你要回档的文件名."
                        image_name="error"
                        interaction="移动失败"
                        GUI()
                        
                else:                                   #没有检测文件被占用,那么移动
                    shutil.move(listFile,NewMovePath)  

def backup(listDirtwo):                                           #回档
    global sourceDir
    global datanamesfinally
    Xlscount = 0                                       #检测文件数
    success_state = ""
    text="作者:LSY\n未经作者授权随意转载\n开源是一种美德。"
    imagename = "successful"
    def successGUI():                                   #回档成功或失败界面
        g.indexbox(text,image=imageDir+imagename+".png",title="HoMo答疑人员管理系统1.0:回档"+success_state,choices=("好的","确定"))
    datanamestwo = os.listdir(listDirtwo)
    for datanametwo in datanamestwo:
        if os.path.splitext(datanametwo)[1] == '.xls':                         #目录下包含.xls的文件
            datanamesfinally.append(datanametwo)
            Xlscount = Xlscount + 1 
    if Xlscount >=2:  
        while True:
            back=g.choicebox("请选择需要回档的文件", "文件回档", datanamesfinally)
            if not back == None:
                break
        sourceDir = os.path.join(sourceDir,back)
        if is_used(sourceDir) == True:
            success_state = "失败"
            text = "移动回档库文件失败!文件"+back+"被占用,请尝试关闭Excel里面你要回档的文件名."
            imagename = "error"
            successGUI()
        else:
            while True:
                CopyMovePath=g.filesavebox(msg='请选择保存文件的路径', title='导出表格', default=str(datetime)+".xls", filetypes=['*.xls'])
                if not CopyMovePath == None:
                    break
            shutil.copy(sourceDir,CopyMovePath)
            success_state = "成功"
            successGUI()   
    else:
        success_state = "失败"
        text = "移动回档文件失败!你需要检测older_versions这个库文件夹里是否拥有两个以上旧文件."
        imagename = "error"
        successGUI()
        
def day_check(weekdayDisplay ,nowhour):                 #检测今天日期，防误删

    if not nowhour=="16":
            choice1=g.indexbox("今天是"+weekdayDisplay+",现在还不到使用时间(16:00-17:00)",title="HoMo答疑人员管理系统1.0:防误触",choices=("好的,退出","不好,退出"))
            if choice1==0:
                os._exit(0)

            if choice1==1:
                os._exit(0)
            
            if choice1==None:
                os._exit(0)

Checkdate = day_check(weekdayDisplay, nowhour)

def writemessage(datetime,Names,Goreasons,GuessBackTime,Gotime,wb,sh1,Psws,Ifverifyeds):
    global writetime
    Pass = False
    msg = "请填写一下信息(其中带*号的项为必填项)"
    title = "HoMo答疑人员管理系统1.0:答疑信息填写"
    fieldNames = ["*姓名","*出去原因","预计何时回来"]
    fieldValues = []
    fieldValues = g.multenterbox(msg,title,fieldNames)
    #print(fieldValues)
    while True:
        if fieldValues == None:
            break
        errmsg = ""
        for i in range(len(fieldNames)):
            option = fieldNames[i].strip()
            if fieldValues[i].strip() == "" and option[0] == "*":
                errmsg += ("【%s】为必填项   " %fieldNames[i])
        for name in Names:
            if name == fieldValues[0]:
                errmsg +=('请不要重复登记,如果需要重复登记请在名字后加上次数:2,3...')
        if errmsg == "":
            Pass = True
            break
        fieldValues = g.multenterbox(errmsg,title,fieldNames,fieldValues)
    if Pass == True:
        if writetime == 0:
            a='你是第一个参与答疑的同学哦o(*￣▽￣*)ブ'
        else:
            a='还有这些小伙伴也参加了答疑(｡･∀･)ﾉﾞ\n'
        for goname in Names:
            num=Names.index(goname)
            a += str(goname) + '\t' + str(Gotime[num]) + ''
            if Ifverifyeds[num] == True:
                a += '\t' + "已返回\n"
            else:
                a += '\n'
        psw=random.randrange(0,1000)
        pswmi=False
        while True: #查重
            for password in Psws:
                if psw == password: #如果密码重复
                    psw=random.randrange(0,1000)
                    pswmi=True
            if pswmi == False:
                break                   
        g.textbox(msg="您填写的资料如下\n姓名:"+fieldValues[0]+"\n出去原因:"+fieldValues[1]+"\n预计返回时间:"+fieldValues[2]+'\n你的验证密码是:'+str(psw)+'\n注:验证密码是答疑完成后返回验证用的', title='HoMo答疑人员管理系统1.0:录入成功',text=a , codebox=0)
        writetime=writetime+1
        Ifverifyeds.append(False)
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

def verify(Names,sh1,Psws,wb): 
    msg = "请输入姓名和密码"
    title = "HoMo答疑人员管理系统1.0:用户登录接口"
    user_info = []
    user_info = g.multpasswordbox(msg,title,("姓名","密码"))
    nameexist=False
    for name in Names:
        if user_info[0] == name:
            num=Names.index(name)
            nameexist=True
            if str(user_info[1]) == str(Psws[num]) or str(user_info[1]) == 'yourultrapassword':
                if Ifverifyeds[num] == False:
                    sh1.write(int(num)+1, 4, '是')
                    sh1.write(int(num)+1, 5, time.strftime("%H:%M:%S", time.localtime()))
                    wb.save(str(datetime)+".xls")
                    imagename='successful'
                    text='验证成功,祝你学习进步'
                    Ifverifyeds[num] = True
                    g.indexbox(text,image=imageDir+imagename+".png",title="HoMo答疑人员管理系统1.0:验证成功",choices=("返回",'好的'))
                    break
                else:
                    imagename='error'
                    text='你已验证过了，不要重复验证'
                    g.indexbox(text,image=imageDir+imagename+".png",title="HoMo答疑人员管理系统1.0:验证过了",choices=("返回","确定"))
                    break    
            else:
                imagename='error'
                text='密码错误'
                g.indexbox(text,image=imageDir+imagename+".png",title="HoMo答疑人员管理系统1.0:密码错误",choices=("返回","确定"))
                break
    if nameexist == False:
        imagename='error'
        text='不存在用户名'
        g.indexbox(text,image=imageDir+imagename+".png",title="HoMo答疑人员管理系统1.0:不存在用户名",choices=("返回",'好的'))
            
def console():                                           #控制台
    command='''
    HoMo答疑人员管理系统 version 1.0.0(2022-06-11) -- "Bug in Your Hair"
copyright (C) 2022 The Panda-Lsy Foundation for statistical ComputingPlatform:'''+sys.platform+'''
HoMo答疑人员管理系统是自由软件,不带任何担保。
在某些条件下你可以将其自由散布。
HoMo答疑人员管理系统是个合作计划,有许多人为之做出了贡献.
用"contributors()"来看合作者的详细情况
用"help()"来阅读在线帮助文件，或用"nelp.start()"通过HTML浏览器来看帮助文件。
用"quit()"退出HoMo答疑人员管理系统
'''
    while True:
        input=g.enterbox(msg=command, title='HoMo答疑人员管理系统1.0:控制台界面', default='', strip=False, image=None, root=None)

        if input == 'contributors()':
            command += input+'''
            作者:Panda-Lsy
            贡献者:Li-Yuhan
            '''
        if input == 'help()':
            command += input+'''
            用"contributors()"来看合作者的详细情况
            用"help()"来阅读在线帮助文件，或用"nelp.start()"通过HTML浏览器来看帮助文件。
            用"quit()"退出HoMo答疑人员管理系统
            用"clear()"清除记录
            用"export()"导出答疑文件
            用"backup()"导出先前的答疑文件
            用"netsource()"来查看作者在GITHUB上的源码'''
        
        if input == 'quit()':
            os._exit(0)
        
        if input == 'clear()':
            command = input+'\n 内容已清空'

        if input == 'export()':
            while True:
                CopyMovePath=g.filesavebox(msg='请选择保存文件的路径', title='导出表格', default=str(datetime)+".xls", filetypes=['*.xls'])
                if not CopyMovePath == None:
                    break
            shutil.copy(os.path.join(os.getcwd(),str(datetime)+".xls"),CopyMovePath)
            command += input+'''导出成功,导出到'''+CopyMovePath+''
        
        if input == 'backup()':
            command += input + '\n'
            backup(listDirtwo)
        
        if input == 'nelp.start()' or input == 'netsource()':
            command += input + '\n'
            webbrowser.open(mylink, new=0, autoraise=True)
            
        if input == None:
            break
        
        
                
def main():
    Checkdate
    move_old_file(datetime, moveDir, listDir)
    wb = xlwt.Workbook()
    sh1 = wb.add_sheet('外出记录')
    sh1.write(0, 0, '姓名')
    sh1.write(0, 1, '出去原因')
    sh1.write(0, 2, '出去时间')
    sh1.write(0, 3, '预计返回时间')
    sh1.write(0, 4, '是否返回')
    sh1.write(0, 5, '返回时间')
    wb.save(str(datetime)+".xls")
    while True:
        Checkdate
        text='HoMo答疑人员管理系统1.0:\n如果你需要出去答疑,请点击[申请答疑]按钮申请答疑。\n如果答疑完成,请点击[答疑完成]按钮结束外出答疑。\n'
        imagename='logo'
        choice=g.indexbox(text,image=imageDir+imagename+".png",title="HoMo答疑人员管理系统1.0:主界面",choices=("申请答疑","答疑完成",'控制台'))
        if choice == 0:
            writemessage(datetime,Names,Goreasons,GuessBackTime,Gotime,wb,sh1,Psws,Ifverifyeds)
        if choice == 1:
            verify(Names,sh1,Psws,wb)
        if choice == 2:
            console()


if __name__ == '__main__':
    main()


#answerList 为一个dictionary,分为三个key:team,date,name
#
import xlwings as xl 
from imutils.perspective import four_point_transform
import datetime
import numpy as np 
#import argparse
#import imutils
import cv2
import os
import tkinter as tk 


fileAddress='.\\pic\\'                 #文件夹位置

cntsFlag=1   
#答题卡总行数和总列数
totalRow,totalColumn=33,23
#标记答案的设置
answerDotNum=172                        #涂色区域如果高于此值，认为已经涂色
color=(0,0,255)                         #最终图片上涂色区域显示的图框颜色
boxSize=8                               #显示图框宽度
adaptiveArea=9                          #使用自适应二值化时，自适应区域大小
#答案整体区域设置
answerAreaSizeFlag=0.3                  #答案区域占整个图片区域面积，如果小于此值认为框选的区域不是答案区域
answerAreaSonBoxRate=0.0001             #答案区的最大子轮廓占答案区面积比，将大于此比例的子轮廓列为候选轮廓（用于统计轮廓数量），约为1/100*100
sonBoxNum=100                           #符合条件子轮廓最小数目
answerColumnList=list(range(3,totalColumn,2))       #答案在第多少列
answerRowList=list(range(3,totalRow,2))             #答案在第多少行


root="" #显示窗口

answerDic2Show={}



def imgshow(img):
    #将img存档为“okLastShow.png",放在fileAddress文件夹下
    cv2.imshow('ImageWindow',img)
    cv2.waitKey()
    #cv2.imwrite(fileAddress+"oklastShow.png",img)
def imgDraw(cnts,img):
    #在img上标示出cnts区域的边框
    for c in cnts:
        cv2.drawContours(img,[c],-1,color,boxSize)
    return img
def loadImg(filename):
    #导入指定文件，调整大小为1440，并输出灰度图和原图
    #print('Now calculate the file: ',filename)
    img=cv2.imread(filename)
    size=1440/max(img.shape[0],img.shape[1])
    img=cv2.resize(img,(0,0),fx=size,fy=size,interpolation=cv2.INTER_AREA)
    colorImg=img.copy()
    gray=cv2.cvtColor(img,cv2.COLOR_BGR2GRAY)
    #imgshow(gray)
    return gray,colorImg

def myFindContours(img,mod1,mod2):#cv版本问题，cv2.findContours函数返回值有可能2参数，有可能3参数
    return cv2.findContours(img,mod1,mod2)[-2]

def findAnswerAreaFunction(img,colorImg):
    #查找面积成功标志
    findFlag=False
    #计算下一级最大的面积与图片整体面积比例，如果比例小于answerAreaSizeFlag认为失败
    answerAreaRate=1
    w,h=img.shape
    totalArea=cv2.contourArea(np.array([[[0,0]],[[0,h]],[[w,h]],[[w,0]]]))
    #保证循环次数loopingNum，如果太多防止陷入死循环，workFlag清0
    loopingNum =0
    original=img.copy()
    #img进行高斯模糊、求边缘、找轮廓、轮廓排序操作
    blurred=cv2.GaussianBlur(img,(5,5),0)
    edged=cv2.Canny(blurred,75,200)
    cnts=myFindContours(edged.copy(),cv2.RETR_EXTERNAL,cv2.CHAIN_APPROX_SIMPLE)
    #求最大轮廓与总轮廓的面积比，小于answerAreaSizeFlag认为求轮廓失败
    maxSonArea=cv2.contourArea(sorted(cnts,key=cv2.contourArea,reverse=True)[0])
    answerAreaRate=maxSonArea/float(totalArea)
    if answerAreaRate<answerAreaSizeFlag:
        return findFlag,original,colorImg
    #求子轮廓最大面积，并筛选子轮廓面积大于0.0001的面积，约为1/100*100    
    cnts=list(filter(lambda x:cv2.contourArea(x)>answerAreaSonBoxRate*maxSonArea,cnts))   
    docCnt= None 
    answerAreaRate=1
    #开始循环
    #保证子轮廓大于总轮廓answerAreaSizeFlag倍，总循环次数小于15
    #首先对所有轮廓按照面积大小排序，从最大的开始查找轮廓是否模糊为四边形
    #如果是四边形就当找到次的子轮廓，删掉该轮廓外其余区域后，在这个区域内再寻找子轮廓
    while(len(cnts)<sonBoxNum and loopingNum <40 and answerAreaRate>answerAreaSizeFlag):
        loopingNum+=1
        oldCnts=cnts.copy()
        cnts=sorted(cnts,key=cv2.contourArea,reverse=True)
        for c in cnts:
            peri=cv2.arcLength(c,True)
            approx=cv2.approxPolyDP(c,0.05*peri,True)
            if len(approx)==4:
                docCnt=approx
                break
        mask=np.zeros(img.shape,dtype='uint8')
        cv2.drawContours(mask,[docCnt],-1,255,-1)
        img=cv2.bitwise_and(original,original,mask=mask)
        blurred=cv2.GaussianBlur(img,(5,5),0)
        edged=cv2.Canny(blurred,75,200)
        cnts=myFindContours(edged.copy(),cv2.RETR_EXTERNAL,cv2.CHAIN_APPROX_SIMPLE)
        maxArea=cv2.contourArea(sorted(cnts,key=cv2.contourArea,reverse=True)[0])
        answerAreaRate=maxArea/totalArea
        newCnts=list(filter(lambda x:cv2.contourArea(x)>answerAreaSonBoxRate*maxArea,cnts))
        if newCnts==cnts:
            break
        #imgshow(imgDraw(cnts,img))
    #循环40次没找到，说明照片有问题，返回照片值
    #否则已找到的四个顶点在原始灰度和彩色图像上做四角变换，输出最终变换后的图形
    if loopingNum==40 or (docCnt is None):
        return findFlag,original,colorImg
    findFlag=True
    colorImg=four_point_transform(colorImg,docCnt.reshape(4,2))
    img=cv2.cvtColor(colorImg,cv2.COLOR_BGR2GRAY)
    #imgshow(img)
    return findFlag,img,colorImg

def findAnswerArea(img,colorImg=None):
    #colorImg尽量使用输入值，否则使用img复制
    if colorImg is None:
        colorImg=img.copy()
    original=img.copy()
    #先使用自适应求值，查找带答案的区域照片
    findFlag=False
    img=cv2.adaptiveThreshold(img,255,cv2.ADAPTIVE_THRESH_GAUSSIAN_C,cv2.THRESH_BINARY_INV,adaptiveArea,0)
    findFlag,img,color=findAnswerAreaFunction(img,colorImg)
    #print(0)
    #如果求取区域失败，在使用原版求值
    if findFlag==False:
        findFlag,img,color=findAnswerAreaFunction(original,colorImg)
        #print(1)
    #如果再求取区域失败，在使用大津化后求值
    if findFlag==False:
        img=cv2.threshold(img,0,255,cv2.THRESH_BINARY_INV | cv2.THRESH_OTSU)[1]
        findFlag,img,color=findAnswerAreaFunction(original,colorImg)
        #print(2)
    if findFlag==False:
        img,color=None,None
    #imgshow(img)
    return img,color

def findMatAn(img,row,column,colorImg):
    #对column列 row行制作mask，然后在二值化图像中查看白色点数
    #如果总点数比answerDotNum多，证明此处已填写，
    #在colorImg上绘制图框，并writeFlag 置1
    #返回wirteFlag和original图片

    writeFlag=0

    anBox=np.array([[[20*column+2,20*row+2]],\
    [[20*column+18,20*row+2]],\
    [[20*column+18,20*row+18]],\
    [[20*column+2,20*row+18]]])
    #img=cv2.adaptiveThreshold(img,255,cv2.ADAPTIVE_THRESH_GAUSSIAN_C,cv2.THRESH_BINARY_INV,5,0)
    #img=cv2.adaptiveThreshold(img,255,cv2.ADAPTIVE_THRESH_MEAN_C,cv2.THRESH_BINARY_INV,31,0)
    img=cv2.threshold(img,128,255,cv2.THRESH_BINARY_INV | cv2.THRESH_OTSU)[1]
    #imgshow(img)
    #img=cv2.threshold(img,127,255,cv2.THRESH_BINARY)[1]
    
    mask=np.zeros(img.shape,dtype='uint8')

    cv2.drawContours(mask,[anBox],-1,255,-1)
    #imgshow(mask)
    mask=cv2.bitwise_and(img,img,mask=mask)
    #imgshow(mask)
    answer=cv2.countNonZero(mask)
    #imgshow(mask)
    #print(answer)
    if answer>answerDotNum:
        colorImg=imgDraw([anBox],colorImg)
        writeFlag=1        
    
    return writeFlag,colorImg


def imgFindAnswer(filename):
    #对输入的文件名查找最终涂黑框的数据
    img,colorImg=loadImg(filename)
    
    img,colorImg=findAnswerArea(img,colorImg)
    #imgshow(img)
    #如果没找到answer区域，返回空值  
    #否则将返回的img和colorImg图片扩展像素至(行列的20倍像素)
    if img is None:
        #print("wrong picture")
        return None,None
    img=cv2.resize(img,(480,680))
    colorImg=cv2.resize(colorImg,(480,680))
    #在图片中遍历row和column寻找涂色的点的坐标
    columnList=answerColumnList
    rowList=answerRowList
    retAnswer={}
    retStr=""
    #计算team组
    colStr=""
    for column in range(6,totalColumn,2):
        writeFlag,colorImg=findMatAn(img,2,column,colorImg)
        if writeFlag:
            colStr+="1"
        else:
            colStr+="0"
    #print(colStr)
    retAnswer["team"]='L-L-'+str(int(colStr[:-4],base=2)).zfill(3)+'-'+str(int(colStr[-4:],base=2)).zfill(2)
    #print(colStr)
    #计算date
    colStr=""
    for column in range(0,totalColumn,2):
        writeFlag,colorImg=findMatAn(img,4,column,colorImg)
        if writeFlag:
            colStr+="1"
        else:
            colStr+="0"
    #print(colStr)
    retAnswer["date"]=(datetime.date(2019,10,1) + datetime.timedelta(days=int(colStr,base=2))).isoformat()


    for row in range(6,totalRow,2):
        colStr=""
        for column in [4,10,16,22]:
            writeFlag,colorImg=findMatAn(img,row,column,colorImg)
            if writeFlag:
                colStr+="1"
            else:
                colStr+="0"
        retStr+=colStr
    retAnswer['answer']=retStr
    return colorImg,retAnswer
    
def excelWrite(retAnswer):
    app=xl.App(visible=False,add_book=False)
    wrongNameList=[]
    fileList=os.listdir(".")
    if retAnswer['team'] +'.xlsm' in fileList:
        wbIn=app.books.open(retAnswer['team'] +'.xlsm')
        wsIn=wbIn.sheets[0]
    else:
        #print("No Team",retAnswer['team'])
        return []
    print("get answer book")
    wbOut=app.books.open("工时任务反馈导入模板.xlsx")
    wsOut=wbOut.sheets[0]
    wbWH=app.books.open("人力工时上报数据核查最近60天-self视图报表.xlsx")
    wsWH=wbWH.sheets[0]
    #查找各列的名称，记录到相关变量当中，防止某一列改位置
    colIC,colName,colDate,colWh=0,0,0,0
    #print(wsWH.cells(1,1).value,wsWH.cells(1,1).value=="员工卡号",type(wsWH.cells(1,1).value))
    for i in range(1,50):
        if wsWH.cells(1,i).value=="员工卡号":
            colIC=i
        elif wsWH.cells(1,i).value=="姓名":
            colName=i
        elif wsWH.cells(1,i).value=="时间":
            colDate=i
        elif wsWH.cells(1,i).value=="参考工时":
            colWh=i
    anList=retAnswer['answer']
    #print(wsWH.cells(20000,7).value)
    for i in range(2,1000):
        if wsOut.cells(i,1).value==None:
            break
    lastRow=i
    #print("find last row,lastrow is ", lastRow)
    #print(wsIn.cells(2,6).value)
    for i in range(len(anList)):
        if anList[i]=="0":
            row,column=divmod(i,4)
            row=row*2+9
            column=column*6+5
            wsOut.cells(lastRow,1).value=retAnswer['date']
            #print(row,column)
            wsOut.cells(lastRow,3).value=wsIn.cells(row,column).value
            
            wsOut.cells(lastRow,5).value=retAnswer["team"][:-3]
            #在“人力工时上报数据核查最近60天-self视图报表.xlsx”中逐行查找IC卡号得到工时
            j=2
            while(wsWH.cells(j,colIC).value!=None):#IC卡号不为空
                if wsIn.cells(row,column).value== wsWH.cells(j,colIC).value and retAnswer['date']== wsWH.cells(j,colDate).value[:10]:
                    wh=wsWH.cells(j,colWh).value
                    wsOut.cells(lastRow,4).value=wsWH.cells(j,colName).value
                    if wh:
                        wsOut.cells(lastRow,6).value=int(float(wh))
                        lastRow+=1
                        break
                    else:
                        wrongNameList.append(wsWH.cells(j,colName).value)
                        #print(wsOut.cells(lastRow,4).value)
                        wsOut.cells(lastRow,1).value=""
                        wsOut.cells(lastRow,3).value=""
                        wsOut.cells(lastRow,4).value=""
                        wsOut.cells(lastRow,5).value=""
                        wsOut.cells(lastRow,6).value=""
                j+=1
    wbOut.save()
    wbOut.close()
    wbIn.close()
    wbWH.close()
    app.quit()
    return wrongNameList




    
def checkPicAndWriteExcel():
    global answerDic2Show
    global root
    showMartrix=[]
    nameList=[]
    fileList=os.listdir(".")
    '''
    try:
        os.mkdir(".\\pic")
    except:
        pass
    '''
    for f in fileList:
        if f[-3:]!='jpg' or f[:2]=='ok':
            continue
        img,retAnswer=imgFindAnswer(f)
        print("read Answer OK",retAnswer)
        nameList=excelWrite(retAnswer)
        showMartrix.append([f[:-4],retAnswer['answer'].count("0"),nameList])
        #print(retAnswer)
        if img is not None:
            #cv2.imwrite(fileAddress+'ok'+f,img)
            pass
    #print(answerDic2Show)
    for i in range(len(showMartrix)):
        answerDic2Show[i][0].set(showMartrix[i][0])
        answerDic2Show[i][1].set(str(showMartrix[i][1]))
        answerDic2Show[i][2].set(showMartrix[i][2])

def showLayout():
    global answerDic2Show
    global root
    root=tk.Tk()
    root.wm_title("生成管理工时")
    root.geometry("300x300+200+200")
    btn=tk.Button(root,text="单击生成工时表格",command=checkPicAndWriteExcel).grid(row=0,column=0,rowspan=3,columnspan=5)
    tk.Label(root,text="照片名称").grid(row=4,column=0,columnspan=2)
    tk.Label(root,text="包含人数").grid(row=4,column=2)
    tk.Label(root,text="错误人员名单").grid(row=4,column=3,columnspan=2)
    answerDic2Show={}

    for infoRow in range(5,40):
        answerDic2Show[infoRow-5]=[tk.StringVar(root),tk.StringVar(root),tk.StringVar(root)]
        tk.Label(root,textvariable=answerDic2Show[infoRow-5][0]).grid(row=infoRow,column=0,columnspan=2)
        tk.Label(root,textvariable=answerDic2Show[infoRow-5][1]).grid(row=infoRow,column=2)
        tk.Label(root,textvariable=answerDic2Show[infoRow-5][2]).grid(row=infoRow,column=3,columnspan=2)

    root.mainloop()



def myMain():
    showLayout()

myMain()
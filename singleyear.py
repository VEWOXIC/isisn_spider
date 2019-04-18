#coding:utf-8
import json
import time
from io import BytesIO as Bytes2Data

import PIL
import pytesseract
import requests
import xlwt
from pylab import *


def listget(keyword):
    url='https://isisn.nsfc.gov.cn/egrantindex/cpt/ajaxload-complete?q='+keyword+'&arloncascade=none&limit=10&timestamp=1555572074679&locale=zh_CN&key=subject_code_index&cacheable=true&sqlParamVal='
    headers={'Origin':'https://isisn.nsfc.gov.cn','Referer':'https://isisn.nsfc.gov.cn/egrantindex/funcindex/prjsearch-list?locale=zh_CN','User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/73.0.3683.103 Safari/537.36'}
    data={'q':keyword,'key':'subject_code_index'}
    s=requests.Session()
    r = s.post(url, data = data,headers=headers)
    rawlist=eval(r.text)
    items={}
    for i in rawlist:
        items[i['id']]=i['title']
    return items
def getcheck():

    def request_download():
        import requests
        r = requests.get('https://isisn.nsfc.gov.cn/egrantindex/validatecode.jpg')   
        cookies = requests.utils.dict_from_cookiejar(r.cookies)
        return cookies,r.content
    def clear_dotnoise(img):
        w,h=img.shape
        for x in range (1,w-1):
            for y in range (1,h-1):
                count=img[x-1,y]+img[x+1,y]+img[x,y-1]+img[x,y+1]#255为白0为黑
                count=count/255#计算周围白色像素数量
                if img[x,y]==0:
                    if count>2:
                        img[x,y]=255#白色多于两个便认为这是个黑色噪点，抹为白色
                else:
                    if count<2:
                        img[x,y]=0#反之认为是白色噪点
        return img
    cookie,img_bytes=request_download()
    img=array(PIL.Image.open(Bytes2Data(img_bytes)).convert('L'))#将目标地址文件以灰阶模式打开，并且转换为ndarray便于二值化处理
    img=select([img>180],[np.uint8(255)],default=0)#筛选灰度大于180的点（可调）置为255认为是白色，其他点认为是黑色0
    img=clear_dotnoise(img)#调用去噪
    result=pytesseract.image_to_string(img,lang='eng',config='-psm 8 digits')#调用ocr采用config=digits（在tesseract根目录修改）psm8为将图片认为是单行文本
    result=result.replace(' ','')#去掉可能的空格
    result=result.replace('S','5')#识别时发现被识别为S的全都是5，强制变换
    return result,cookie

def i_am_the_spider(item_title,item_id,grant_code,which_year):
    check_code,thiscookie=getcheck()
    url='https://isisn.nsfc.gov.cn/egrantindex/funcindex/prjsearch-list?flag=grid&checkcode='+check_code
    headers={'Origin':'https://isisn.nsfc.gov.cn','Referer':'https://isisn.nsfc.gov.cn/egrantindex/funcindex/prjsearch-list','User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/73.0.3683.103 Safari/537.36'}
    result_date='resultDate^:prjNo:,ctitle:,psnName:,orgName:,subjectCode:'+item_title+',f_subjectCode_hideId:'+item_id+',subjectCode_hideName:'+item_title+',keyWords:,checkcode:'+check_code+',grantCode:'+grant_code+',subGrantCode:,helpGrantCode:,year:'+which_year+',sqdm:'+item_id
    FormData={'sidx:':'','_search':'false','rows':'200','page':'1','searchString':result_date+'[tear]sort_name1^:psnName[tear]sort_name2^:prjNo[tear]sort_order^:desc'}
    s=requests.Session()
    r = s.post(url, data = FormData,cookies=thiscookie,headers=headers)
    if 'html' in r.text:
        endcode=1
        print('Wrong check code...')
    else:
        endcode=0
        print('Data get...')
    return r.text,endcode

#yearlist=('2008','2009','2010','2011','2012','2013','2014','2015','2016','2017','2018')
datadict=listget(input('Input the parent index:'))
whichyear=input('Input a single year:')
#grantlist=('218','220','222','339','429','432','433','649','579','632','635','51','52','2699','70','7161')
grantlist=('218','220','222','339','429','432','433','649','579','632')

dataworkbook=xlwt.Workbook(encoding='ascii')
xlsname='D:/'+whichyear+'data.xls'
for grantcode in grantlist:#遍历级别
    nowsheet=dataworkbook.add_sheet(grantcode)
    linecount=0
    for itemid,itemtitle in datadict.items():#遍历级别下的所有id
        endcode=1#保证可以进行第一次请求
        while endcode:#若结束码为1则为失败请求
            print('Sending request for',itemtitle,grantcode,whichyear)
            try:#捕获所有错误，可能导致其他问题无法抛出
                textdata,endcode=i_am_the_spider(itemtitle,itemid,grantcode,whichyear)#调用spider，其中自动调用识别
            except:
                print('Due to timeout, Get a cup of coffee and take a nap zzz...')
                time.sleep(5)
                endcode=1
        textdata=textdata.replace('\n','').replace('\t<cell>','')
        templist=textdata.split('<row id="">')#若没有项目可能报错 待解决 已解决，读取records数量 洗过后形如初步分段，将其分为一个个项目块
        itemnumber=templist.pop(0)#取第一个包括
        print('Acquired %d records...' % (len(templist)))
        if '<records>0</records>' in itemnumber:
            #print('Acquired 0 record...')
            continue#若被注明无记录，则不写入
        for i in templist:
            objlist=i.split('</cell>')#单个项目块切分后 ['81270057', 'H0101', 'KLF5调控上皮极性介导肺上皮管腔形成的机制', '万华靖', '四川大学', '70', '2013-01至2016-12']
            objlist.pop(-1)
            for wordindex in range (0,len(objlist)):
                nowsheet.write(linecount,wordindex,label=objlist[wordindex])#按照规则填入
            linecount=linecount+1
        dataworkbook.save(xlsname)

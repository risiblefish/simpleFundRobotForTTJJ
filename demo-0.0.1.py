import xlrd
import requests
from bs4 import BeautifulSoup
import time


fileLocation = r'D:\YZX\application\Git\pytest\基金净值跟踪20190329.xlsx'

def get_one_page(url):
    response = requests.get(url)
    response.encoding = 'utf-8'
    if response.status_code == 200:
        return response.text
    return None

def myParseInt(val):
    if(type(val) == float):
        val = int(val)
    val = str(val)
    firstLeftChineseBracketIndex = val.rfind('（')
    fristLeftEnglishBracketIndex = val.rfind('(')
    leftIndex = firstLeftChineseBracketIndex if firstLeftChineseBracketIndex > fristLeftEnglishBracketIndex else fristLeftEnglishBracketIndex
    #print(leftIndex)
    if leftIndex == -1:
        return val
    else:
        val = str(val)
        return val[0:leftIndex]

def write_txt(content):
    with open(r'D:\YZX\application\Git\pytest\data.txt', 'a',encoding='utf-8') as f:
        f.write(content)

def read_excel():

    wb = xlrd.open_workbook(filename=fileLocation)#打开文件

    sheet1 = wb.sheet_by_name('2019年')#通过名字获取表格
    #print(sheet1.name,sheet1.nrows,sheet1.ncols)

    cols = sheet1.col_values(0)#获取列内容
    startIndex = 2
    for i in range(startIndex,len(cols)):
        thisFundNum = myParseInt(cols[i])
        thisFund = get_one_fund_info(thisFundNum)
        thisFundTxtInfo = thisFund[0] + ' ' +thisFund[1] + ' ' + thisFund[2] +  '\n'
        print(thisFundTxtInfo)
        write_txt(thisFundTxtInfo)
        time.sleep(1)

def get_one_fund_info(fundNum):
    url = 'http://fund.eastmoney.com/' + fundNum + '.html?spm=search'
    html = get_one_page(url)

    soup = BeautifulSoup(html,'lxml')
    body = soup.body
    mytext = body.find(style='float: left').get_text()
    myvalue = body.find(class_='dataItem02').find(class_='dataNums')
    #v1 = myvalue.find(class_='ui-font-large ui-color-green ui-num').get_text()
    v1 = myvalue.contents[0].get_text()
    #v2 = myvalue.find(class_='ui-font-middle ui-color-green ui-num').get_text()
    #print(type(myvalue))
    #print(v1)
    #print(v2)
    startIndex = mytext.rfind("(")
    #print(startIndex)
    fundName = mytext[0:startIndex]
    fundNum = mytext[startIndex+1:len(mytext)]
    #print('名字：'+fundName)
    #print('编号：'+fundNum)
    #print('净值：'+v1)
    fund = [fundName,fundNum,v1]
    return fund


read_excel()
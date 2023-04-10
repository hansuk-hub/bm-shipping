from selenium import webdriver
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains


import urllib.request
from bs4 import BeautifulSoup

from datetime import datetime

import random
import time
import os, sys
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from urllib.parse import quote_plus
import re

from pynput.keyboard import Key, Controller



from selenium import webdriver
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
import math
import urllib.request
from bs4 import BeautifulSoup

from datetime import datetime

import random
import time
import os, sys
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from urllib.parse import quote_plus
import re



api = 'gspread.json'

scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name(api, scope)
client = gspread.authorize(creds)

##########################################################################################

while (1):
    # getIndate = str(quote_plus(input("밀크런 달릴 날짜? ex)2020-03-15 : ")))
    # chkString = re.compile("\d{4}[\s-]?\d{2}[-]?\d{2}")
    # chkString = chkString.findall(getIndate)
    #
    # if chkString and (len(getIndate) == 10):
    #     print(str(chkString) + " 밀크런-쉽먼트를 신청합니다.")
    #     break
    # else:
    #     print("날짜를 다시 입력하세요")

    getIndate = '2023-04-17'
    break



options = webdriver.ChromeOptions()
# options.add_argument(r"--user-data-dir=C:\Users\tempe\AppData\Local\Google\Chrome\User Data") #e.g. C:\Users\You\AppData\Local\Google\Chrome\User Data

driver = webdriver.Chrome(executable_path=r'C:\works\chromedriver.exe', chrome_options=options)

def waitTime(stayTime):
    if stayTime == 's':
        ranVal = random.randint(1, 1)
    else:
        ranVal = random.randint(1, 2)
    time.sleep(ranVal)


def mainLogin():

    driver.get('https://supplier.coupang.com/')
    driver.find_element(By.NAME, 'username').send_keys('manyalittle')
    driver.find_element(By.NAME,'password').send_keys('wsjang555#')
    waitTime('s')
    driver.find_element(By.CLASS_NAME,'btn.btn-primary').click()  # 로그린 버튼 클릭
    waitTime('s')

    while 1 :
        try :
            # 페이지 갯수 알아내기
            driver.get(
                    'https://supplier.coupang.com/scm/purchase/order/list?page=1&searchDateType=WAREHOUSING_PLAN_DATE&searchStartDate=' + getIndate + '&searchEndDate=' + getIndate + '&centerCode=&purchaseOrderSeq=&vendorPaymentInfoSeq=&purchaseOrderStatus=')
            waitTime('l')
            html = driver.page_source
            soup = BeautifulSoup(html, 'html.parser')
            endPageNum = len(soup.select('#pagination ul li')) - 1


            ## 해당 입고일 발주서 URL 가져오기
            doc = client.open_by_url('https://docs.google.com/spreadsheets/d/1FR68zd-iqwmk1MxoMnHnIdgDkzHlHJnZv2Syk2iny7k/edit#gid=0')
            waitTime('s')
            manageSheet = doc.worksheet('발주리스트관리')


            try:
                urlData = manageSheet.find(getIndate)
            except:
                print('발주 관리 리스트에 해당 날짜가 없습니다.')
                driver.close()
                sys.exit()
            break
        except : continue

    global targetDateUrl
    while 1 :
        try :
            bulletinLastLine = len(manageSheet.col_values(1)) + 1
            for i in range(2, bulletinLastLine, 1):
                if manageSheet.cell(i, 1).value == getIndate:
                    targetDateUrl = manageSheet.cell(i, 2).value
                    # 입고일 파일로 이동 후 작업시작
                    doc = client.open_by_url(targetDateUrl)
                    waitTime('s')
                    break
                else:
                    targetDateUrl = "no data"
            break
        except : continue

    if (targetDateUrl == "no data") :
        print("일치하는 날짜가 없어 프로그램 종료합니다.")
        driver.close()
        sys.exit()

    while 1:
        try :
            targetDateUrl = manageSheet.cell(urlData.row,2 ).value
            break
        except:
            print("error : 발주리스트 관리 - 링크 Load error")
            exit()
            continue

    # print( targetDateUrl)
    print('해당 날짜의 시트 열기 성공')
    doc = client.open_by_url(targetDateUrl)
    waitTime('s')
    while 1 :
        try :
            allSheets = doc.worksheets()
            break
        except : continue

    print ( manageSheet)

    waitTime('s')

    checkedSheet = []
    for oneSheet in allSheets :
        if oneSheet.title != r'Bulletin' and oneSheet.title != r'전센터물량체크' :
            while 1 :
                try :
                    getSheet = doc.worksheet(oneSheet.title)
                    break
                except : continue
            if getSheet.cell(2,13).value is not None :
                print ( oneSheet.title + " 의 시트를 저장합니다.")
                checkedSheet.append(oneSheet.title)
    print ( checkedSheet )

    # 쉽먼트가 신청된 센터만 작업시작

    for pickCenter in checkedSheet :
        while 1 :
            try :
                targetCenter = doc.worksheet( pickCenter )
                line13th = targetCenter.col_values(13)
                line13th.remove('쉽먼트 박스 번호')

                break
            except : continue

    waitTime('s')

    line13th = list(filter(None, line13th))
    print( line13th)

    #제일큰 박스 번호를 알아낸다.
    maximumBox = 0
    setBoolean = 0

    for oneBox in line13th :
        oneBox = str(oneBox).strip()
        if str(oneBox).find("-") > 0 :
            doSplit = str(oneBox).split(',')

            for tempData in doSplit:
                pickBox = tempData.split('-')

                for tempData2 in pickBox:
                    setBoolean += 1
                    if int(maximumBox) < int(tempData2) and setBoolean % 2 == 1:
                        maximumBox = int(tempData2)

        else :
            if maximumBox < int(oneBox) :
                maximumBox = int(oneBox)

    print (maximumBox)
    waitTime('s')
mainLogin()

from selenium import webdriver
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By

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
import shutil

# from goto import goto, label

# 발주 관리 리스트
# https://docs.google.com/spreadsheets/d/1MFc8cukd7Cir3ZpUl27MfIj1m3iEc7UiypDx3nsYpAA/edit#gid=0

api = 'gspread.json'

scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name(api, scope)
client = gspread.authorize(creds)
global centerCnt

targetDir = r'C:\\works\\bm-shipping\\document\\milkrun\\'
downloadDir = 'C:/Users/tempe/Downloads/'



##########################################################################################

while (1):
    getIndate = str(quote_plus(input("발주서 파일 자동 다운로드 날짜? ex)2020-03-15 : ")))
    chkString = re.compile("\d{4}[\s-]?\d{2}[-]?\d{2}")
    chkString = chkString.findall(getIndate)

    if chkString and (len(getIndate) == 10):
        print(str(chkString) + " 발주서 파일 다운로드 합니다..")
        break
    else:
        print("날짜를 다시 입력하세요")

#############################################################################################


while True :
    try :
        driver: WebDriver = webdriver.Chrome('C:\works\chromedriver.exe')
        action = ActionChains(driver)
        break
    except :
        continue

def waitTime(stayTime):
    if stayTime == 's':
        ranVal = random.randint(1, 1)
    else:
        ranVal = random.randint(3, 5)
    time.sleep(ranVal)


def mainLogin():
    driver.get('https://supplier.coupang.com/')
    driver.find_element(By.NAME, 'username').send_keys('manyalittle')
    driver.find_element(By.NAME, 'password').send_keys('wsjang555#')
    waitTime('s')
    driver.find_element(By.CLASS_NAME, 'btn.btn-primary').click()  # 로그린 버튼 클릭
    waitTime('s')


    try:
        # 페이지 갯수 알아내기
        driver.get(
            'https://supplier.coupang.com/scm/purchase/order/list')
    except:  # 에러인 경우 한번 더 한다
        driver.get(
            'https://supplier.coupang.com/scm/purchase/order/list')
    waitTime('s')

    #입고예정일 클릭
    driver.find_element(By.ID,'searchDateType').click()
    action.key_down(Keys.ARROW_DOWN).perform()
    waitTime('s')

    #날짜 설정
    driver.find_element(By.XPATH,"//input[@name='searchStartDate']").clear()
    driver.find_element(By.XPATH,"//input[@name='searchStartDate']").send_keys(getIndate)
    driver.find_element(By.XPATH,"//input[@name='searchEndDate']").clear()
    driver.find_element(By.XPATH,"//input[@name='searchEndDate']").send_keys(getIndate)
    waitTime('s')
    #센터 설정
    getAllText = driver.find_element(By.XPATH,"//select[@id='centerCode']")
    getOneText = str(getAllText.text).replace(' ','').replace('\n',' ').split()


    #센터 이름 가져오기

    #
    doc = client.open_by_url(
        'https://docs.google.com/spreadsheets/d/1FR68zd-iqwmk1MxoMnHnIdgDkzHlHJnZv2Syk2iny7k/edit#gid=0')
    waitTime('s')
    manageSheet = doc.worksheet('발주리스트관리')

    try:
        urlData = manageSheet.find(getIndate)
    except:
        print('발주 관리 리스트에 해당 날짜가 없습니다.')
        driver.close()
        exit()

    global targetDateUrl
    global getCenters
    getCenters = []

    while 1 :
        try :
            bulletinLastLine = len(manageSheet.col_values(1)) + 1
            if bulletinLastLine > 40:
                stLine = 40
            elif bulletinLastLine <= 40:
                stLine = 2

            for i in range(stLine, bulletinLastLine, 1):

                if manageSheet.cell(i, 1).value == getIndate:
                    targetDateUrl = manageSheet.cell(i, 4).value

                    print ( targetDateUrl)
                    break
                else:
                    targetDateUrl = "no data"
            doc = client.open_by_url( targetDateUrl )
            break
        except : continue

    try:
        manageSheet = doc.worksheet('밀크런-쉽먼트')
    except:
        try:
            manageSheet = doc.worksheet('밀크런-쉽먼트')
        except:
            try:
                manageSheet = doc.worksheet('밀크런-쉽먼트')
            except:
                manageSheet = doc.worksheet('밀크런-쉽먼트')

    print ( manageSheet )
    print ( manageSheet.row_count )
    for i in range( 2,manageSheet.row_count ,1 ) :
        while True :
            try :
                checkData = manageSheet.cell(i ,2).value
                print(checkData)
                break
            except : continue


        if checkData == None :
            break
        getCenters.append( checkData )



##########################################################
    for getCenter in getCenters :
        centerCnt = 0
        clickCnt = 0

        for itema in getOneText :
            centerCnt += 1
            if (itema == getCenter) :
                print ( "# " + getCenter + '센터 클릭 ')
                break

        driver.find_element(By.XPATH,'//*[ @id="centerCode"]').click()
        driver.find_element(By.XPATH,'//*[ @id="centerCode"]/option[' + str( centerCnt ) +']').click()
        waitTime('s')

        #검색 버튼 클릭
        driver.find_element(By.ID, 'search').click()
        waitTime('s')
        dealCount = driver.page_source
        dealCount = ( dealCount.count('일반') -4)  /2
        print("   - 총 " + str(dealCount) + "개의 밀크런 접수내역 파일을 다운 받습니다.")
        waitTime('s')

        #전체 선택
        driver.find_element(By.NAME, 'checkAll').click()
        waitTime('s')

        # 다운로드
        driver.find_element(By.ID, 'btn-confirm-download').click()

        alert = driver.switch_to.alert
        if (alert):
            alert.accept()
            waitTime('l')

        #폴더 생성
        mp4Dir = targetDir + str(getCenter).strip()


        #파일 옮기기
        makeList = []
        for f_name in os.listdir(downloadDir):

            if f_name.startswith('PO_FOR_CONFIRM'):
                makeList.append(downloadDir + f_name)




        ## 파일 옮기기
        for getFile in makeList:
            print ( "new-file " + getFile)

            if getCenter == "인천13(RC)" :
                # old_name = targetDir + f_name
                # new_name = targetDir + "PO_FOR_CONFIRM_인천13-RC.xls"
                #
                # os.rename(old_name, new_name)
                mp4Dir = 'c:/coupang/bm-shipping/milkrun/인천13'
            shutil.move(getFile, mp4Dir)
#########################################################################################
    print("발주서 다운로드 끝.")

    driver.close()


# while True:
#     try:
#         mainLogin()
#         break
#     except:
#         continue
mainLogin()
# driver.close()


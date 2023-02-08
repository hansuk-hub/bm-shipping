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


# from goto import goto, label

# 발주 관리 리스트
# https://docs.google.com/spreadsheets/d/1MFc8cukd7Cir3ZpUl27MfIj1m3iEc7UiypDx3nsYpAA/edit#gid=0

api = 'gspread.json'

scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name(api, scope)
client = gspread.authorize(creds)

##########################################################################################

while (1):
    getIndate = str(quote_plus(input("밀크런 달릴 날짜? ex)2020-03-15 : ")))
    chkString = re.compile("\d{4}[\s-]?\d{2}[-]?\d{2}")
    chkString = chkString.findall(getIndate)

    if chkString and (len(getIndate) == 10):
        print(str(chkString) + " 밀크런 날짜 달립니다.")
        break
    else:
        print("날짜를 다시 입력하세요")

#############################################################################################


driver: WebDriver = webdriver.Chrome('C:\chromedriver.exe')

# def centerInfo():

def waitTime(stayTime):
    if stayTime == 's':
        ranVal = random.randint(1, 1)
    else:
        ranVal = random.randint(1, 2)
    time.sleep(ranVal)


def mainLogin():
    driver.get('https://supplier.coupang.com/')
    driver.find_element_by_name('username').send_keys('manyalittle')
    driver.find_element_by_name('password').send_keys('wsjang5566#')
    waitTime('s')
    driver.find_element_by_class_name('btn.btn-primary').click()  # 로그린 버튼 클릭
    waitTime('s')

    # 전세터 발주 파일 가져오기
    doc = client.open_by_url(
        'https://docs.google.com/spreadsheets/d/1FR68zd-iqwmk1MxoMnHnIdgDkzHlHJnZv2Syk2iny7k/edit#gid=0')
    waitTime('s')

    manageSheet = doc.worksheet('발주리스트관리')

    try:
        urlData = manageSheet.find(getIndate)
        # urlData = manageSheet.find('2020-07-30')
    except:
        print('발주요일이 없습니다. 프로그램을 중단합니다.')
        exit()

    while True:
        try:
            targetDateUrl = manageSheet.cell(urlData.row, 4).value
            print ( targetDateUrl)
            # targetDateUrl = 'https://docs.google.com/spreadsheets/d/1LEJ2NJ8PeIhaqveFhUDWm7cUsSMSjxg8yVLbUdCi0Ak/edit#gid=0'
            break
        except:
            print("error : 발주리스트 관리 - 링크 Load error")
            continue

    # 입고일 파일로 이동 후 작업시작
    errcnt = 1
    while 1 :
        try:
            doc = client.open_by_url(targetDateUrl)
            manageSheet = doc.worksheet('밀크런-쉽먼트')
            break
        except:
            if errcnt == 10 :
                print ('발주 확정 관리 리스트 링크를 확인해주세요')
                exit()
            errcnt += 1
            continue


    for i in range(23, 100, 1):
        while 1 :
            try:
                driver.get(
                    'https://supplier.coupang.com/milkrun/saveform')
                break
            except:
                waitTime('s')
                continue
        waitTime('s')

        # 체크박스 확인
        driver.find_element_by_xpath('//*[@id="checkbox"]').click()
        waitTime('s')
        driver.find_element_by_xpath('/html/body/div[4]/div[2]/div/div[3]/div/div/button').click()
        waitTime('s')

        getType = manageSheet.cell(i, 2).value
        getStatus = manageSheet.cell(i, 5).value  # 상태
        getCenter = manageSheet.cell(i, 2).value  # 센터 이름

        # 센터 없을경우 작업종료
        if getCenter == None:
            print('센터가 없음으로 작업종료')
            break  # 밀크런 작업 종료

        print ( getType)

        if getType != None :
            getPtCnt = manageSheet.cell(i, 3).value  # 발렛트 수량
            getQty = manageSheet.cell(i, 4).value  # 총입고 수량
            if getQty == None:
                print('입고 수량 표기 누락으로 작업 종료')
                exit()
                break  # 밀크런 작업 종료
            sendWeight = float(getQty) * 0.8 # 무게


            # 날짜 설정
            driver.find_element_by_id('warehousingPlannedAt').clear()
            driver.find_element_by_id('warehousingPlannedAt').send_keys(getIndate)

            waitTime('s')


            if getStatus != 'v' or getStatus == '불가능' :
                print(getCenter + " 센터 " + getPtCnt + " 개 팔렛트를 확정 시작 합니다")
                try:
                    # s1 = driver.find_element_by_xpath('//*[@id="centerCodeSearch"]')
                    # s1.send_keys(getCenter)
                    select = Select(driver.find_element_by_id('centerCodeSearch'))

                    # select by visible text
                    if getCenter == 'XRC01(RC)' :
                        getCenter = 'XRC01'
                    elif getCenter =='인천13(RC)' :
                        getCenter ='인천13'

                    select.select_by_visible_text(getCenter)
                    waitTime('s')
                except:
                    print("센터이름을 정확하게 기입해주세요. 프로그램 종료!")
                    exit()

                # 검색 버튼 클릭
                driver.find_element_by_id('search').click()
                waitTime('s')

                # 내용 입력 시작
                # 해당 발주서 클릭
                aaa = driver.find_elements_by_name('addPurchaseOrder')
                btnCnt = 0
                for bb in aaa:
                    bb.click()
                    btnCnt += 1
                waitTime('s')

                if btnCnt == 0:
                    print("센터에 접수할 발주서가 없어서 종료합니다.")
                    exit()

                # 출고지선택
                while True:
                    try:
                        driver.find_element_by_id('releaseAddressImport').click()
                        waitTime('s')
                        # driver.find_element_by_xpath("//button[@data-supplier-milkrun-location-seq='27819']").click()
                        driver.find_element_by_xpath("//button[@data-supplier-milkrun-location-seq='27820']").click()
                        waitTime('s')
                        break
                    except:
                        print("출고지 다시 클릭")
                        continue
                waitTime('s')
                # 총박스 수 입력
                driver.find_element_by_id('boxCountBox').send_keys(str(getPtCnt))
                # 박스 선택및 믹스 파렛트 수량 입력
                driver.find_element_by_id('box_3').send_keys(str(getPtCnt))
                driver.find_element_by_id('mixPltCount').send_keys(str(getPtCnt))

                # 수량 및 무게 입력

                driver.find_element_by_id('weightBox').send_keys(str(sendWeight))
                driver.find_element_by_id('contentsBox').send_keys('인형')

                # # 팔렛트 추가
                # driver.find_element_by_id('addPallet').click()
                # waitTime('s')
                # waitTime('s')
                #
                # # 팔렛트 수량 기입
                # driver.find_element_by_id('count_1').clear()
                # alert = driver.switch_to.alert
                # if (alert):
                #     alert.accept()
                #     waitTime('s')
                # driver.find_element_by_id('count_1').send_keys(getPtCnt)
                # waitTime('s')

                # 버튼 클릭 및 팔렛트 입력
                driver.find_element_by_xpath(
                    "//select[@id='allocationReplyBox']/option[@value='Y']").click()
                driver.find_element_by_xpath(
                    "//select[@id='loadUpBox']/option[@value='Y']").click()
                driver.find_element_by_xpath(
                    "//select[@id='forkliftBox']/option[@value='Y']").click()
                driver.find_element_by_xpath(
                    "//select[@id='pltRentalCompanyBox']/option[@value='아주팔레트']").click()
                driver.find_element_by_id('milkrunGuidCheck1').click()
                driver.find_element_by_id('milkrunGuidCheck2').click()
                driver.find_element_by_id('milkrunGuidCheck3').click()
                driver.find_element_by_id('milkrunGuidCheck4').click()

                dealCount = driver.page_source
                dealFindText = dealCount.count('밀크런 수정 가능한 시간이 마감되었습니다.')

                if dealFindText > 0:
                    print('밀크런 신청 불가능 센터(또는 신청시간 마감) 입니다.')
                    manageSheet.update_cell(i, 4, "불가능")
                elif dealFindText == 0:
                    driver.find_element_by_id('saveMilkrun').click()
                    waitTime('s')
                    waitTime('s')
                    waitTime('s')
                    driver.find_element_by_xpath('//*[@id="happyFeedback"]/i').click()
                    waitTime('s')

                # 확정 상태 체크
                manageSheet.update_cell(i, 5, "v")

                # 저장버튼
                alert = driver.switch_to.alert
                alert.dismiss()
                waitTime('s')
                print('--- ' + getCenter + ' 센터 확정 되었습니다')
        elif getType == None :
            print ("더이상 작업할 센터가 없습니다.")


    waitTime('s')

    print("밀크런 달리기 끝.")

    driver.close()


mainLogin()
# while True :
#     try :
#         mainLogin()
#         break
#     except : continue


# driver.close()


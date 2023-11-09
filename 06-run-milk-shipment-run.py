from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
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


options = webdriver.ChromeOptions()
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option("useAutomationExtension", False)
service = ChromeService(executable_path=r'C:\works\chromedriver.exe')
driver = webdriver.Chrome(service=service, options=options)

# def centerInfo():

def waitTime(stayTime):
    if stayTime == 's':
        ranVal = random.randint(1, 1)
    else:
        ranVal = random.randint(1, 2)
    time.sleep(ranVal)


def mainLogin():
    driver.get('https://supplier.coupang.com/')
    driver.find_element(By.NAME,'username').send_keys('manyalittle')
    driver.find_element(By.NAME,'password').send_keys('wsjang555#')
    waitTime('s')
    driver.find_element(By.CLASS_NAME,'btn.btn-primary').click()  # 로그린 버튼 클릭
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


    for i in range(23, 54, 1):
        while 1 :
            try:
                driver.get(
                    'https://supplier.coupang.com/milkrun/saveform')
                waitTime('s')
                # 체크박스 확인
                driver.find_element(By.XPATH, '//label[@for="checkbox"]').click()
                # driver.find_element(By.XPATH,'//*[label]').click()
                waitTime('s')
                driver.find_element(By.XPATH,'/html/body/div[4]/div[2]/div/div[3]/div/div/button').click()
                waitTime('s')
                break
            except Exception as e:  # 모든 예외의 에러 메시지를 출력할 때는 Exception을 사용
                print('예외가 발생했습니다.', e)
                waitTime('s')
                continue
        waitTime('s')


        while 1 :
            try :

                getType = manageSheet.cell(i, 2).value
                getStatus = manageSheet.cell(i, 5).value  # 상태
                getCenter = manageSheet.cell(i, 2).value  # 센터 이름
                break
            except : continue

        # 센터 없을경우 작업종료
        if getCenter == None:
            print('더이상 확정 센터가 없음으로 작업종료')
            break  # 밀크런 작업 종료

        print ( getType)

        if getType != None :
            while 1 :
                try :

                    getQty = manageSheet.cell(i, 3).value  # 총 박스 수량 (C열)
                    get2boxCnt = manageSheet.cell(i, 6).value  # 2호 박스 수량
                    get3boxCnt = manageSheet.cell(i, 7).value  # 3호 박스 수량
                    break
                except :
                    print ("구글 시트로 부터 데이터 로딩중 오류 - Don't worry")
                    continue

            if getQty == None:
                print('입고 수량 표기 누락으로 작업 종료')
                exit()
                break  # 밀크런 작업 종료
            # sendWeight = float(getQty) * 0.8 # 무게


            # 날짜 설정
            driver.find_element(By.ID,'warehousingPlannedAt').clear()
            driver.find_element(By.ID,'warehousingPlannedAt').send_keys(getIndate)

            waitTime('s')


            if getStatus != 'v' or getStatus == '불가능' :
                print(getCenter + " 센터 " + getQty + " 개 박스를 확정 시작 합니다")
                try:
                    # s1 = driver.find_element(By.XPATH,'//*[@id="centerCodeSearch"]')
                    # s1.send_keys(getCenter)
                    select = Select(driver.find_element(By.ID,'centerCodeSearch'))

                    # select by visible text
                    if getCenter == 'XRC01(RC)':
                        getCenter = 'XRC01'
                    elif getCenter == 'XRC02(RC)':
                        getCenter = 'XRC02'
                    elif getCenter == 'XRC03(RC)':
                        getCenter = 'XRC03'
                    elif getCenter == '인천13(RC)':
                        getCenter = '인천13'
                    elif getCenter == '곤지암2(RC)':
                        getCenter = '곤지암2'
                    elif getCenter == '경기광주1':
                        getCenter = '광주'

                    select.select_by_visible_text(getCenter)
                    waitTime('s')
                except:
                    print("센터이름을 정확하게 기입해주세요. 프로그램 종료!")
                    exit()

                # 검색 버튼 클릭
                driver.find_element(By.ID,'search').click()
                waitTime('s')

                # 내용 입력 시작
                # 해당 발주서 클릭
                aaa = driver.find_elements(By.NAME, 'addPurchaseOrder')
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
                        driver.find_element(By.ID,'releaseAddressImport').click()
                        waitTime('s')
                        driver.find_element(By.XPATH,"//button[@data-supplier-milkrun-location-seq='27820']").click()
                        waitTime('s')
                        waitTime('s')
                        waitTime('s')
                        break
                    except:
                        print("출고지 다시 클릭")
                        continue

                # 수량 및 무게 입력력

                driver.find_element(By.ID,'boxCountBox').send_keys(str(getQty))    #총 박스 수량 (발주확정시트 C열의 값)
                driver.find_element(By.ID,'weightBox').send_keys(str(10))    #총 중량
                driver.find_element(By.ID,'mixPltCount').send_keys(str(1))   #믹스 팔렛트 수량
                driver.find_element(By.ID,'contentsBox').send_keys('인형')
                driver.find_element(By.ID,'box_2').send_keys(str(get2boxCnt))
                driver.find_element(By.ID,'box_3').send_keys(str(get3boxCnt))


                # 버튼 클릭 및 팔렛트 입력
                driver.find_element(By.XPATH,
                    "//select[@id='allocationReplyBox']/option[@value='Y']").click()
                driver.find_element(By.XPATH,
                    "//select[@id='loadUpBox']/option[@value='Y']").click()
                driver.find_element(By.XPATH,
                    "//select[@id='forkliftBox']/option[@value='Y']").click()
                driver.find_element(By.XPATH,
                    "//select[@id='pltRentalCompanyBox']/option[@value='아주팔레트']").click()
                driver.find_element(By.ID,'milkrunGuidCheck1').click()
                driver.find_element(By.ID,'milkrunGuidCheck2').click()
                driver.find_element(By.ID,'milkrunGuidCheck3').click()
                driver.find_element(By.ID,'milkrunGuidCheck4').click()

                dealCount = driver.page_source
                dealFindText = dealCount.count('밀크런 수정 가능한 시간이 마감되었습니다.')


                if dealFindText > 0:
                    print('밀크런 신청 불가능 센터(또는 신청시간 마감) 입니다.')
                    print("작업중 문제!!!!!!!!!!!!!!!!!!!!!")
                    manageSheet.update_cell(i, 4, "불가능")
                elif dealFindText == 0:
                    driver.find_element(By.ID,'saveMilkrun').click()
                    waitTime('s')
                    waitTime('s')
                    waitTime('s')
                    driver.find_element(By.XPATH,'//*[@id="happyFeedback"]/i').click()
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

    print("밀크런 - 쉽먼트 달리기 끝.")

    driver.close()


mainLogin()
# while True :
#     try :
#         mainLogin()
#         break
#     except : continue


# driver.close()
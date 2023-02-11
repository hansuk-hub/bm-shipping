from selenium import webdriver
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.common.keys import Keys
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


# from goto import goto, label

# 발주 관리 리스트
# https://docs.google.com/spreadsheets/d/1MFc8cukd7Cir3ZpUl27MfIj1m3iEc7UiypDx3nsYpAA/edit#gid=0

api = 'gspread.json'

scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name(api, scope)
client = gspread.authorize(creds)

##########################################################################################



while(1) :
   getIndate = str(quote_plus(input ("쉽먼트 달릴 날짜? ex)2020-03-15 : ")))
   chkString = re.compile("\d{4}[\s-]?\d{2}[-]?\d{2}")
   chkString = chkString.findall(getIndate)

   if chkString and (len(getIndate)== 10 ):
       print(str(chkString) + " 쉽먼트 날짜 달립니다.")
       break
   else :
       print("날짜를 다시 입력하세요")
   # getIndate = '2021-07-15'
   break

#############################################################################################


driver: WebDriver = webdriver.Chrome('C:\chromedriver.exe')
# options = webdriver.ChromeOptions()
# options.binary_location = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
# chrome_driver_path = r"C:\work\python\auto-worker\coupang\venv\chromedriver.exe"
#
# driver = webdriver.Chrome(chrome_driver_path, chrome_options=options)


def waitTime(stayTime):
   if stayTime == 's':
       ranVal = random.randint(1, 1)
   else:
       ranVal = random.randint(1, 2)
   time.sleep(ranVal)



def mainLogin() :
    driver.get('https://supplier.coupang.com/')
    driver.find_element_by_name('username').send_keys('manyalittle')
    driver.find_element_by_name('password').send_keys('wsjang5566#')
    waitTime('s')
    driver.find_element_by_class_name('btn.btn-primary').click()  # 로그린 버튼 클릭
    waitTime('s')

    complex = {}
    complex = dict()



    #전세터 발주 파일 가져오기
    doc = client.open_by_url(
       'https://docs.google.com/spreadsheets/d/1FR68zd-iqwmk1MxoMnHnIdgDkzHlHJnZv2Syk2iny7k/edit#gid=0')
    waitTime('s')

    manageSheet = doc.worksheet('발주리스트관리')

    try :
        urlData = manageSheet.find(string(getIndate))
        # urlData = manageSheet.find('2020-07-30')
    except :
       print('발주요일이 없습니다. 프로그램을 중단합니다.')
       exit()

    while 1:
        try :
            targetDateUrl = manageSheet.cell(urlData.row, ).value
            break
        except:
            print("error : 발주리스트 관리 - 링크 Load error")
            exit()
            continue


    # 입고일 파일로 이동 후 작업시작
    doc = client.open_by_url(targetDateUrl)
    try :
        manageSheet = doc.worksheet('밀크런-쉽먼트')
    except :
        try :
            manageSheet = doc.worksheet('밀크런-쉽먼트')
        except :
            try:
                manageSheet = doc.worksheet('밀크런-쉽먼트')
            except:
                manageSheet = doc.worksheet('밀크런-쉽먼트')

    print ('관리시트를 찾았습니다. 다음으로 진행합니다.')

    for findShipmentText in range( 23, 55, 1) :

       #  발주 타입중 쉽먼트를 가져온다.

       for LastCenter in range ( findShipmentText  , 100, 1) :
           while True:
               try:
                   isEmpty = manageSheet.cell(LastCenter, 2).value
                   break
               except : continue

           if isEmpty == None :
               print('작업 완료')
               exit()
               break

           try:
               # 쉽먼트 페이지 이동
               driver.get('https://supplier.coupang.com/ibs/asn/active')
           except:  # 에러인 경우 한번 더 한다
               driver.get('https://supplier.coupang.com/ibs/asn/active')
           waitTime('s')

           # 쉽먼트 신청 클릭
           while True:
               try:
                   driver.find_element_by_xpath('//*[@id="createShipmentBtn"]').click()
                   break
               except: continue

           waitTime('s')
           # 출고지 센터 클릭
           while True:
               try:
                   driver.find_element_by_xpath('//*[@id="goCreate"]').click()
                   break
               except:  continue
           waitTime('s')

           while True:
               try:
                   getStatus = manageSheet.cell(LastCenter , 5).value  # 상태 가져오기
                   break
               except:  continue


           #송장 번호 배열 생성

           getPtCnt = 0
           deliveryCompany = 0
           boxCurrent = 1
           if getStatus != 'v':
               # 센터 이름
               getCenter = manageSheet.cell(LastCenter , 2).value

               #박스 수량 알아내기
               getBoxCnt = manageSheet.cell(LastCenter , 3).value

               print(getCenter + " 센터 / " + str(getBoxCnt) + " 박스 작업 합니다")

               globals()[getCenterData] = {'a': 100, 'b': 200}


               #혼재내역
               for bq in range(8, 16,1):
                   while true :
                       try :
                            getSKUinfo = manageSheet.cell(LastCenter , bq).value
                            break
                       except : continue
                   if getSKUinfo != None :
                       splitData = getSKUinfo.split(',')
                       for getNano in splitData:
                           getCenterData = { getNano.strip() : boxCurrent }  #현재의 스큐정보가 몇번 박스 인지 저장한다.
                   elif getSKUinfo == None :
                       break    # 더이상 박스 SKU 수집하지 않음
                   print (getCenterData)
                   boxCurrent += 1



               # 센터이름 바꾸기
               if getCenter == '평택' :
                   getCenter = '평택1'


               while True :
                   try :
                       waitTime('s')
                       # 센터 클릭
                       driver.find_element_by_xpath(
                           "//select[@name='fcCode']/option[text()='" + getCenter + "']").click()
                       break
                   except :
                       continue

               driver.find_element_by_id('edd').click()
               driver.find_element_by_id('edd').clear()
               driver.find_element_by_id('edd').send_keys(getIndate)

               waitTime('s')
               driver.find_element_by_id('searchPO').click()
               driver.find_element_by_id('check-all').click()
               waitTime('s')
               driver.find_element_by_id('check-all').click()
               waitTime('s')
               driver.find_element_by_id('check-all').click()
               waitTime('s')
               driver.find_element_by_name('confirmed').click()
               waitTime('s')

               #  박스 수량 입력
               driver.find_element_by_id('parcelCount').clear()
               driver.find_element_by_id('parcelCount').send_keys(getBoxCnt)
               driver.find_element_by_name('confirmed').click()
               waitTime('s')
               waitTime('s')


               # 상품 갯수 알아내기
               # proRows = len(driver.find_elements_by_xpath("//*[@id='skuSelectTable']/table/tbody/tr"))
               rowCnt = 1
               currentBoxCnt =1

               # 한박스 넘을때만 작업
               if getBoxCnt > 1 :

                   # 한박스에 여러개의 SKU 가 있는 상황

                   for confirmBox in sperateBoxCode :
                        getTableSku = driver.find_element_by_xpath('/html[1]/body/div[1]/div[2]/div[2]/div[2]/div[4]/table/tbody/tr[' + str(rowCnt) + ']/td[2]').text
                        if getTableSku == confirmBox :
                            driver.find_element_by_xpath(
                                '/html[1]/body[1]/div[1]/div[2]/div[2]/div[2]/div[4]/table/tbody/tr[' + str(
                                    rowCnt) + ']/td[7]/div[' + str(
                                    targetBox) + ']/select/option[' + str(currentBoxCnt) + ']').click()  # 박스 선택
                        getIntQty = driver.find_element_by_xpath('/html[1]/body/div[1]/div[2]/div[2]/div[2]/div[4]/table/tbody/tr[' + str(rowCnt) + ']/td[6]').text

                        driver.find_element_by_xpath(
                            '/html[1]/body[1]/div[1]/div[2]/div[2]/div[2]/div[4]/table/tbody/tr[' + str(rowCnt) + ']/td[8]/input').clear()  # 기존수량 삭제

                        driver.find_element_by_xpath('/html[1]/body[1]/div[1]/div[2]/div[2]/div[2]/div[4]/table/tbody/tr[' + str(rowCnt) + ']/td[8]/input').send_keys(
                                getIntQty)  # 수량입력



                        rowCnt += 1
                   currentBoxCnt += 1





               driver.find_element_by_name('confirmed').click()
               waitTime('s')
               waitTime('s')

               #택배사 클릭
               # driver.find_element_by_id('deliveryCompany_chosen').click()
               # driver.find_element_by_id('deliveryCompany').click()
               driver.find_element_by_xpath('/html[1]/body[1]/div[1]/div[2]/div[2]/div[2]/div[3]/div/dl[1]/dd/div/a').click()
               waitTime('s')

               action = ActionChains(driver)

               driver.find_element_by_class_name('chosen-search-input').click()
               waitTime('s')


               # for aaa in range(0, deliveryCompany) :
               action.key_down(Keys.DOWN).pause(1).perform()



               action.send_keys(Keys.RETURN).perform()

               waitTime('s')

               #발송 시간 클릭
               driver.find_element_by_id('shipTime').click()
               waitTime('s')
               driver.find_element_by_xpath('/html[1]/body[1]/div[1]/div[2]/div[2]/div[2]/div[3]/div/dl[2]/dd/div[2]/div/span').click()
               waitTime('s')

               for inputInvoice in range ( 0, len(getInvoiceArray) - 1) :
                    if getInvoiceArray[inputInvoice] == '' :
                        print ('송장번호가 입력이 안되어 있습니다')
                        driver.close()
                        break

                    while 1 :
                        try :
                            print('\t 송장번호 입력 : '+ getInvoiceArray[inputInvoice])
                            driver.find_element_by_xpath(
                                '/html[1]/body[1]/div[1]/div[2]/div[2]/div[2]/div[4]/table/tbody/tr[' + str(
                                    inputInvoice + 1) + ']/td[2]/input').clear()
                            driver.find_element_by_xpath('/html[1]/body[1]/div[1]/div[2]/div[2]/div[2]/div[4]/table/tbody/tr['+ str(inputInvoice+1 ) + ']/td[2]/input').send_keys( getInvoiceArray[inputInvoice] )
                            waitTime('s')
                            break
                        except : continue
               driver.find_element_by_name('confirmed').click()
               waitTime('s')
               while True :
                   try :
                       alert = driver.switch_to.alert
                       if (alert):
                           alert.accept()
                           waitTime('s')
                       break
                   except : continue


               while True :
                    try :
                        manageSheet.update_cell(LastCenter , 15, 'v')
                        break
                    except : continue

               print(getCenter + " 작업 완료!")


               waitTime('s')
       break


    print("쉽먼트 달리기 끝.")

    driver.close()


mainLogin()
# while True :
#     try :
#         mainLogin()
#         break
#     except : continue



# driver.close()

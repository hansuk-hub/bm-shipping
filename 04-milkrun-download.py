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
import shutil

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
    getIndate = str(quote_plus(input("밀크런 파일 자동 다운로드 날짜? ex)2020-03-15 : ")))
    # getIndate = '2020-10-22'
    chkString = re.compile("\d{4}[\s-]?\d{2}[-]?\d{2}")
    chkString = chkString.findall(getIndate)

    if chkString and (len(getIndate) == 10):
        print(str(chkString) + " 밀크런 날짜 달립니다.")
        break
    else:
        print("날짜를 다시 입력하세요")

#############################################################################################


# driver: WebDriver = webdriver.Chrome('C:\chromedriver.exe')

options = webdriver.ChromeOptions()

options.add_experimental_option("prefs", {
  "download.default_directory": r"C:\coupang\bm-shipping\milkrun",
  "download.prompt_for_download": False,
  "download.directory_upgrade": False,
  "plugins.always_open_pdf_externally": True,
  "safebrowsing.enabled": False
})

driver = webdriver.Chrome(chrome_options=options, executable_path="C:\chromedriver.exe")


action = ActionChains(driver)

def waitTime(stayTime):
    if stayTime == 's':
        ranVal = random.randint(1, 1)
    else:
        ranVal = random.randint(5, 7)
    time.sleep(ranVal)


def mainLogin():
    driver.get('https://supplier.coupang.com/')
    driver.find_element_by_name('username').send_keys('manyalittle')
    driver.find_element_by_name('password').send_keys('wsjang5566#')
    waitTime('s')
    driver.find_element_by_class_name('btn.btn-primary').click()  # 로그린 버튼 클릭
    waitTime('s')

    try:
        # 페이지 갯수 알아내기
        driver.get(
            'https://supplier.coupang.com/milkrun/milkrunList')
    except:  # 에러인 경우 한번 더 한다
        driver.get(
            'https://supplier.coupang.com/milkrun/milkrunList')
    waitTime('s')

    driver.find_element_by_name('startDate').clear()
    driver.find_element_by_name('startDate').send_keys(getIndate)
    driver.find_element_by_name('endDate').clear()
    driver.find_element_by_name('endDate').send_keys(getIndate)
    driver.find_element_by_id('search').click()

    waitTime('s')

    dealCount = driver.page_source
    dealCount = dealCount.count('PALLET')

    targetDir = 'c:/coupang/bm-shipping/milkrun/'
    downloadDir = 'c:/coupang/bm-shipping/milkrun/'

    print ( "총 "+ str(dealCount) + "개의 밀크런 접수내역 파일을 다운 받습니다.")
    for i in range(0, dealCount, 1):
        try :
            postionChk = driver.find_element_by_xpath('//*[@id="milkrunListTable"]/tbody/tr[' + str(i + 1) + ']/td[19]/span[6]')
            clickPosition = 4
        except :
            clickPosition = 3


        print("클릭 포지션 체크 : " + str(clickPosition))
        driver.find_element_by_xpath(
                '//*[@id="milkrunListTable"]/tbody/tr[' + str(i + 1) + ']/td[19]/span[' + str(clickPosition) +']').click()
        getCenter = driver.find_element_by_xpath('//*[@id="milkrunListTable"]/tbody/tr[' + str(i + 1) + ']/td[10]').text
        driver.switch_to.window(driver.window_handles[1])

        waitTime('l')
        #프린트 취소
        print ('저장 합니다.')

        action.key_down(Keys.TAB).pause(2).key_down(Keys.ENTER).perform()

        waitTime('l')

        driver.switch_to.window(driver.window_handles[1])
        driver.close()


        waitTime('s')

        # 폴더 생성
        mp4Dir = targetDir + str(getCenter).strip()
        print ( mp4Dir )

        try:
            os.mkdir(mp4Dir)
            print('   - 폴더 생성 완료')
        except:
            print('폴더 생성 실패 - 위치 확인 바랍니다')
            exit()

        driver.switch_to.window(driver.window_handles[0])
        driver.find_element_by_xpath('//*[@id="milkrunListTable"]/tbody/tr[' + str(i + 1) + ']/td[19]/span[' + str(clickPosition+2) +']').click()  # 로그린 버튼 클릭

        waitTime('s')

        chkDelete = 0

        # 파일 옮기기
        makeList = []
        for f_name in os.listdir(downloadDir):
            if f_name.endswith('pdf'):
                makeList.append(downloadDir + f_name)

        ## 파일 옮기기
        for getFile in makeList:
            if chkDelete == 0 :
                shutil.move(getFile, mp4Dir)
                chkDelete += 1
            else :
                os.remove( getFile )




        # 파일 옮기기
        makeList = []
        for f_name in os.listdir(downloadDir):
            if f_name.endswith('xlsx'):
                makeList.append(downloadDir + f_name)

        ## 파일 옮기기
        for getFile in makeList:
            shutil.move(getFile, mp4Dir)






    waitTime('s')

    print("밀크런 달리기 끝.")

    driver.close()


mainLogin()


# driver.close()


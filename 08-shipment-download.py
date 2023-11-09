from selenium import webdriver
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

from bs4 import BeautifulSoup
import random
import time
import os

from urllib.parse import quote_plus
import re
import shutil


##########################################################################################

## 생성될 위치를 꼭 미리 만들어 놓는다.

targetDir = 'C:\\works\\bm-shipping\\document\\shipment\\'
downloadDir = 'C:\\Users\\tempe\\Downloads\\'
##########################################################################################

while (1):
    getIndate = str(quote_plus(input("쉽먼트 파일 자동 다운로드 날짜? ex)2020-03-15 : ")))
    chkString = re.compile("\d{4}[\s-]?\d{2}[-]?\d{2}")
    chkString = chkString.findall(getIndate)

    if chkString and (len(getIndate) == 10):
        print(str(chkString) + " 쉽먼트 날짜 파일 다운 받습니다.")
        break
    else:
        print("날짜를 다시 입력하세요")

#############################################################################################


driver: WebDriver = webdriver.Chrome('C:\works\chromedriver.exe')


def waitTime(stayTime):
    if stayTime == 's':
        ranVal = random.randint(1, 1)
    else:
        ranVal = random.randint(5, 7)
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
            'https://supplier.coupang.com/ibs/asn/active')
        waitTime('s')
    except:  # 에러인 경우 한번 더 한다
        driver.get(
            'https://supplier.coupang.com/ibs/asn/active')
        waitTime('s')

    driver.find_element(By.ID,'edd').click()
    driver.find_element(By.ID,'edd').clear()
    driver.find_element(By.ID,'edd').send_keys(getIndate)
    waitTime('s')
    driver.find_element(By.ID,'edd').send_keys(Keys.ENTER)

    waitTime('l')

    html = driver.page_source
    soup = BeautifulSoup(html, 'html.parser')
    endPageNum = len(soup.select('nav ul li a')) - 2

    curPage = 2

    for i in range(0, endPageNum, 1):
        waitTime('s')
        newPage = driver.page_source

        if newPage.count('발송 완료') > 0 :
            pageDownCount = newPage.count('발송 완료') + 1
            btnPostion1 = 2
            btnPostion2 = 3
        else :
            pageDownCount = newPage.count('발송 가능') + 1
            btnPostion1 = 4
            btnPostion2 = 5

        for curPosition in range(1, pageDownCount, 1):
            getCenterName = driver.find_element(By.XPATH,
                '//*[@id="parcel-tab"]/tbody/tr[' + str(curPosition) + ' ]/td[7]').text
            mp4Dir = targetDir + str(getCenterName).strip()
            print(mp4Dir)
            try:
                os.mkdir(mp4Dir)
                print('  - 폴더 생성 완료')
            except:
                print('폴더 생성 실패 - 위치 확인 바랍니다')
                exit()

            # 임시 테스트   - 발송 완료 상태
            # driver.find_element(By.XPATH,'//*[@id="parcel-tab"]/tbody/tr['+ str(curPosition) +']/td[10]/button[2]').click()
            #
            # waitTime('s')
            # driver.find_element(By.XPATH,
            #     '//*[@id="parcel-tab"]/tbody/tr[' + str(curPosition) + ']/td[10]/button[3]').click()
            # waitTime('s')

            # 기존것
            driver.find_element(By.XPATH,
                '//*[@id="parcel-tab"]/tbody/tr[' + str(curPosition) + ']/td[10]/button[' + str(btnPostion1) + ']').click()
            waitTime('s')
            driver.find_element(By.XPATH,
                '//*[@id="parcel-tab"]/tbody/tr[' + str(curPosition) + ']/td[10]/button[' + str(btnPostion2) + ']').click()
            waitTime('s')

            ## 다운로드 파일 찾기
            makeList = []
            for f_name in os.listdir(downloadDir):
                if f_name.startswith('shipment'):
                    makeList.append(downloadDir + f_name)

            ## 파일 옮기기
            for getFile in makeList:
                shutil.move(getFile, mp4Dir)

        curPage = curPage + 2
        if (endPageNum > 1):  driver.find_element(By.XPATH,'//*[@id="parcel-pagination"]/ul/li[' + str(curPage) + ']/a').click()

    waitTime('l')

    print("쉽먼트 달리기 끝.")

    driver.close()


mainLogin()

# driver.close()



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

api = 'gspread.json'

scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name(api, scope)
client = gspread.authorize(creds)


##########################################################################################
def find_index(data, target):
    res = []
    lis = data
    while True:
        try:
            res.append(lis.index(target) + (res[-1] + 1 if len(res) != 0 else 0))
            lis = data[res[-1] + 1:]
        except:
            break
    return res


while (1):
    getIndate = str(quote_plus(input("밀크런 달릴 날짜? ex)2020-03-15 : ")))
    chkString = re.compile("\d{4}[\s-]?\d{2}[-]?\d{2}")
    chkString = chkString.findall(getIndate)

    if chkString and (len(getIndate) == 10):
        print(str(chkString) + " 밀크런-쉽먼트를 신청합니다.")
        break
    else:
        print("날짜를 다시 입력하세요")

    # getIndate = '2023-04-17'
    break

options = webdriver.ChromeOptions()
# options.add_argument(r"--user-data-dir=C:\Users\tempe\AppData\Local\Google\Chrome\User Data") #e.g. C:\Users\You\AppData\Local\Google\Chrome\User Data
options.add_argument("--window-size=1700,1000")
driver = webdriver.Chrome(executable_path=r'C:\works\chromedriver.exe', chrome_options=options)


def waitTime(stayTime):
    if stayTime == 's':
        ranVal = random.randint(1, 1)
    else:
        ranVal = random.randint(1, 2)
    time.sleep(ranVal)


def mainLogin():
    action = ActionChains(driver)
    driver.get('https://supplier.coupang.com/')
    driver.find_element(By.NAME, 'username').send_keys('manyalittle')
    driver.find_element(By.NAME, 'password').send_keys('wsjang555#')
    waitTime('s')
    driver.find_element(By.CLASS_NAME, 'btn.btn-primary').click()  # 로그린 버튼 클릭
    waitTime('s')

    while 1:
        try:
            # 페이지 갯수 알아내기
            driver.get(
                'https://supplier.coupang.com/scm/purchase/order/list?page=1&searchDateType=WAREHOUSING_PLAN_DATE&searchStartDate=' + getIndate + '&searchEndDate=' + getIndate + '&centerCode=&purchaseOrderSeq=&vendorPaymentInfoSeq=&purchaseOrderStatus=')
            waitTime('l')
            html = driver.page_source
            soup = BeautifulSoup(html, 'html.parser')
            endPageNum = len(soup.select('#pagination ul li')) - 1

            ## 해당 입고일 발주서 URL 가져오기
            doc = client.open_by_url(
                'https://docs.google.com/spreadsheets/d/1FR68zd-iqwmk1MxoMnHnIdgDkzHlHJnZv2Syk2iny7k/edit#gid=0')
            waitTime('s')
            manageSheet = doc.worksheet('발주리스트관리')

            try:
                urlData = manageSheet.find(getIndate)
            except:
                print('발주 관리 리스트에 해당 날짜가 없습니다.')
                driver.close()
                sys.exit()
            break
        except:
            continue

    global targetDateUrl
    while 1:
        try:
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
        except:
            continue

    if (targetDateUrl == "no data"):
        print("일치하는 날짜가 없어 프로그램 종료합니다.")
        driver.close()
        sys.exit()

    while 1:
        try:
            targetDateUrl = manageSheet.cell(urlData.row, 2).value
            break
        except:
            print("error : 발주리스트 관리 - 링크 Load error")
            exit()
            continue

    # print( targetDateUrl)
    print('해당 날짜의 시트 열기 성공')
    doc = client.open_by_url(targetDateUrl)
    waitTime('s')
    while 1:
        try:
            allSheets = doc.worksheets()
            break
        except:
            continue

    waitTime('s')

    checkedSheet = []
    for oneSheet in allSheets:
        if oneSheet.title != r'Bulletin' and oneSheet.title != r'전센터물량체크':
            while 1:
                try:
                    getSheet = doc.worksheet(oneSheet.title)
                    if getSheet.cell(2, 13).value is not None:
                        print(oneSheet.title + " 의 시트를 저장합니다.")
                        checkedSheet.append(oneSheet.title)
                    break
                except:
                    continue


    # 쉽먼트가 신청된 센터만 작업시작

    for pickCenter in checkedSheet:
        while 1:
            try:
                targetCenter = doc.worksheet(pickCenter)
                line13th = targetCenter.col_values(13)
                line13th.remove('쉽먼트 박스 번호')

                break
            except:
                continue

        line13th = list(filter(None, line13th))

        # 제일큰 박스 번호를 알아낸다.
        maximumBox = 0
        setBoolean = 0

        for oneBox in line13th:
            oneBox = str(oneBox).strip()
            if str(oneBox).find("-") > 0:
                doSplit = str(oneBox).split(',')

                for tempData in doSplit:
                    pickBox = tempData.split('-')

                    for tempData2 in pickBox:
                        setBoolean += 1
                        if int(maximumBox) < int(tempData2) and setBoolean % 2 == 1:
                            maximumBox = int(tempData2)

            else:
                if maximumBox < int(oneBox):
                    maximumBox = int(oneBox)

        # print(maximumBox)
        waitTime('s')

        ############ 쉽먼트 작업 시작

        try:
            # 쉽먼트 페이지 이동
            driver.get('https://supplier.coupang.com/ibs/asn/active')
        except:  # 에러인 경우 한번 더 한다
            driver.get('https://supplier.coupang.com/ibs/asn/active')
        waitTime('s')

        while True:
            try:
                getStatus = targetCenter.cell(1, 14).value  # 상태 가져오기
                break
            except:
                continue

        # 미확정 상태라면 작업시작
        if getStatus != 'v':
            # 쉽먼트 신청 클릭
            while True:
                try:
                    driver.get('https://supplier.coupang.com/ibs/asn/active')
                    driver.find_element(By.XPATH, '//*[@id="createShipmentBtn"]').click()
                    break
                except:
                    continue

            waitTime('s')
            # 출고지 센터 클릭
            while True:
                try:
                    driver.find_element(By.XPATH, '//*[@id="goCreate"]').click()
                    break
                except:
                    continue
            waitTime('s')

            print(pickCenter + " 센터 / " + str(maximumBox) + " 박스 작업 합니다")

            # 센터이름 바꾸기
            if pickCenter == 'XRC01(RC)':
                pickCenter = 'XRC01'
            elif pickCenter == '인천13(RC)':
                pickCenter = '인천13'
            elif pickCenter == 'XRC02(RC)':
                pickCenter = 'XRC02'
            elif pickCenter == 'XRC03(RC)':
                pickCenter = 'XRC03'
            elif pickCenter == 'XRC04(RC)':
                pickCenter = 'XRC04'
            elif pickCenter == 'XRC05(RC)':
                pickCenter = 'XRC05'
            elif pickCenter == '곤지암2(RC)':
                pickCenter = '곤지암2'




            while True:
                try:
                    waitTime('s')
                    # 센터 클릭
                    driver.find_element(By.XPATH,
                                        "//select[@name='fcCode']/option[text()='" + pickCenter + "']").click()
                    break
                except:
                    continue

            driver.find_element(By.ID, 'edd').click()
            driver.find_element(By.ID, 'edd').clear()
            driver.find_element(By.ID, 'edd').send_keys(getIndate)

            waitTime('s')
            driver.find_element(By.ID, 'searchPO').click()
            driver.find_element(By.ID, 'check-all').click()
            waitTime('s')
            driver.find_element(By.ID, 'check-all').click()
            waitTime('s')
            driver.find_element(By.ID, 'check-all').click()
            waitTime('s')
            driver.find_element(By.NAME, 'confirmed').click()
            waitTime('s')

            #  박스 수량 입력
            driver.find_element(By.ID, 'parcelCount').clear()
            driver.find_element(By.ID, 'parcelCount').send_keys(maximumBox)
            driver.find_element(By.NAME, 'confirmed').click()
            waitTime('s')
            waitTime('s')

            # 상품 갯수 알아내기
            proRows = len(driver.find_elements(By.XPATH, "//*[@id='skuSelectTable']/table/tbody/tr"))

            # 구글 드라이브에 입력된 모든 발주번호 가져오기
            while 1:
                try:
                    getAllbill = targetCenter.col_values(7)
                    break
                except:
                    continue
            # 구글 드라이브에 입력된 모든 SKU 가져오기
            # getAllSKU = targetCenter.col_values(2)

            for trCnt in range(1, (proRows + 1)):
                getInputCnt = driver.find_element(By.XPATH,
                                                  '/html[1]/body/div[1]/div[2]/div[2]/div[2]/div[4]/table/tbody/tr[' + str(
                                                      trCnt) + ']/td[6]').text
                print("\t 수량 입력 " + str(trCnt) + " : " + getInputCnt)

                if int(maximumBox) == 1:  # 박스 수량 1개일때
                    driver.find_element(By.XPATH,
                                        '/html/body/div[1]/div[2]/div[2]/div[2]/div[4]/table/tbody/tr[' + str(
                                            trCnt) + ']/td[8]/input').clear()
                    driver.find_element(By.XPATH,
                                        '/html/body/div[1]/div[2]/div[2]/div[2]/div[4]/table/tbody/tr[' + str(
                                            trCnt) + ']/td[8]/input').send_keys(getInputCnt)
                elif int(maximumBox) > 1:  # 박스 신청 수량 2개 넘을때
                    ###################################################################################
                    print(' 다중 박스 입력 시작')
                    # 서플라이에서 한개의 발주번호가져오기
                    lineBill = driver.find_element(By.XPATH,
                                                   '/html[1]/body/div[1]/div[2]/div[2]/div[2]/div[4]/table/tbody/tr[' + str(
                                                       trCnt) + ']/td[1]').text
                    lineSKU = driver.find_element(By.XPATH,
                                                  '/html[1]/body/div[1]/div[2]/div[2]/div[2]/div[4]/table/tbody/tr[' + str(
                                                      trCnt) + ']/td[2]').text
                    # 가져온 발주번호를 드라이브에서 찾기
                    filteredBill = find_index(getAllbill, lineBill)

                    for loopI in filteredBill:
                        while 1:
                            try:
                                tempData3 = targetCenter.cell(int(loopI) + 1, 2).value
                                break
                            except:
                                continue
                        if tempData3 == lineSKU:
                            while 1:
                                try:
                                    getBoxData = targetCenter.cell(int(loopI) + 1, 13).value
                                    break
                                except:
                                    continue

                            if getBoxData.find(',') > 0:
                                print(getBoxData)
                                commaSplit = getBoxData.split(',')

                                currentBoxCnt = len(commaSplit)
                                waitTime('s')
                                # 박스 분할
                                if currentBoxCnt > 1:
                                    driver.find_element(By.XPATH,
                                                        '/html[1]/body[1]/div[1]/div[2]/div[2]/div[2]/div[4]/table[1]/tbody[1]/tr[' + str(
                                                            trCnt) + ']/td[7]/div[1]/button[1]').click()
                                    waitTime('s')
                                    driver.find_element(By.ID, 'splitCount').send_keys(currentBoxCnt)
                                    driver.find_element(By.ID, 'split').click()
                                    waitTime('s')

                                # 수량 나눠서 입력하기
                                for targetBox in range(1, currentBoxCnt + 1):
                                    getBoxInfo = commaSplit[targetBox - 1].strip()
                                    print(getBoxInfo + " ==> 선택합니다.")
                                    dashSplit = getBoxInfo.split('-')

                                    driver.find_element(By.XPATH,
                                                        '/html[1]/body[1]/div[1]/div[2]/div[2]/div[2]/div[4]/table/tbody/tr[' + str(
                                                            trCnt) + ']/td[7]/div[' + str(
                                                            targetBox) + ']/select/option[' + str(
                                                            dashSplit[0]) + ']').click()  # 박스 선택
                                    driver.find_element(By.XPATH,
                                                        '/html[1]/body[1]/div[1]/div[2]/div[2]/div[2]/div[4]/table/tbody/tr[' + str(
                                                            trCnt) + ']/td[8]/input[' + str(
                                                            targetBox) + ']').clear()  # 기존수량 삭제

                                    driver.find_element(By.XPATH,
                                                        '/html[1]/body[1]/div[1]/div[2]/div[2]/div[2]/div[4]/table/tbody/tr[' + str(
                                                            trCnt) + ']/td[8]/input[' + str(targetBox) + ']').send_keys(
                                        str(dashSplit[1]))  # 수량입력

                                waitTime('s')
                            elif getBoxData.find(',') < 0:
                                getBoxData = getBoxData.strip()
                                boxNumber = getBoxData.split('-')
                                while 1:
                                    try:
                                        driver.find_element(By.XPATH,
                                                            '/html/body/div[1]/div[2]/div[2]/div[2]/div[4]/table/tbody/tr[' + str(
                                                                trCnt) + ']/td[8]/input').clear()

                                        driver.find_element(By.XPATH,
                                                            '/html[1]/body[1]/div[1]/div[2]/div[2]/div[2]/div[4]/table/tbody/tr[' + str(
                                                                trCnt) + ']/td[7]/div/select/option[' + str(
                                                                boxNumber[0]) + ']').click()  # 박스 선택

                                        driver.find_element(By.XPATH,
                                                            '/html/body/div[1]/div[2]/div[2]/div[2]/div[4]/table/tbody/tr[' + str(
                                                                trCnt) + ']/td[8]/input').send_keys(getInputCnt)
                                        break
                                    except Exception as e:
                                        print('큰 오류는 아니지만 박스 선택중 오류 : ' + e)
                                        waitTime('s')
                                        continue

            driver.find_element(By.NAME, 'confirmed').click()

            waitTime('s')

            # 택배사 클릭
            # driver.find_element_by_id('deliveryCompany_chosen').click()
            # driver.find_element_by_id('deliveryCompany').click()

            # driver.find_element(By.XPATH,
            #     '/html[1]/body[1]/div[1]/div[2]/div[2]/div[2]/div[3]/div/dl[1]/dd/div/a').click()
            driver.execute_script("window.scrollTo(0, 0);")
            waitTime('s')

            driver.find_element(By.XPATH, '//a[@class="chosen-single"]').click()
            # driver.find_element(By.XPATH, '/html/body/div[1]/div[2]/div[2]/div[2]/div[3]/div/dl[1]/dd/div/a').click()

            waitTime('s')

            action.send_keys('밀').perform()
            # action.send_keys(Keys.BACKSPACE).perform()

            action.send_keys(Keys.RETURN).perform()

            waitTime('s')

            # 발송 시간 클릭
            driver.find_element(By.ID, 'shipTime').click()
            waitTime('s')

            for xCnt in range(1, maximumBox + 1, 1):
                while 1:
                    try:
                        driver.find_element(By.XPATH,
                                            '/html[1]/body/div[1]/div[2]/div[2]/div[2]/div[4]/table/tbody/tr[' + str(
                                                xCnt) + ']/td[2]/select').click()
                        driver.find_element(By.XPATH,
                                            '/html[1]/body/div[1]/div[2]/div[2]/div[2]/div[4]/table/tbody/tr[' + str(
                                                xCnt) + ']/td[2]/select/option[' + str(xCnt + 1) + ']').click()
                        break
                    except:
                        continue
                waitTime('s')

            driver.find_element(By.NAME, 'confirmed').click()
            waitTime('s')
            while True:
                try:
                    alert = driver.switch_to.alert
                    if (alert):
                        alert.accept()
                        waitTime('s')
                    break
                except:
                    continue

            while True:
                try:
                    targetCenter.update_cell(1, 14, 'v')
                    break
                except:
                    continue

            print(pickCenter + " 작업 완료!")

    waitTime('s')
    driver.close()


mainLogin()
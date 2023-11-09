from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import urllib.request

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
    global getIndate
    getIndate = str(quote_plus(input("입고일을 입력하세요 ex)2020-03-31 : ")))
    # getIndate = "2020-07-16"
    chkString = re.compile("\d{4}[\s-]?\d{2}[-]?\d{2}")
    chkString = chkString.findall(getIndate)

    # getInTargetNum = quote_plus(input("발주번호 입력 : "))
    break
    # if chkString and (len(getIndate)== 10 ):
    #     print(str(chkString) + " 데이터를 취합 시작 합니다.")
    #     break
    # else :
    #     print("날짜를 다시 입력하세요")

#############################################################################################


options = webdriver.ChromeOptions()
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option("useAutomationExtension", False)
service = ChromeService(executable_path=r'C:\works\chromedriver.exe')

driver = webdriver.Chrome(service=service, options=options)
onlyCheckGo = 0


def waitTime(stayTime):
    if stayTime == 's':
        ranVal = random.randint(1, 1)
    else:
        ranVal = random.randint(2, 2)
    time.sleep(ranVal)


def mainLogin(getfTargetNum, getIndate):
    ## 해당 입고일 발주서 URL 가져오기
    doc = client.open_by_url(
        'https://docs.google.com/spreadsheets/d/1MFc8cukd7Cir3ZpUl27MfIj1m3iEc7UiypDx3nsYpAA/edit#gid=0')
    waitTime('s')

    manageSheet = doc.worksheet('발주리스트관리')
    urlData = manageSheet.find(getIndate)
    while True:
        try:
            targetDateUrl = manageSheet.cell(urlData.row, 2).value
            break
        except:
            print("error : 발주리스트 관리 - 링크 Load error - wait moment")
            continue

    # 입고일 파일로 이동 후 작업시작
    doc = client.open_by_url(targetDateUrl)
    waitTime('s')

    # 발주리스트 페이지로 이동
    # for fi in range(1,dealCount +1 , 1) :
    #     getOrderNumber = driver.find_element_by_xpath('/html/body/div[1]/div[2]/div[2]/div[3]/table/tbody/tr['+ str(fi) + ']/td[2]/a')
    #     getOrderCenter = driver.find_element_by_xpath('/html/body/div[1]/div[2]/div[2]/div[3]/table/tbody/tr['+ str(fi) + ']/td[11]')
    #     getOrderQty = driver.find_element_by_xpath('/html/body/div[1]/div[2]/div[2]/div[3]/table/tbody/tr['+ str(fi) + ']/td[13]')
    #     getOrderMoney = driver.find_element_by_xpath('/html/body/div[1]/div[2]/div[2]/div[3]/table/tbody/tr['+ str(fi) + ']/td[15]')
    #     getNowState = driver.find_element_by_xpath('/html/body/div[1]/div[2]/div[2]/div[3]/table/tbody/tr['+ str(fi) + ']/td[4]')
    #
    #
    #     dealCenter.append(getOrderCenter.text)   #센터 저장
    #     dealBox.append(getOrderNumber.text)     #발주 번호 리스트 저장     # 지속된 오류로 코드 변경
    #     dealQty.append(getOrderQty.text)        #발주 수량 저장
    #     dealMoney.append(getOrderMoney.text)    #발주 금액 저장
    #     dealState.append(getNowState.text)       #발주 상태 저장 - 발주 확정 요청 이외의 상황에서 -3 으로 변경해야한다.

    while True:
        try:  # 시트를 선택하거나, 생성한다
            BulletinSheet = doc.worksheet('Bulletin')
            break
        except:
            continue
    # waitTime('s')
    getfTargetNum = getfTargetNum.replace('C', '')
    getfTargetNum = getfTargetNum.replace(' ', '')

    while True:
        try:
            targetCenterRow = BulletinSheet.find(getfTargetNum)
            break
        except:
            print("error : Bulltein에서 deal number load error")
            continue

    # 센터 이름
    targetCenterName = BulletinSheet.cell(targetCenterRow.row, 2).value
    print("작업센터 : " + targetCenterName + " / " + getfTargetNum)
    confirmSheet = doc.worksheet(targetCenterName)
    now = str(datetime.now())[0:19]  # 시간저장

    CenterDeal = confirmSheet.findall(getfTargetNum)
    GDBarcodeRoom = []  # 취합된 GD의 바코드 배열
    GDDealQtyRoom = []  # 발주 수량
    GDInQtyRoom = []  # 입고 가능 수량
    GDInWayRoom = []  # 입고 방법
    GDQtyMode = []  #

    # 해당 센터의 전체 내용을 가져온다
    for curData in CenterDeal:

        while True:
            try:
                values_list = confirmSheet.row_values(curData.row)
                break
            except:
                print("error : Get Row values")
                continue

        while True:
            try:
                getBarcode = confirmSheet.cell(curData.row + 1, 3).value
                break
            except:
                print("error : barcode  Into")
                continue

        getInQty = int(values_list[11])
        getDealQty = int(values_list[5])
        getWay = values_list[12]

        if getWay == "":
            print(str(curData.row) + " : 입고 방법이 비어 있습니다.")
            exit()

        if (len(GDBarcodeRoom) > 1):
            if GDInWayRoom[0] != getWay:
                print("출고 방법이 다릅니다. 확인해주세요")
                driver.close()
                exit()

        GDBarcodeRoom.append(getBarcode)
        GDDealQtyRoom.append(getDealQty)
        GDInQtyRoom.append(getInQty)
        GDInWayRoom.append(getWay)

        if (getDealQty > getInQty):  # 수량이 모자르게 입고 되는 경우
            GDQtyMode.append('less')
        elif (getDealQty == getInQty):  # 발주 수량과 입고 수량이 동일
            GDQtyMode.append('good')
        elif (getDealQty < getInQty):
            print(getBarcode + " Error  - 입고수량 확인 요청!!")
            driver.close()
            exit()

    #  입고 유형선택
    if (getWay == '택배') or (getWay == '쉽먼트'):
        transportType = 'SHIPMENT'
    elif getWay == '트럭':
        transportType = 'TRUCK'
    elif getWay == '밀크런':
        transportType = 'MILKRUN'
    else:
        transportType = 'PARCEL'

    driver.get('https://supplier.coupang.com/scm/purchase/order/get/' + str(getfTargetNum))
    dealCount = driver.page_source
    dealCount = dealCount.count('예약완료')

    if dealCount >= 1 :
        print("이미 발주 확정된 상태 입니다.")
        driver.close()
        exit()

    # dealCount = driver.page_source
    # onlyCheckGo = dealCount.count('상품명/이미지/바코드')
    # waitTime('s')

    # for fi in range(1,endPageNum, 1) :
    driver.get('https://supplier.coupang.com/scm/purchase/order/modify/' + str(getfTargetNum))
    waitTime('s')




    driver.find_element(By.XPATH,"//select[@name='transportType']/option[@value='" + transportType + "']").click()
    waitTime('s')
    driver.find_element(By.XPATH,"//select[@name='transportType']/option[@value='" + transportType + "']").click()

    allsum = 0  # 총 발주 수량
    codeChk = 0  # 상품명과 바코드 저장시 토글읠 위한 변수
    barCodeBox = []

    spName = driver.find_elements(By.CLASS_NAME, 'list-group-item')
    for item in spName:
        chkText = item.text
        if (chkText.find('회송됩니다.') > 1):
            break
        if (item.text != '직매입') and (item.text != '과세'):
            if codeChk == 0:
                barCodeBox.append(item.text)  # 바코드 저장
                codeChk = 1
            else:
                # proNameBox.append(item.text)  # 상품명 저장
                codeChk = 0

    # 수량 입력20
    for ai in range(0, len(GDBarcodeRoom), 1):
        findRoom = barCodeBox.index(GDBarcodeRoom[ai])
        makeLineNum = (ai * 2) + 1
        print("barcode Postion : " + str(ai))
        # insertQty = driver.find_element_by_xpath("/html/body/div/div[2]/div[2]/table[6]/tbody/tr[" + str( makeLineNum ) + "]/td[7]/input")
        insertQty = driver.find_element(By.XPATH,
            "//tbody/tr[" + str(makeLineNum) + "]/td[7]/input")
        insertQty.clear()
        # 수량 입력
        insertQty.send_keys(GDInQtyRoom[findRoom])

        allsum = allsum + GDInQtyRoom[findRoom]  # 수량 입력 0일때 체크

        # 수량 모자란경우
        if (GDQtyMode[ai] == 'less'):
            # 사유 입력
            driver.find_element(By.XPATH,
                "//tr[" + str(makeLineNum + 1) + "]/td[2]/button").click()
            # "/html/body/div/div[2]/div[2]/table[5]/tbody/tr["+ str(makeLineNum + 1) +"]/td[2]/button").click()
            waitTime('s')
            waitTime('s')
            driver.find_element(By.XPATH,
                "//select[@name='inSufficiencyCause1']/option[@value='32']").click()
            waitTime('s')
            # driver.find_element_by_xpath('/html/body/div/div[2]/div[2]/div[7]/div[2]/div/div[3]/button[3]').click()
            driver.find_element(By.XPATH, "//button[3]").click()
            waitTime('s')
    # // *[ @ id = "causeModal"] / div[2] / div / div[3] / button[3]

    # 발주 확정
    driver.find_element(By.CSS_SELECTOR,
        '#app > div.contentsArea > div.scmContentsArea > div.tabBottom > button.btn.btn-warning.btn-save > span').click()
    waitTime('s')
    alert = driver.switch_to.alert
    alert.accept()
    waitTime('l')

    if allsum == 0:  # 발주 합계가 0 인경우
        alert = driver.switch_to.alert
        alert.accept()
        # driver.close()
    else:
        # 체크박스 클릭
        getChkBox = driver.find_elements(By.CLASS_NAME, 'checkAgree')
        for clickBox in getChkBox:
            clickBox.click()

        # 체크박스 확인용
        for clickBox in getChkBox:
            if (clickBox.get_attribute("checked") == 0):
                clickBox.click()
        waitTime('s')

        # 버튼 클릭하며 마무리
        driver.find_element(By.ID, 'btnPartnerAgree').click()
        waitTime('s')

        while True:
            try:
                alert = driver.switch_to.alert
                break
            except:
                print("error : last confirm error")
                continue

        alert.accept()
        waitTime('s')
        alert = driver.switch_to.alert
        alert.dismiss()
        waitTime('s')

    print("작업이 완료 되었습니다.")


# driver = webdriver.Chrome()

driver.get('https://supplier.coupang.com/')
driver.find_element(By.NAME, 'username').send_keys('manyalittle')
driver.find_element(By.NAME,'password').send_keys('wsjang555#')
waitTime('s')
driver.find_element(By.CLASS_NAME,'btn.btn-primary').click()  # 로그린 버튼 클릭
waitTime('s')

def autoConfirm(getIndate):
    ## 해당 입고일 발주서 URL 가져오기
    doc = client.open_by_url(
        'https://docs.google.com/spreadsheets/d/1FR68zd-iqwmk1MxoMnHnIdgDkzHlHJnZv2Syk2iny7k/edit#gid=0')
    waitTime('s')

    manageSheet = doc.worksheet('발주리스트관리')
    urlData = manageSheet.find(getIndate)

    while True:
        try:
            targetDateUrl = manageSheet.cell(urlData.row, 2).value

            # targetDateUrl = 'https://docs.google.com/spreadsheets/d/1LEJ2NJ8PeIhaqveFhUDWm7cUsSMSjxg8yVLbUdCi0Ak/edit#gid=0'
            print(targetDateUrl)
            break
        except:
            print("error : 발주리스트 관리 - 링크 Load error / don't worry")
            continue

    # 입고일 파일로 이동 후 작업시작
    doc = client.open_by_url(targetDateUrl)
    waitTime('s')

    # 시트를 선택하거나, 생성한다
    while True:
        try:
            BulletinSheet = doc.worksheet('Bulletin')
            break
        except:
            print("error : BulletinSheet Select Error - don't worry")
            continue

    makeAllDealNumber = []

    while True:
        try:
            curLastLine = len(BulletinSheet.col_values(2)) + 1
            break
        except:
            print("error : BulletinSheet Get Row values - don't worry")
            continue

    for ia in range(2, curLastLine):
        time.sleep(1)
        while True:
            try:
                values_list = BulletinSheet.row_values(ia)
                print(values_list)


                values_list = BulletinSheet.row_values(ia)
                print(values_list)

                print("test1")
                getInTargetNum = values_list[2]
                getInType = values_list[7]
                getCenterName = values_list[1]
                break
            except:
                print('발주서에 입고 유형이 기입 안되어 있는지 확인 바랍니다.')
                print("error : BulletinSheet Data Current Row Read Error - don't worry ")
                continue

        if getInType == '':
            print('Error : 입고 유형이 기입되지 않았습니다.')
            exit()

        try:  # 완료 상태 v가 비어 있으면 error 발생 함으로 따로 처리한다
            getConfrimSt = values_list[8]
        except:
            getConfrimSt = "X"

        print(
            " BulletinSheetget Read : " + getInTargetNum + " " + getInType + " " + getCenterName + " " + "/ 확정 상태 :" + getConfrimSt)

        if getConfrimSt != "v":
            ########################################################################################

            # targetCenterName = BulletinSheet.cell(targetCenterRow.row, 2).value
            print("Auto Confirm Start : Center " + getCenterName + " / " + getInTargetNum)
            confirmSheet = doc.worksheet(getCenterName)  # 타겟 센터 이동 후 저장
            now = str(datetime.now())[0:19]  # 시간저장




            CenterDeal = confirmSheet.findall(getInTargetNum)
            GDBarcodeRoom = []  # 취합된 GD의 바코드 배열
            GDDealQtyRoom = []  # 발주 수량
            GDInQtyRoom = []  # 입고 가능 수량
            GDInWayRoom = []  # 입고 방법
            GDQtyMode = []  #

            # 해당 센터의 전체 내용을 가져온다
            for curData in CenterDeal:
                while True:
                    try:
                        values_list = confirmSheet.row_values(curData.row)
                        break
                    except:
                        print("error : Get Row values - don't worry")
                        continue

                while True:
                    try:
                        getBarcode = confirmSheet.cell(curData.row + 1, 3).value
                        break
                    except:
                        print("error : barcode  Into  - don't worry")
                        continue

                getInQty = int(values_list[11])  # 사용자가 입력한 입고 수량 저장
                if getInQty == None :
                    print('입고 수량을 기입하지 않아 프로그램을 종료합니다')
                    driver.close()
                    exit()

                # print("원규가 저장한 입고수량 :" + str(getInQty))

                getDealQty = int(values_list[5])  # 발주서가 원하는 수량 저장

                # getWay = values_list[12]          # 입고 방법 저장

                # if getWay == "":
                #     print(str(curData.row) + " : 입고 방법이 비어 있습니다.")
                #     exit()

                # if (len(GDBarcodeRoom) > 1):
                #     if GDInWayRoom[0] != getWay:
                #         print("출고 방법이 다릅니다. 확인해주세요")
                #         driver.close()
                #         exit()

                GDBarcodeRoom.append(getBarcode)
                GDDealQtyRoom.append(getDealQty)
                GDInQtyRoom.append(getInQty)
                GDInWayRoom.append(getInType)

                if (getDealQty > getInQty):  # 수량이 모자르게 입고 되는 경우
                    GDQtyMode.append('less')
                elif (getDealQty == getInQty):  # 발주 수량과 입고 수량이 동일
                    GDQtyMode.append('good')
                elif (getDealQty < getInQty):
                    print(getBarcode + " Error  - 입고수량 확인 요청!!")
                    exit()

            #  입고 유형선택
            if getInType == '택배' or getInType == '쉽먼트':
                transportType = 'SHIPMENT'
            elif getInType == '트럭':
                transportType = 'TRUCK'
            elif getInType == '밀크런':
                transportType = 'MILKRUN'
            elif getInType == '밀크런-쉽먼트':
                transportType = 'MILKRUN'
            else:
                print("Error - 입고 방법 명시 오류 : 센터명 - " + getCenterName + " / 발주번호 : " + getInTargetNum)
                exit()
                transportType = 'PARCEL'

            # for fi in range(1,endPageNum, 1) :
            driver.get('https://supplier.coupang.com/scm/purchase/order/modify/' + getInTargetNum)
            waitTime('s')

            pageSource = driver.page_source

            if ( pageSource.count('잘못된') > 1 ) :
                print ('잘못된 발주서 상태 입니다. 발주서가 확정되어 있을수 있으니 확인해주세요!')
                exit()

            # 크롬에서 수정 버튼 클릭하여, 발주 확정 작업으로 넘어간다.
            # while 1:
            #     try:
            #         driver.find_element_by_class_name('btn.btn-default').click()
            #         waitTime('s')
            #         break
            #
            #     except:
            #         continue




            dealCount = driver.page_source
            dealFindText = dealCount.count('예약완료') + dealCount.count('잘못된 발주서수정화면으로') + dealCount.count('발주서 조회에')
            # 확정 상태인지 한번더 확인하는 구문

            if dealFindText >= 1:
                print("이미 발주 확정 및 확정 불가된 상태 발주 입니다.")
                ## 에러 리포트 작성하기
            #                 driver.close()
            else:


                # 가끔 이미 거래처 수정 상태로 체크박스만 눌러야 하는 상태가 있다.
                if onlyCheckGo == 0:




                    allsum = 0  # 총 발주 수량
                    codeChk = 0  # 상품명과 바코드 저장시 토글읠 위한 변수
                    barCodeBox = []

                    spName = driver.find_elements(By.CLASS_NAME, 'list-group-item')
                    for item in spName:
                        chkText = item.text
                        if (chkText.find('회송됩니다.') > 1):
                            break
                        if (item.text != '직매입') and (item.text != '과세'):
                            if codeChk == 0:
                                barCodeBox.append(item.text)  # 바코드 저장
                                codeChk = 1
                            else:
                                # proNameBox.append(item.text)  # 상품명 저장
                                codeChk = 0

                    # # 발주 확정 클릭
                    # while True:
                    #     try:
                    #         driver.find_element_by_css_selector(
                    #             '#app > div.contentsArea > div.scmContentsArea > div.tabBottom > button.btn.btn-warning.btn-save > span').click()
                    #         break
                    #     except:
                    #         waitTime('s')
                    #         continue
                    # waitTime('l')

                    ## 입고 유형 선택택
                    while True:
                       try:
                            driver.find_element(By.XPATH,
                                "//select[@name='transportType']/option[@value='" + transportType + "']").click()
                            break
                       except:
                            print("센터에 입고 " + transportType + " 방법이 없습니다. ")
                            sys.exit()


                    # 가끔 입고 유형이 잘못될때가 있다. 그래서 다시 한번 입고 유형을 클릭한다.
                    while True:
                        try:
                            driver.find_element(By.XPATH,
                                "//select[@name='transportType']/option[@value='" + transportType + "']").click()
                            break
                        except:
                            continue


                    # 반품 주소지 변경
                    # while True:
                    #     try:
                    #         getReturnAddress = driver.find_element_by_name('returnContact')
                    #         getReturnAddress.clear()
                    #         waitTime('s')
                    #         getReturnAddress.send_keys('변경하는 전화번호')
                    #         getReturnAddress = driver.find_element_by_name('returnLocation')
                    #         getReturnAddress.clear()
                    #         waitTime('s')
                    #         getReturnAddress.send_keys('변경할 주소지')
                    #         break
                    #     except :
                    #         continue

                    # 수량 입력
                    curUrl = driver.current_url  # 페이지 리로드를 위해 URL 저장
                    while True:
                        try:
                            for ai in range(0, len(GDBarcodeRoom), 1):
                                findRoom = barCodeBox.index(GDBarcodeRoom[ai])
                                makeLineNum = (ai * 2) + 1
                                print("barcode Postion -> " + str(ai))

                                while True:
                                    try:
                                        insertQty = driver.find_element(By.XPATH,
                                            "//tbody/tr[" + str(makeLineNum) + "]/td[7]/input")
                                        insertQty.clear()
                                        # 수량 입력
                                        insertQty.send_keys(GDInQtyRoom[findRoom])
                                        break
                                    except:
                                        continue

                                allsum = allsum + GDInQtyRoom[findRoom]  # 수량 입력 0일때 체크
                                waitTime('s')
                                # 수량 모자란경우
                                if (GDQtyMode[ai] == 'less'):
                                    # 사유 입력
                                    driver.find_element(By.XPATH,
                                        "//tr[" + str(makeLineNum + 1) + "]/td[2]/button").click()
                                    waitTime('s')
                                    driver.find_element(By.XPATH,
                                        "//select[@name='inSufficiencyCause1']/option[@value='32']").click()
                                    waitTime('s')
                                    driver.find_element(By.XPATH,"//button[3]").click()
                                    waitTime('s')

                            driver.find_element(By.CSS_SELECTOR,
                                '#app > div.contentsArea > div.scmContentsArea > div.tabBottom > button.btn.btn-warning.btn-save > span').click()

                            waitTime('s')

                            alert = driver.switch_to.alert
                            if (alert):
                                alert.accept()
                            waitTime('l')
                            ################## 끝
                            break
                        except:
                            driver.get(curUrl)
                            continue

                if allsum == 0:  # 발주 합계가 0 인경우
                    while True:
                        try :
                            alert = driver.switch_to.alert
                            break
                        except :
                            waitTime('l')
                            continue


                    if (alert):
                        alert.accept()


                else:
                    curUrl = driver.current_url  # 페이지 리로드를 위해 URL 저장
                    while True:
                        try:
                            # 새로 로드하여 확정한다.
                            driver.get(curUrl)  # 저장한 확정 페이지 URL로 이동한다
                            getSource = driver.page_source
                            getConfirmStatus = getSource.count('[발주확정] 업체 발주확정')

                            if getConfirmStatus >= 1:
                                print('이미 발주확정 상태입니다')
                                break

                   #          while True:
                   #              driver.get(curUrl)  # 저장한 확정 페이지 URL로
                   #              waitTime('l')
                   #
                   #              # 전체 체크 박스 클릭 ( 최종동의 )
                   #              # try :
                   #              #     driver.find_element_by_name('checkAgreeAll').click()
                   #              #     if (driver.find_element_by_name('checkAgreeAll').is_selected() != True):
                   #              #         waitTime('s')
                   #              #         a = 0/ 0 #일부러 에러를 일으켜 다시 체크하도록 한다.
                   #              #     print("최종체크박스 클릭중 에러러")
                   #              #     break
                   #              # except:
                   #              #     waitTime('s')
                   #              #     continue
                   #
                   # #-------------# 기존 체크 진행하기 시작---------------------------------------
                   #              checkCnt = 0
                   #              try:
                   #                  # 체크 박스 클릭
                   #
                   #                  ###
                   #
                   #                  checkboxes = driver.find_element(By.XPATH, ".//*[@type='checkbox']")
                   #                  for checkbox in checkboxes:
                   #                      checkbox.click()
                   #                      checkCnt = checkCnt + 1
                   #                  getSource = driver.page_source
                   #                  isMax6 = getSource.count('시즌상품')
                   #                  isMax6 = isMax6 + getSource.count('사전 발주서 확인')
                   #                  if isMax6 > 0:
                   #                      maxCheck = 8
                   #                  else:
                   #                      maxCheck = 7
                   #                  if checkCnt < maxCheck:
                   #                      print("체크박스 체크 오류 페이지 리로드 ")
                   #                      a = 0 / 0  # 에러를 발생 시킨다
                   #
                   #                  elif checkCnt == maxCheck:  # 5개 또는 6개 정상 체크시
                   #                      break
                   #              except:
                   #                  waitTime('s')
                   #                  continue

                            # ############################################ 체크박스 클릭 + 1
                            # checkboxes = driver.find_elements(By.XPATH,"//input[@name='checkAgree']")
                            #
                            # for checkbox in checkboxes:
                            #     if (checkbox.is_selected() != True):
                            #         checkbox.click()
                            # waitTime('s')
                            # ############################################ 체크박스 클릭 + 2
                            # checkboxes = driver.find_elements(By.XPATH,"//input[@name='checkAgree']")
                            #
                            # for checkbox in checkboxes:
                            #     if (checkbox.is_selected() != True):
                            #         checkbox.click()
                    # -------------# 기존 체크 진행하기 끝---------------------------------------

                            driver.find_element(By.XPATH, "//input[@name='checkAgreeAll']").click()

                            ###################################
                            waitTime('s')

                            # 버튼 클릭하며 마무리
                            driver.find_element(By.ID, 'btnPartnerAgree').click()
                            waitTime('l')
                            while True:
                                try:
                                    # alert = driver.switch_to.alert
                                    # alert.accept()

                                    driver.find_element(By.CLASS_NAME,'btn.primary-button.btn-lg').click()
                                    waitTime('s')
                                    break
                                except:
                                    continue

                            # # 밀크런 일때만 해임 버튼을 누룬다
                            # if transportType == 'MILKRUN':
                            #     driver.find_element_by_class_name('btn.primary-button.btn-lg').click()
                            #     # alert = driver.switch_to.alert
                            #     # alert.dismiss()
                            #     waitTime('s')
                            # break
                        except:
                            waitTime('l')
                            driver.get(curUrl)
                            continue

            print(getInTargetNum + " 작업이 완료 되었습니다.")

        while True:
            try:
                BulletinSheet.update_cell(ia, 9, "v")
                break
            except:
                print("error : v check error- don't worry.")
                continue

        ###########################################################################################


# while True :
#     try :
#         autoConfirm(getIndate)
#         break
#     except :
#         continue

autoConfirm(getIndate)
# autoConfirm(getIndate)


closecnt = 0

if driver:
    while True:
        closecnt = closecnt + 1
        if closecnt == 20:  # 브라우져 닫을려는 횟수가 20 넘으면 아웃
            break
        try:
            driver.close()
            break
        except:
            print("error : close ")
            continue

print("작업이 완료되었습니다.")

# driver.close()
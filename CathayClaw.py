import os
import re
import cv2
import time
import pandas
import shutil
import datetime
import pytesseract
import selenium.webdriver as WebDriver
import selenium.webdriver.common.by as WebBy
import selenium.webdriver.support.ui as WebUI
import selenium.webdriver.common.action_chains as WebAction
import selenium.webdriver.support.expected_conditions as WebConditions

pytesseract.pytesseract.tesseract_cmd = "C:\\Program Files\\Tesseract-OCR\\tesseract.exe"
ChromeDriver = WebDriver.Chrome("chromedriver_Cathay.exe")
ChromeDriver.maximize_window()
print("程式開始 ", datetime.datetime.now())
# 登入金控
if True:
    ChromeDriver.get("https://w3.cathaylife.com.tw")
    while True:
        ChromeDriver.find_element(WebBy.By.ID, "UID").send_keys("A120923934")
        ChromeDriver.find_element(WebBy.By.ID, "KEY").send_keys("0938196099")
        ChromeDriver.find_element(WebBy.By.ID, "btnLogin").click()
        time.sleep(0.1)
        if len(ChromeDriver.find_elements(WebBy.By.ID, "UID")) == 0:
            break
# 登入金控
# 進保戶關係管理系統
if True:
    time.sleep(0.1)
    ChromeDriver.execute_script(ChromeDriver.find_element(WebBy.By.XPATH, "//*[@id='idx_menu']/ul[2]/li[4]/ul/li[3]/ul/li[2]/a").get_attribute("onclick"))
    time.sleep(10)
    ChromeDriver.switch_to.window(ChromeDriver.window_handles[1])
    ChromeDriver.close()
    ChromeDriver.switch_to.window(ChromeDriver.window_handles[1])
# 進保戶關係管理系統
# 重讀保戶總表
if 1 == 1:
    CustomerTable = pandas.read_csv("TableCustomer.csv", index_col="ID", encoding="utf_8_sig", quoting=1, dtype=str).fillna("")
    CustomerTable = CustomerTable[CustomerTable.index != CustomerTable.index]
    ChromeDriver.switch_to.default_content()
    ChromeDriver.switch_to.frame(ChromeDriver.find_element(WebBy.By.XPATH, "/html/frameset/frame[2]"))
    ChromeDriver.find_element(WebBy.By.XPATH, "/html/body/table[1]/tbody/tr[2]/td[1]/input").click()
    time.sleep(0.1)
    ChromeDriver.switch_to.default_content()
    ChromeDriver.switch_to.frame(ChromeDriver.find_element(WebBy.By.XPATH, "/html/frameset/frame[2]"))
    ChromeDriver.switch_to.frame(ChromeDriver.find_element(WebBy.By.XPATH, "/html/frameset/frame[2]"))
    for CustomerRowTag in ChromeDriver.find_elements(WebBy.By.XPATH, "/html/body/form/table[1]/tbody/tr"):
        if len(CustomerRowTag.find_elements(WebBy.By.XPATH, "td[2]/div/a")) == 0:
            continue
        CustomerID = CustomerRowTag.find_element(WebBy.By.XPATH, "td[2]/div/a/span").text
        CustomerName = CustomerRowTag.find_element(WebBy.By.XPATH, "td[3]/div").text.replace("*", "").replace("#", "").replace("$", "")
        CustomerTel = CustomerRowTag.find_element(WebBy.By.XPATH, "td[6]/div").text.replace("(", "").replace(")", "")
        CustomerMob = CustomerRowTag.find_element(WebBy.By.XPATH, "td[7]/div").text.replace("(", "").replace(")", "")
        if CustomerTel != "" and CustomerMob != "":
            CustomerTel = CustomerTel + ";" + CustomerMob
        elif CustomerMob != "":
            CustomerTel = CustomerMob
        CustomerBirth = [int(Date) for Date in CustomerRowTag.find_element(WebBy.By.XPATH, "td[9]/div").text.split("/")]
        CustomerBirth = CustomerBirth[0] * 10000 + CustomerBirth[1] * 100 + CustomerBirth[2] + 19110000
        CustomerTable.loc[CustomerID] = [CustomerName, CustomerBirth, CustomerTel, "", ""]
    CustomerTable.sort_index(inplace=True)
    CustomerTable.to_csv("TableCustomer.csv", encoding="utf_8_sig", quoting=1)
# 重讀保戶總表
print("重讀保戶總表 ", datetime.datetime.now())
# 讀取保戶情報
if 1 == 1:
    NumberMapDict = {Key + 14: Value for Key, Value in {32: " ", 40: "(", 41: ")", 45: "-", 48: "0", 49: "1", 50: "2", 51: "3", 52: "4", 53: "5", 54: "6", 55: "7", 56: "8", 57: "9", 70: "F"}.items()}
    CustomerTable = pandas.read_csv("TableCustomer.csv", index_col="ID", encoding="utf_8_sig", quoting=1, dtype=str).fillna("")
    for CustomerID in CustomerTable.index:
        [CustomerName, CustomerBirth, CustomerTel, CustomerAddress, CustomerFamilyHead] = CustomerTable.loc[CustomerID]
        CustomerTel = [] if CustomerTel == "" else CustomerTel.split(";")
        CustomerAddress = [] if CustomerAddress == "" else CustomerAddress.split(";")
        # 回 CRM 首頁
        ChromeDriver.refresh()
        time.sleep(0.1)
        # 回 CRM 首頁
        # 指定客戶 ID
        ChromeDriver.switch_to.default_content()
        ChromeDriver.switch_to.frame(ChromeDriver.find_element(WebBy.By.XPATH, "/html/frameset/frame[2]"))
        ChromeDriver.find_element(WebBy.By.XPATH, "/html/body/table[1]/tbody/tr[4]/td[2]/input[1]").send_keys(CustomerID)
        time.sleep(0.1)
        ChromeDriver.find_element(WebBy.By.XPATH, "/html/body/table[1]/tbody/tr[4]/td[2]/input[2]").click()
        time.sleep(0.1)
        ChromeDriver.switch_to.default_content()
        ChromeDriver.switch_to.frame(ChromeDriver.find_element(WebBy.By.XPATH, "/html/frameset/frame[2]"))
        ChromeDriver.switch_to.frame(ChromeDriver.find_element(WebBy.By.XPATH, "/html/frameset/frame[2]"))
        ChromeDriver.find_element(WebBy.By.XPATH, "/html/body/form/table[1]/tbody/tr[1]/td[2]/div/a").click()
        time.sleep(0.1)
        # 指定客戶 ID
        # 讀取客戶資料
        ChromeDriver.switch_to.default_content()
        ChromeDriver.switch_to.frame(ChromeDriver.find_element(WebBy.By.XPATH, "/html/frameset/frame[2]"))
        ChromeDriver.switch_to.frame(ChromeDriver.find_element(WebBy.By.XPATH, "/html/frameset/frame[2]"))
        ChromeDriver.save_screenshot("ChromeScreen.png")
        CellTagList = ChromeDriver.find_elements(WebBy.By.XPATH, "//*[@id='TB_inside']/tbody/tr/td")
        for IndexTag, CellTag in enumerate(CellTagList):
            CellText = CellTag.text
            if CellText.find("\n") >= 0:
                CellText = CellText[: CellText.find("\n")]
            if re.match(".*地址.*", CellText):
                GoogleMapTagList = CellTagList[IndexTag + 1].find_elements(WebBy.By.XPATH, "a")
                if len(GoogleMapTagList) == 1:
                    Address = GoogleMapTagList[0].get_attribute("onclick")
                    Address = Address[Address.find("maps/search/") + 12: Address.find("','popup'")]
                    CustomerAddress.append(Address)
                elif len(GoogleMapTagList) > 0:
                    raise Exception()
            elif re.match(".*(電話|手機).*", CellText) and not re.match(".*(電話行銷).*", CellText):
                Telephone = ""
                for ImageTag in CellTagList[IndexTag + 1].find_elements(WebBy.By.XPATH, "img"):
                    NumberRGB = cv2.imread("ChromeScreen.png")
                    NumberRGB = NumberRGB[ImageTag.location_once_scrolled_into_view["y"]: ImageTag.location_once_scrolled_into_view["y"] + ImageTag.size["height"], ImageTag.location_once_scrolled_into_view["x"]: ImageTag.location_once_scrolled_into_view["x"] + ImageTag.size["width"]]
                    cv2.imwrite("Telephone.png", NumberRGB)
                    Temp = pytesseract.image_to_string(NumberRGB, config="-psm 7 -c tessedit_char_whitelist=-()0123456789")
                    NumberList = ImageTag.get_attribute("src").split("|")
                    NumberList.pop(len(NumberList) - 1)
                    NumberList[0] = NumberList[0][len(NumberList[0]) - 2:]
                    Telephone += "".join([NumberMapDict[int(Number)] for Number in NumberList])
                if Telephone != "":
                    CustomerTel.append(Telephone)
        # 讀取客戶資料
        CustomerTel = ";".join(set(CustomerTel))
        CustomerAddress = ";".join(set(CustomerAddress))
        CustomerTable.loc[CustomerID] = [CustomerName, CustomerBirth, CustomerTel, CustomerAddress, CustomerFamilyHead]
        CustomerTable.sort_index(inplace=True)
        CustomerTable.to_csv("TableCustomer.csv", encoding="utf_8_sig", quoting=1)
# 讀取保戶情報
print("讀取保戶情報 ", datetime.datetime.now())
# 讀取保單健檢
if 1 == 1:
    PolicyTable = pandas.read_csv("TablePolicy.csv", encoding="utf_8_sig", quoting=1, dtype=str).fillna("")
    PolicyTable = PolicyTable[PolicyTable.index != PolicyTable.index]
    CustomerTable = pandas.read_csv("TableCustomer.csv", index_col="ID", encoding="utf_8_sig", quoting=1, dtype=str).fillna("")
    IndexFamily = 0
    ChromeDriver.switch_to.default_content()
    ChromeDriver.switch_to.frame(ChromeDriver.find_element(WebBy.By.XPATH, "/html/frameset/frame[1]"))
    ChromeDriver.find_element(WebBy.By.XPATH, "/html/body/table[1]/tbody/tr[2]/td[5]/a").click()
    try:
        WebUI.WebDriverWait(ChromeDriver, 10).until(WebConditions.alert_is_present())
        ChromeDriver.switch_to.alert.accept()
    except:
        ChromeDriver.switch_to.default_content()
    time.sleep(0.1)
    while True:
        # 列出所有家庭
        ChromeDriver.switch_to.default_content()
        ChromeDriver.switch_to.frame(ChromeDriver.find_element(WebBy.By.XPATH, "/html/frameset/frame[2]"))
        ChromeDriver.find_element(WebBy.By.XPATH, "//*[@id='aspnetForm']/table/tbody/tr/td[2]/table/tbody/tr[1]/td[2]/input").click()
        time.sleep(0.1)
        # 列出所有家庭
        # 點選最新家庭
        ChromeDriver.switch_to.default_content()
        ChromeDriver.switch_to.frame(ChromeDriver.find_element(WebBy.By.XPATH, "/html/frameset/frame[2]"))
        FamilyTagList = ChromeDriver.find_elements(WebBy.By.XPATH, "//*[@id='ctl00_HeaderContentHolder_Table1']/tbody/tr/td/a")
        if IndexFamily >= len(FamilyTagList):
            break
        FamilyTagList[IndexFamily].click()
        WebUI.WebDriverWait(ChromeDriver, 10).until(WebConditions.alert_is_present())
        ChromeDriver.switch_to.alert.accept()
        time.sleep(0.1)
        # 點選最新家庭
        # 讀取家庭成員資料
        CustomerRowList = ChromeDriver.find_elements(WebBy.By.XPATH, "//*[@id='ctl00_ctl00_HeaderContentHolder_PageContentHolder_FamilyMembersViewControl1_FamilyMembersGrid']/tbody/tr")
        CustomerRowList.pop(0)
        FamilyHead = CustomerRowList[0].find_element(WebBy.By.XPATH, "td[3]/img").get_attribute("src")
        FamilyHead = FamilyHead[FamilyHead.find("Key=") + 4:]
        FamilyHead = FamilyHead[: 10]
        for CustomerRow in CustomerRowList:
            CustomerID = CustomerRow.find_element(WebBy.By.XPATH, "td[3]/img").get_attribute("src")
            CustomerID = CustomerID[CustomerID.find("Key=") + 4:]
            CustomerID = CustomerID[: 10]
            [CustomerName, CustomerBirth, CustomerTel, CustomerAddress] = [[], [], [], []]
            if CustomerID in CustomerTable.index:
                [CustomerName, CustomerBirth, CustomerTel, CustomerAddress, CustomerFamilyHead] = [[] if Value == "" else Value.split(";") for Value in CustomerTable.loc[CustomerID]]
            CustomerName.append(CustomerRow.find_element(WebBy.By.XPATH, "td[2]").text)
            CustomerBirthDate = [int(Date) for Date in CustomerRow.find_element(WebBy.By.XPATH, "td[5]").text.split("/")]
            CustomerBirth.append(str(CustomerBirthDate[0] * 10000 + CustomerBirthDate[1] * 100 + CustomerBirthDate[2] + 19110000))
            CustomerTable.loc[CustomerID] = [";".join(set(Value)) for Value in [CustomerName, CustomerBirth, CustomerTel, CustomerAddress, {FamilyHead}]]
        # 讀取家庭成員資料
        # 點選契約內容
        ChromeDriver.switch_to.default_content()
        ChromeDriver.switch_to.frame(ChromeDriver.find_element(WebBy.By.XPATH, "/html/frameset/frame[2]"))
        ChromeDriver.find_element(WebBy.By.XPATH, "//*[@id='ctl00_ctl00_HeaderContentHolder_Menu']/table[2]/tbody/tr/td[2]/a").click()
        time.sleep(0.1)
        # 點選契約內容
        # 讀取契約內容
        ChromeDriver.switch_to.default_content()
        ChromeDriver.switch_to.frame(ChromeDriver.find_element(WebBy.By.XPATH, "/html/frameset/frame[2]"))
        PolicyRowList = ChromeDriver.find_elements(WebBy.By.XPATH, "//*[@id='aspnetForm']/table[2]/tbody/tr/td[2]/table[1]/tbody/tr[2]/td/table/tbody/tr")
        if len(PolicyRowList) > 0:
            LoadPolicyList = PolicyTable[PolicyTable.index != PolicyTable.index]
            PolicyRowList.pop(0)
            InsuredID = ""
            for PolicyRow in PolicyRowList:
                if PolicyRow.text.startswith("被保險人"):
                    InsuredID = PolicyRow.find_element(WebBy.By.XPATH, "td/img").get_attribute("src")
                    InsuredID = InsuredID[InsuredID.find("Key=") + 4:]
                    InsuredID = InsuredID[: 10]
                    continue
                PolicyNumber = PolicyRow.find_element(WebBy.By.XPATH, "td[1]/img").get_attribute("src")
                PolicyNumber = PolicyNumber[PolicyNumber.find("Key=") + 4:]
                PolicyNumber = PolicyNumber[: 10]
                ProductName = PolicyRow.find_element(WebBy.By.XPATH, "td[2]").text
                PayTerm = PolicyRow.find_element(WebBy.By.XPATH, "td[3]").text
                Relation = PolicyRow.find_element(WebBy.By.XPATH, "td[4]").text
                EffectiveDate = [int(Date) for Date in PolicyRow.find_element(WebBy.By.XPATH, "td[5]").text.split("/")]
                EffectiveDate = EffectiveDate[0] * 10000 + EffectiveDate[1] * 100 + EffectiveDate[2] + 19110000
                SumAssured = PolicyRow.find_element(WebBy.By.XPATH, "td[6]").text.replace("\n", ";")
                PayPremium = PolicyRow.find_element(WebBy.By.XPATH, "td[7]").text
                PayType = PolicyRow.find_element(WebBy.By.XPATH, "td[8]").text
                Beneficiary = PolicyRow.find_element(WebBy.By.XPATH, "td[9]").text.replace("\n", ";")
                PayRoute = ""
                OwnerName = ""
                if ProductName.startswith("**要保人"):
                    IndexString = ProductName.find("\n")
                    OwnerName = ProductName[6: IndexString]
                    ProductName = ProductName[IndexString + 1:].replace("\n", ";")
                    IndexString = PayPremium.find("元\n")
                    if IndexString < 0:
                        PayRoute = PayPremium.replace("\n", "")
                        PayPremium = ""
                    else:
                        PayRoute = PayPremium[IndexString + 2:] .replace("\n", "")
                        PayPremium = PayPremium[:IndexString + 1]
                else:
                    ProductName = ProductName[4:].replace("\n    ", ";")
                LoadPolicyList = LoadPolicyList.append(pandas.Series(["國泰人壽", InsuredID,  PolicyNumber, OwnerName,  ProductName,  PayTerm,  Relation,  EffectiveDate,  SumAssured, PayPremium, PayType,  PayRoute,  Beneficiary], index=LoadPolicyList.columns), ignore_index=True)
            for [Company, InsuredID, PolicyNumber] in LoadPolicyList.groupby(["Company", "InsuredID", "PolicyNumber"]).indices:
                PolicyTable = PolicyTable[~((PolicyTable.Company == Company) & (PolicyTable.InsuredID == InsuredID) & (PolicyTable.PolicyNumber == PolicyNumber))]
            PolicyTable = PolicyTable.append(LoadPolicyList, ignore_index=True)
        # 讀取契約內容
        IndexFamily += 1
        PolicyTable.sort_values(by=["InsuredID", "Company", "PolicyNumber"], inplace=True)
        PolicyTable.to_csv("TablePolicy.csv", index=False, encoding="utf_8_sig", quoting=1)
        CustomerTable.sort_index(inplace=True)
        CustomerTable.to_csv("TableCustomer.csv", encoding="utf_8_sig", quoting=1)
# 讀取保單健檢
print("讀取保單健檢 ", datetime.datetime.now())
ChromeDriver.quit()
print("程式結束 ", datetime.datetime.now())

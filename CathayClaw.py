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
DownloadPath = os.path.abspath("..\\Box Sync\\CathayClawData")
if not os.path.exists(DownloadPath):
    os.makedirs(DownloadPath)
for FileName in os.listdir(DownloadPath):
    if FileName.lower()[len(FileName) - 4:] not in [".csv"]:
        os.unlink(DownloadPath + "\\" + FileName)
NumberMapDict = {46: " ", 54: "(", 55: ")", 59: "-", 62: "0", 63: "1", 64: "2", 65: "3", 66: "4", 67: "5", 68: "6", 69: "7", 70: "8", 71: "9", 84: "F"}

StartTime = datetime.datetime.now()

CustomerTable = pandas.read_csv(DownloadPath + "\\Customer.csv", encoding="utf_8_sig", quoting=1, dtype=str)
CustomerTable.set_index("ID", inplace=True, verify_integrity=True)
CustomerTable.fillna("", inplace=True)

ChromeDriver = WebDriver.Chrome("chromedriver.exe")
ChromeDriver.maximize_window()
ChromeDriver.implicitly_wait(10)

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
    time.sleep(5)
    ChromeDriver.execute_script(ChromeDriver.find_element(WebBy.By.XPATH, "//*[@id='idx_menu']/ul[2]/li[4]/ul/li[3]/ul/li[2]/a").get_attribute("onclick"))
    time.sleep(5)
    ChromeDriver.switch_to.window(ChromeDriver.window_handles[1])
    ChromeDriver.close()
    ChromeDriver.switch_to.window(ChromeDriver.window_handles[1])
# 進保戶關係管理系統
# 列出所有保戶
if 1 == 0:
    ChromeDriver.switch_to.default_content()
    ChromeDriver.switch_to.frame(ChromeDriver.find_element(WebBy.By.XPATH, "/html/frameset/frame[2]"))
    ChromeDriver.find_element(WebBy.By.XPATH, "/html/body/table[1]/tbody/tr[2]/td[1]/input").click()
    time.sleep(1)
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
        CustomerTable.loc[CustomerID] = [CustomerName, CustomerBirth, CustomerTel, ""]
    CustomerTable.to_csv(DownloadPath + "\\Customer.csv", encoding="utf_8_sig", quoting=1)
# 列出所有保戶
# 讀取保戶情報
if 1 == 0:
    for CustomerID in CustomerTable.index:
        [CustomerName, CustomerBirth, CustomerTel, CustomerAddress] = CustomerTable.loc[CustomerID]
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
        ChromeDriver.save_screenshot(DownloadPath + "\\ChromeScreen.png")
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
                    NumberRGB = cv2.imread(DownloadPath + "\\ChromeScreen.png")
                    NumberRGB = NumberRGB[ImageTag.location_once_scrolled_into_view["y"]: ImageTag.location_once_scrolled_into_view["y"] + ImageTag.size["height"], ImageTag.location_once_scrolled_into_view["x"]: ImageTag.location_once_scrolled_into_view["x"] + ImageTag.size["width"]]
                    cv2.imwrite(DownloadPath + "\\Telephone.png", NumberRGB)
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
        CustomerTable.loc[CustomerID] = [CustomerName, CustomerBirth, CustomerTel, CustomerAddress]
        CustomerTable.to_csv(DownloadPath + "\\Customer.csv", encoding="utf_8_sig", quoting=1)
# 讀取保戶情報
# 讀取保單健檢
if 1 == 1:
    IndexFamily = 0
    ChromeDriver.switch_to.default_content()
    ChromeDriver.switch_to.frame(ChromeDriver.find_element(WebBy.By.XPATH, "/html/frameset/frame[1]"))
    ChromeDriver.find_element(WebBy.By.XPATH, "/html/body/table[1]/tbody/tr[2]/td[5]/a").click()
    time.sleep(1)
    while True:
        ChromeDriver.switch_to.default_content()
        ChromeDriver.switch_to.frame(ChromeDriver.find_element(WebBy.By.XPATH, "/html/frameset/frame[2]"))
        ChromeDriver.find_element(WebBy.By.XPATH, "//*[@id='aspnetForm']/table/tbody/tr/td[2]/table/tbody/tr[1]/td[2]/input").click()
        time.sleep(1)
        ChromeDriver.switch_to.default_content()
        ChromeDriver.switch_to.frame(ChromeDriver.find_element(WebBy.By.XPATH, "/html/frameset/frame[2]"))
        FamilyTagList = ChromeDriver.find_elements(WebBy.By.XPATH, "//*[@id='ctl00_HeaderContentHolder_Table1']/tbody/tr/td/a")
        if IndexFamily >= len(FamilyTagList):
            break
        FamilyTagList[IndexFamily].click()
        WebUI.WebDriverWait(ChromeDriver, 10).until(WebConditions.alert_is_present())
        ChromeDriver.switch_to.alert.accept()
        time.sleep(1)
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
                [CustomerName, CustomerBirth, CustomerTel, CustomerAddress] = [[] if Value == "" else Value.split(";") for Value in CustomerTable.loc[CustomerID]]
            CustomerName.append(CustomerRow.find_element(WebBy.By.XPATH, "td[2]").text)
            CustomerBirthDate = [int(Date) for Date in CustomerRow.find_element(WebBy.By.XPATH, "td[5]").text.split("/")]
            CustomerBirth.append(str(CustomerBirthDate[0] * 10000 + CustomerBirthDate[1] * 100 + CustomerBirthDate[2] + 19110000))
            CustomerTable.loc[CustomerID] = [";".join(set(Value)) for Value in [CustomerName, CustomerBirth, CustomerTel, CustomerAddress]]
        IndexFamily += 1
        CustomerTable.to_csv(DownloadPath + "\\Customer.csv", encoding="utf_8_sig", quoting=1)
# 讀取保單健檢
ChromeDriver.quit()

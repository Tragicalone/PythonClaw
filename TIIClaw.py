import os
import re
import cv2
import time
import numpy
import pandas
import shutil
import datetime
import pytesseract
import selenium.webdriver as WebDriver
import selenium.webdriver.common.by as WebBy
import selenium.webdriver.support.ui as WebUI
import selenium.webdriver.common.keys as WebKey
import selenium.webdriver.remote.webelement as WebElement
import selenium.webdriver.support.expected_conditions as WebCondition


pytesseract.pytesseract.tesseract_cmd = "C:\\Program Files\\Tesseract-OCR\\tesseract.exe"
DownloadPath = os.path.abspath("..\\Box Sync\\TIIClawData")
if not os.path.exists(DownloadPath):
    os.makedirs(DownloadPath)
for FileName in os.listdir(DownloadPath):
    if re.match(".*.pdf|.*.png",  FileName.lower()):
        os.unlink(DownloadPath + "\\" + FileName)

StartTime = datetime.datetime.now()
ChromeOptions = WebDriver.ChromeOptions()
ChromeOptions.add_experimental_option("prefs", {"plugins.plugins_list": [{"enabled": False, "name": "Chrome PDF Viewer"}], "download.default_directory": DownloadPath})

FileError = open("ErrorTIIClaw.log", mode="w", encoding="UTF-8")
FileError.close()
DownloadTable = pandas.read_csv("TableDLHistory.csv", encoding="utf_8_sig", quoting=1, dtype=numpy.str)
DownloadTable.set_index("TIISerial", inplace=True, verify_integrity=True)

ChromeDriver = WebDriver.Chrome("chromedriver_TII.exe", options=ChromeOptions)
ChromeDriver.maximize_window()
ChromeDriver.get("http://insprod.tii.org.tw/database/insurance/query.asp")

CompanySelect = WebUI.Select(ChromeDriver.find_element(WebBy.By.XPATH, "//*[@id='printContext']/table/tbody/tr/td/table[2]/tbody/tr/td[1]/table[3]/tbody/tr/td/table/tbody/tr[1]/td/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr[2]/td/table/tbody/tr/td[2]/select"))
CompanyList = [option.text.replace("-", "_") for option in CompanySelect.options]

for IndexCompany, CompanyName in enumerate(CompanyList):
    # 篩選壽險 Begin
    if int(CompanyName[0:3]) < 200:
        continue
    # 篩選壽險 End
    if not os.path.exists(DownloadPath + "\\" + CompanyName):
        os.makedirs(DownloadPath + "\\" + CompanyName)
    ChromeDriver.get("chrome://settings/clearBrowserData")
    time.sleep(3)
    ClearButton = ChromeDriver.execute_script("return arguments[0].shadowRoot", ChromeDriver.find_element(WebBy.By.XPATH, "/html/body/settings-ui"))
    ClearButton = ChromeDriver.execute_script("return arguments[0].shadowRoot", ClearButton.find_element(WebBy.By.ID, "main"))
    ClearButton = ChromeDriver.execute_script("return arguments[0].shadowRoot", ClearButton.find_element(WebBy.By.CSS_SELECTOR, "settings-basic-page"))
    ClearButton = ChromeDriver.execute_script("return arguments[0].shadowRoot", ClearButton.find_element(WebBy.By.CSS_SELECTOR, "settings-privacy-page"))
    ClearButton = ChromeDriver.execute_script("return arguments[0].shadowRoot", ClearButton.find_element(WebBy.By.CSS_SELECTOR, "settings-clear-browsing-data-dialog"))
    ClearButton = ClearButton.find_element(WebBy.By.ID, "clearBrowsingDataConfirm")
    ClearButton.click()
    time.sleep(15)
    IndexPage = 1
    while IndexPage > 0:
        # 破驗證碼 Begin
        while True:
            ChromeDriver.get("http://insprod.tii.org.tw/database/insurance/query.asp")
            CompanySelect = WebUI.Select(ChromeDriver.find_element(WebBy.By.XPATH, "//*[@id='printContext']/table/tbody/tr/td/table[2]/tbody/tr/td[1]/table[3]/tbody/tr/td/table/tbody/tr[1]/td/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr[2]/td/table/tbody/tr/td[2]/select"))
            CompanySelect.select_by_index(IndexCompany)
            # 只下載現售商品 Begin
            if False:
                CurrentTag = ChromeDriver.find_element(WebBy.By.XPATH, "//*[@id='printContext']/table/tbody/tr/td/table[2]/tbody/tr/td[1]/table[3]/tbody/tr/td/table/tbody/tr[1]/td/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr[3]/td/table/tbody/tr[3]/td[2]/label[2]/input[1]")
                while not bool(CurrentTag.get_attribute("checked")):
                    CurrentTag.click()
                    time.sleep(0.1)
            # 只下載現售商品 End
            ImageTag = ChromeDriver.find_element(WebBy.By.XPATH, "//*[@id='printContext']/table/tbody/tr/td/table[2]/tbody/tr/td[1]/table[3]/tbody/tr/td/table/tbody/tr[1]/td/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr[5]/td/img")
            ChromeDriver.save_screenshot("ChromeScreen.png")
            RGBOrigin = cv2.imread("ChromeScreen.png")
            RGBOrigin = RGBOrigin[ImageTag.location["y"] + 1: ImageTag.location["y"] + ImageTag.size["height"] - 1, ImageTag.location["x"] + 1: ImageTag.location["x"] + ImageTag.size["width"] - 1]
            RGBBlurred = cv2.cvtColor(RGBOrigin, cv2.IMREAD_GRAYSCALE)
            RGBBlurred = cv2.medianBlur(RGBBlurred, 5)
            RGBBlurred = cv2.threshold(RGBBlurred, 127, 255, cv2.THRESH_BINARY)[1]
            ValidateText = ""
            for IndexH in range(4):
                RGBChar = RGBBlurred[0: RGBBlurred.shape[0], int(IndexH * RGBBlurred.shape[1] / 4): int((IndexH + 1) * RGBBlurred.shape[1] / 4)]
                ValidateText += pytesseract.image_to_string(RGBChar, config="-psm 10 -c tessedit_char_whitelist=0123456789")
            ChromeDriver.find_element(WebBy.By.XPATH, "//*[@id='printContext']/table/tbody/tr/td/table[2]/tbody/tr/td[1]/table[3]/tbody/tr/td/table/tbody/tr[1]/td/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr[5]/td/input[1]").send_keys(ValidateText)
            time.sleep(1)
            ChromeDriver.find_element(WebBy.By.ID, "Go2225").click()
            try:
                WebUI.WebDriverWait(ChromeDriver, 3).until(WebCondition.alert_is_present())
                ChromeDriver.switch_to.alert.accept()
            except Exception as exception:
                break
            time.sleep(0.1)
        # 破驗證碼 End
        # 下載 Begin
        while True:
            ChromeDriver.get("http://insprod.tii.org.tw/database/insurance/resultQueryAll.asp?page=" + str(IndexPage) + "&fQueryAll=&CompanyID=" + CompanyName[:3] + "&categoryId=")
            # 驗證碼超時處理 Begin
            try:
                WebUI.WebDriverWait(ChromeDriver, 3).until(WebCondition.alert_is_present())
                ChromeDriver.switch_to.alert.accept()
                break
            except Exception as exception:
                time.sleep(0.1)
            # 驗證碼超時處理 End
            ProductNameURLList = []
            # 下載當頁所有商品 URL Begin
            for ProductTag in ChromeDriver.find_elements(WebBy.By.XPATH, "//*[@id='printContext']/table/tbody/tr/td/table[2]/tbody/tr/td/table[3]/tbody/tr/td/table/tbody/tr[3]/td/table[2]/tbody/tr"):
                if ProductTag.size["height"] < 2:
                    continue
                ProductName = ProductTag.find_element(WebBy.By.XPATH, "td[2]").text.lstrip().replace("/", "／").replace("?", "？").replace("\\", "＼").replace("*", "＊").replace("#", "＃").replace("<", "＜").replace(">", "＞").replace("|", "｜").replace("\"", "”").replace(":", "：")
                LaunchDate = ProductTag.find_element(WebBy.By.XPATH, "td[4]").text.strip().replace("/", "")
                CloseDate = ProductTag.find_element(WebBy.By.XPATH, "td[6]").text.strip().replace("/", "")
                ProductURL = ProductTag.find_element(WebBy.By.XPATH, "td[2]/a").get_attribute("href")
                TIISerial = ProductURL[ProductURL.find("productId=") + 10:]
                ProductNameURLList.append([ProductName, LaunchDate, CloseDate, TIISerial, ProductURL])
            # 下載當頁所有商品 URL End
            if len(ProductNameURLList) == 0:
                IndexPage = -1
                break
            # 下載當頁所有商品 Begin
            for [ProductName, LaunchDate, CloseDate, TIISerial, ProductURL] in ProductNameURLList:
                # 該 TIISerial 已下載 Begin
                if TIISerial in DownloadTable.index:
                    Company = DownloadTable.loc[TIISerial]
                    os.rename(DownloadPath + "\\" + CompanyName + "\\" + TIISerial + "_" + Company.values[1][:20] + "_" + Company.values[2] + "_" + Company.values[3] + ".pdf", DownloadPath + "\\" + CompanyName + "\\" + TIISerial + "_" + ProductName[:20] + "_" + LaunchDate + "_" + CloseDate + ".pdf")
                    DownloadTable.loc[TIISerial] = [CompanyName, ProductName, LaunchDate, CloseDate]
                    continue
                # 該 TIISerial 已下載 End
                ChromeDriver.get(ProductURL)
                # 驗證碼超時處理 Begin
                try:
                    WebUI.WebDriverWait(ChromeDriver, 3).until(WebCondition.alert_is_present())
                    ChromeDriver.switch_to.alert.accept()
                    break
                except Exception as exception:
                    time.sleep(0.1)
                # 驗證碼超時處理 End
                ProvisionTag = None
                # 下載處理 Begin
                try:
                    ProvisionTag = ChromeDriver.find_element(WebBy.By.XPATH, "//*[@id='printContext']/table/tbody/tr/td/table[2]/tbody/tr/td/table[3]/tbody/tr/td/table[1]/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr[19]/td/table/tbody/tr/td/a")
                except Exception as exception:
                    with open("ErrorTIIClaw.log", mode="a", encoding="UTF-8") as FileError:
                        FileError.write("無條款 " + CompanyName + "\\" + ProductName + "\n")
                if ProvisionTag:
                    ProvisionFileName = ProvisionTag.text
                    ProvisionTag.click()
                    if len(ChromeDriver.find_elements(WebBy.By.ID, "printContext")) == 0:
                        with open("ErrorTIIClaw.log", mode="a", encoding="UTF-8") as FileError:
                            FileError.write("無連結 " + CompanyName + "\\" + ProductName + "\n")
                        continue
                    ProvisionFileExists = False
                    for IndexWait in range(120):
                        ProvisionFileExists = os.path.isfile(DownloadPath + "\\" + ProvisionFileName)
                        time.sleep(1)
                        if ProvisionFileExists:
                            shutil.move(DownloadPath + "\\" + ProvisionFileName, DownloadPath + "\\" + CompanyName + "\\" + TIISerial + "_" + ProductName[:20] + "_" + LaunchDate + "_" + CloseDate + ".pdf")
                            DownloadTable.loc[TIISerial] = [CompanyName, ProductName, LaunchDate, CloseDate]
                            break
                    if not ProvisionFileExists:
                        with open("ErrorTIIClaw.log", mode="a", encoding="UTF-8") as FileError:
                            FileError.write("下載失敗 " + CompanyName + "\\" + ProductName + ".pdf\n")
                # 下載處理 End
                time.sleep(0.1)
            # 下載當頁所有商品 End
            DownloadTable.sort_index(inplace=True)
            DownloadTable.to_csv("TableDLHistory.csv", encoding="utf_8_sig", quoting=1)
            IndexPage += 1
        # 下載 End
    print("已下蛓 ", CompanyName, " 目前時間", datetime.datetime.now(), " 共使用 ", (datetime.datetime.now() - StartTime).total_seconds(), " 秒")
ChromeDriver.close()
print("共使用 ", (datetime.datetime.now() - StartTime).total_seconds(), " 秒")

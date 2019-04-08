import os
import re
import time
import pandas
import shutil
import datetime
import selenium.webdriver as WebDriver
import selenium.webdriver.common.by as WebBy
import selenium.webdriver.support.ui as WebUI
import selenium.webdriver.common.action_chains as WebAction
import selenium.webdriver.support.expected_conditions as WebConditions

ChromeDriver = WebDriver.Chrome("chromedriver_Cathay.exe")
ChromeDriver.maximize_window()
print("程式開始 ", datetime.datetime.now())
# 登入金控 Begin
ChromeDriver.get("https://w3.cathaylife.com.tw")
ChromeDriver.find_element(WebBy.By.ID, "UID").send_keys("A120923934")
ChromeDriver.find_element(WebBy.By.ID, "KEY").send_keys("0938196099")
ChromeDriver.find_element(WebBy.By.ID, "btnLogin").click()
time.sleep(0.1)
# 登入金控 End
# 重讀保戶總表情報 Begin
if 1 == 1:
    # 進保戶關係管理系統 Begin
    ChromeDriver.execute_script(ChromeDriver.find_element(WebBy.By.XPATH, "//*[@id='idx_menu']/ul[2]/li[4]/ul/li[3]/ul/li[2]/a").get_attribute("onclick"))
    time.sleep(10)
    if len(ChromeDriver.window_handles) == 3:
        ChromeDriver.switch_to.window(ChromeDriver.window_handles[1])
        ChromeDriver.close()
    elif len(ChromeDriver.window_handles) > 3:
        raise Exception()
    ChromeDriver.switch_to.window(ChromeDriver.window_handles[1])
    # 進保戶關係管理系統 End
    CustomerTable = pandas.read_csv("TableCustomer.csv", index_col="ID", encoding="utf_8_sig", quoting=1, dtype=str).fillna("")
    CustomerTable = CustomerTable[CustomerTable.index != CustomerTable.index]
    # 所有客戶列表 Begin
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
        CustomerTable.loc[CustomerID, ["Name", "Birthday", "Telephone"]] = [CustomerName, CustomerBirth, CustomerTel]
    CustomerTable.sort_index(inplace=True)
    CustomerTable.to_csv("TableCustomer.csv", encoding="utf_8_sig", quoting=1)
    # 所有客戶列表 End
    # 逐一客戶情報 Begin
    # Tue 12 Wed 14 Thu 16 Mon 7 Tue 9 Mon 14 Tue 16 Wed 8
    NumberMapDict = {Key + 8: Value for Key, Value in {32: " ", 40: "(", 41: ")", 45: "-", 48: "0", 49: "1", 50: "2", 51: "3", 52: "4", 53: "5", 54: "6", 55: "7", 56: "8", 57: "9", 70: "F"}.items()}
    for CustomerID in CustomerTable.index:
        [CustomerName, CustomerBirth, CustomerTel] = CustomerTable.loc[CustomerID, ["Name", "Birthday", "Telephone"]]
        CustomerTel = [] if CustomerTel == "" else CustomerTel.split(";")
        CustomerAddress = []
        # 回 CRM 首頁 Begin
        ChromeDriver.refresh()
        time.sleep(0.1)
        # 回 CRM 首頁 End
        # 指定客戶 ID Begin
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
        # 指定客戶 ID End
        # 讀取客戶資料 Begin
        ChromeDriver.switch_to.default_content()
        ChromeDriver.switch_to.frame(ChromeDriver.find_element(WebBy.By.XPATH, "/html/frameset/frame[2]"))
        ChromeDriver.switch_to.frame(ChromeDriver.find_element(WebBy.By.XPATH, "/html/frameset/frame[2]"))
        CellTagList = ChromeDriver.find_elements(WebBy.By.XPATH, "//*[@id='TB_inside']/tbody/tr/td")
        CellTagList += ChromeDriver.find_elements(WebBy.By.XPATH, "//*[@id='TB_o_customer']/tbody/tr/td")
        for IndexTag, CellTag in enumerate(CellTagList):
            CellText = CellTag.text
            if CellText.find("\n") >= 0:
                CellText = CellText[: CellText.find("\n")]
            if re.match(".*地址.*", CellText):
                GoogleMapTagList = CellTagList[IndexTag + 1].find_elements(WebBy.By.XPATH, "a")
                if len(GoogleMapTagList) == 1:
                    Address = GoogleMapTagList[0].get_attribute("onclick")
                    Address = Address[Address.find("maps/search/") + 12: Address.find("','popup'")]
                    Address = Address.replace("1", "１").replace("2", "２").replace("3", "３").replace("4", "４").replace("5", "５").replace("6", "６").replace("7", "７").replace("8", "８").replace("9", "９").replace("0", "０")
                    CustomerAddress.append(CellText + "：" + Address)
                elif len(GoogleMapTagList) > 0:
                    raise Exception()
            elif re.match(".*(電話|手機).*", CellText) and not re.match(".*(電話行銷).*", CellText):
                Telephone = ""
                for ImageTag in CellTagList[IndexTag + 1].find_elements(WebBy.By.XPATH, "img"):
                    NumberList = ImageTag.get_attribute("src").split("|")
                    NumberList.pop(len(NumberList) - 1)
                    NumberList[0] = NumberList[0][len(NumberList[0]) - 2:]
                    Telephone += "".join([NumberMapDict[int(Number)] for Number in NumberList])
                if Telephone != "":
                    Telephone = Telephone.replace("-", " ").replace("(", " ").replace(")", " ").replace("  ", " ").replace("  ", " ").strip()
                    CustomerTel.append(Telephone)
        # 讀取客戶資料 End
        CustomerTel = ";".join(set(CustomerTel))
        CustomerAddress = ";".join(set(CustomerAddress))
        CustomerTable.loc[CustomerID, ["Name", "Birthday", "Telephone", "Address"]] = [CustomerName, CustomerBirth, CustomerTel, CustomerAddress]
        CustomerTable.to_csv("TableCustomer.csv", encoding="utf_8_sig", quoting=1)
    # 逐一客戶情報 End
    # 關閉戶關係管理系統 Begin
    CustomerTable.sort_index(inplace=True)
    CustomerTable.to_csv("TableCustomer.csv", encoding="utf_8_sig", quoting=1)
    ChromeDriver.close()
    ChromeDriver.switch_to.window(ChromeDriver.window_handles[0])
    # 關閉戶關係管理系統 End
# 重讀保戶總表情報 End
print("重讀保戶總表情報 ", datetime.datetime.now())
# 讀取保單健檢 Begin
if 1 == 1:
    # 進保戶關係管理系統 Begin
    ChromeDriver.execute_script(ChromeDriver.find_element(WebBy.By.XPATH, "//*[@id='idx_menu']/ul[2]/li[4]/ul/li[3]/ul/li[2]/a").get_attribute("onclick"))
    time.sleep(10)
    if len(ChromeDriver.window_handles) == 3:
        ChromeDriver.switch_to.window(ChromeDriver.window_handles[1])
        ChromeDriver.close()
    elif len(ChromeDriver.window_handles) > 3:
        raise Exception()
    ChromeDriver.switch_to.window(ChromeDriver.window_handles[1])
    # 進保戶關係管理系統 End
    CustomerTable = pandas.read_csv("TableCustomer.csv", index_col="ID", encoding="utf_8_sig", quoting=1, dtype=str).fillna("")
    PolicyTable = pandas.read_csv("TablePolicy.csv", encoding="utf_8_sig", quoting=1, dtype=str).fillna("")

    PolicyTable = PolicyTable[PolicyTable.index != PolicyTable.index]
    IndexFamily = 0

    # 進保單健檢頁 Begin
    ChromeDriver.switch_to.default_content()
    ChromeDriver.switch_to.frame(ChromeDriver.find_element(WebBy.By.XPATH, "/html/frameset/frame[1]"))
    ChromeDriver.find_element(WebBy.By.XPATH, "/html/body/table[1]/tbody/tr[2]/td[5]/a").click()
    try:
        WebUI.WebDriverWait(ChromeDriver, 10).until(WebConditions.alert_is_present())
        ChromeDriver.switch_to.alert.accept()
    except:
        ChromeDriver.switch_to.default_content()
    time.sleep(0.1)
    # 進保單健檢頁 End
    while True:
        # 點選最新家庭 Begin
        ChromeDriver.switch_to.default_content()
        ChromeDriver.switch_to.frame(ChromeDriver.find_element(WebBy.By.XPATH, "/html/frameset/frame[2]"))
        ChromeDriver.find_element(WebBy.By.XPATH, "//*[@id='aspnetForm']/table/tbody/tr/td[2]/table/tbody/tr[1]/td[2]/input").click()
        time.sleep(0.1)
        ChromeDriver.switch_to.default_content()
        ChromeDriver.switch_to.frame(ChromeDriver.find_element(WebBy.By.XPATH, "/html/frameset/frame[2]"))
        FamilyTagList = ChromeDriver.find_elements(WebBy.By.XPATH, "//*[@id='ctl00_HeaderContentHolder_Table1']/tbody/tr/td/a")
        if IndexFamily >= len(FamilyTagList):
            break
        FamilyTagList[IndexFamily].click()
        WebUI.WebDriverWait(ChromeDriver, 10).until(WebConditions.alert_is_present())
        ChromeDriver.switch_to.alert.accept()
        time.sleep(0.1)
        # 點選最新家庭 End
        # 讀取家庭成員資料 Begin
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
                [CustomerName, CustomerBirth, CustomerTel, CustomerAddress] = [[] if Value == "" else Value.split(";") for Value in CustomerTable.loc[CustomerID, ["Name", "Birthday", "Telephone", "Address"]]]
            CustomerName.append(CustomerRow.find_element(WebBy.By.XPATH, "td[2]").text)
            CustomerBirthDate = [int(Date) for Date in CustomerRow.find_element(WebBy.By.XPATH, "td[5]").text.split("/")]
            CustomerBirth.append(str(CustomerBirthDate[0] * 10000 + CustomerBirthDate[1] * 100 + CustomerBirthDate[2] + 19110000))
            [CustomerName, CustomerBirth, CustomerTel, CustomerAddress] = [";".join(set(Value)) for Value in [CustomerName, CustomerBirth, CustomerTel, CustomerAddress]]
            CustomerTable.loc[CustomerID] = [CustomerName, CustomerBirth, CustomerTel, CustomerAddress, FamilyHead]
        CustomerTable.to_csv("TableCustomer.csv", encoding="utf_8_sig", quoting=1)
        # 讀取家庭成員資料 End
        # 進入契約內容頁 Begin
        ChromeDriver.switch_to.default_content()
        ChromeDriver.switch_to.frame(ChromeDriver.find_element(WebBy.By.XPATH, "/html/frameset/frame[2]"))
        ChromeDriver.find_element(WebBy.By.XPATH, "//*[@id='ctl00_ctl00_HeaderContentHolder_Menu']/table[2]/tbody/tr/td[2]/a").click()
        time.sleep(0.1)
        # 進入契約內容頁 End
        LoadPolicyList = PolicyTable[PolicyTable.index != PolicyTable.index]
        ChromeDriver.switch_to.default_content()
        ChromeDriver.switch_to.frame(ChromeDriver.find_element(WebBy.By.XPATH, "/html/frameset/frame[2]"))
        for TagTable in ChromeDriver.find_elements(WebBy.By.XPATH, "//*[@id='aspnetForm']/table[2]/tbody/tr/td[2]/table[1]/tbody/tr/td"):
            # 讀取國泰契約內容 Begin
            if TagTable.find_element(WebBy.By.XPATH, "span").text == "國泰保單":
                PolicyRowList = TagTable.find_elements(WebBy.By.XPATH, "table/tbody/tr")
                PolicyRowList.pop(0)
                InsuredID = ""
                InsuredName = ""
                for PolicyRow in PolicyRowList:
                    if PolicyRow.text.startswith("被保險人"):
                        InsuredID = PolicyRow.find_element(WebBy.By.XPATH, "td/img").get_attribute("src")
                        InsuredID = InsuredID[InsuredID.find("Key=") + 4:]
                        InsuredID = InsuredID[: 10]
                        InsuredName = PolicyRow.text[5: PolicyRow.text.find("(")].strip()
                        continue
                    PolicyNumber = PolicyRow.find_element(WebBy.By.XPATH, "td[1]/img").get_attribute("src")
                    PolicyNumber = PolicyNumber[PolicyNumber.find("Key=") + 4:]
                    PolicyNumber = PolicyNumber[: 10]
                    ProductName = PolicyRow.find_element(WebBy.By.XPATH, "td[2]").text
                    PayTerm = PolicyRow.find_element(WebBy.By.XPATH, "td[3]").text
                    Relation = PolicyRow.find_element(WebBy.By.XPATH, "td[4]").text
                    EffectiveDate = [int(Date) for Date in PolicyRow.find_element(WebBy.By.XPATH, "td[5]").text.split("/")]
                    EffectiveDate = EffectiveDate[0] * 10000 + EffectiveDate[1] * 100 + EffectiveDate[2] + 19110000
                    SumAssured = PolicyRow.find_element(WebBy.By.XPATH, "td[6]").text.replace("\n", " ")
                    PayPremium = PolicyRow.find_element(WebBy.By.XPATH, "td[7]").text
                    PayType = PolicyRow.find_element(WebBy.By.XPATH, "td[8]").text
                    Beneficiary = PolicyRow.find_element(WebBy.By.XPATH, "td[9]").text.replace("\n", ";")
                    PayRoute = ""
                    OwnerName = ""
                    if ProductName.startswith("**要保人"):
                        IndexString = ProductName.find("\n")
                        OwnerName = ProductName[6: IndexString]
                        ProductName = ProductName[IndexString + 1:]
                        IndexString = PayPremium.find("元\n")
                        if IndexString >= 0:
                            PayRoute = PayPremium[IndexString + 2:] .replace("\n", "")
                            PayPremium = PayPremium[:IndexString + 1]
                        elif not PayPremium.endswith("元"):
                            PayRoute = PayPremium.replace("\n", "")
                            PayPremium = ""
                    Occupation = ""
                    CRMPolicyValue = ""
                    ProductNameList = ProductName.split("\n")
                    ProductName = ProductNameList.pop(0)
                    if ProductName.startswith("    "):
                        ProductName = "(附約)" + ProductName[4:]
                    for Item in ProductNameList:
                        if Item.startswith("    職業類別："):
                            Occupation = Item[9:]
                        elif Item.startswith("保單價值:"):
                            CRMPolicyValue = Item[5:]
                        else:
                            print("未知的商品附註 " + Item)
                    LoadPolicyList = LoadPolicyList.append(pandas.Series(["國泰人壽", InsuredID, InsuredName, PolicyNumber, OwnerName, ProductName, PayTerm, Relation, EffectiveDate, SumAssured, Occupation, CRMPolicyValue, PayPremium, PayType, PayRoute, Beneficiary, 0, 0, 0, 0], index=LoadPolicyList.columns), ignore_index=True)
            # 讀取國泰契約內容 End
            # 讀取他家公司契約內容 Begin
            elif TagTable.find_element(WebBy.By.XPATH, "span").text == "其他保單":
                PolicyRowList = TagTable.find_elements(WebBy.By.XPATH, "table/tbody/tr")
                PolicyRowList.pop(0)
                InsuredID = ""
                InsuredName = ""
                for PolicyRow in PolicyRowList:
                    if PolicyRow.text.startswith("被保險人"):
                        InsuredID = PolicyRow.find_element(WebBy.By.XPATH, "td/img").get_attribute("src")
                        InsuredID = InsuredID[InsuredID.find("Key=") + 4:]
                        InsuredID = InsuredID[: 10]
                        InsuredName = PolicyRow.text[5: PolicyRow.text.find("(")].strip()
                        continue
                    Company = PolicyRow.find_element(WebBy.By.XPATH, "td[1]").text
                    PolicyNumber = PolicyRow.find_element(WebBy.By.XPATH, "td[1]/img").get_attribute("src")
                    PolicyNumber = PolicyNumber[PolicyNumber.find("Key=(") + 5:]
                    PolicyNumber = PolicyNumber[: PolicyNumber.find(")&BC=")]
                    ProductName = PolicyRow.find_element(WebBy.By.XPATH, "td[2]").text
                    PayTerm = PolicyRow.find_element(WebBy.By.XPATH, "td[3]").text
                    Relation = PolicyRow.find_element(WebBy.By.XPATH, "td[4]").text
                    EffectiveDate = [int(Date) for Date in PolicyRow.find_element(WebBy.By.XPATH, "td[5]").text.split("/")]
                    EffectiveDate = EffectiveDate[0] * 10000 + EffectiveDate[1] * 100 + EffectiveDate[2] + 19110000
                    SumAssured = PolicyRow.find_element(WebBy.By.XPATH, "td[6]").text.replace("\n", " ")
                    PayPremium = PolicyRow.find_element(WebBy.By.XPATH, "td[7]").text
                    PayType = PolicyRow.find_element(WebBy.By.XPATH, "td[8]").text
                    DeathBeneficiary = PolicyRow.find_element(WebBy.By.XPATH, "td[9]").text
                    MatureBeneficiary = PolicyRow.find_element(WebBy.By.XPATH, "td[10]").text
                    AnnuityBeneficiary = PolicyRow.find_element(WebBy.By.XPATH, "td[11]").text
                    ExpireBeneficiary = PolicyRow.find_element(WebBy.By.XPATH, "td[12]").text
                    Beneficiary = []
                    if DeathBeneficiary != "":
                        Beneficiary.append("身故：" + DeathBeneficiary)
                    if MatureBeneficiary != "":
                        Beneficiary.append("滿期：" + MatureBeneficiary)
                    if AnnuityBeneficiary != "":
                        Beneficiary.append("年金：" + AnnuityBeneficiary)
                    if ExpireBeneficiary != "":
                        Beneficiary.append("祝壽：" + ExpireBeneficiary)
                    Beneficiary = ";".join(Beneficiary)
                    PayRoute = ""
                    OwnerName = ""
                    if ProductName.startswith("**要保人"):
                        IndexString = ProductName.find("\n")
                        OwnerName = ProductName[6: IndexString]
                        ProductName = ProductName[IndexString + 1:]
                    Occupation = ""
                    CRMPolicyValue = ""
                    ProductNameList = ProductName.split("\n")
                    ProductName = ProductNameList.pop(0)
                    if ProductName.startswith("    "):
                        ProductName = "(附約)" + ProductName[4:]
                    for Item in ProductNameList:
                        if Item.startswith("    職業類別："):
                            Occupation = Item[9:]
                        elif Item.startswith("保單價值:"):
                            CRMPolicyValue = Item[5:]
                        else:
                            print("未知的商品附註 " + Item)
                    LoadPolicyList = LoadPolicyList.append(pandas.Series([Company, InsuredID, InsuredName,  PolicyNumber, OwnerName,  ProductName,  PayTerm,  Relation, EffectiveDate,  SumAssured, Occupation, CRMPolicyValue, PayPremium, PayType,  PayRoute,  Beneficiary, 0, 0, 0, 0], index=LoadPolicyList.columns), ignore_index=True)
            # 讀取他家公司契約內容 End
        for [Company, InsuredID, PolicyNumber] in LoadPolicyList.groupby(["Company", "InsuredID", "PolicyNumber"]).indices:
            PolicyTable = PolicyTable[~((PolicyTable.Company == Company) & (PolicyTable.InsuredID == InsuredID) & (PolicyTable.PolicyNumber == PolicyNumber))]
        PolicyTable = PolicyTable.append(LoadPolicyList, ignore_index=True)
        PolicyTable.to_csv("TablePolicy.csv", index=False, encoding="utf_8_sig", quoting=1)
        IndexFamily += 1
    # 關閉戶關係管理系統 Begin
    ChromeDriver.close()
    ChromeDriver.switch_to.window(ChromeDriver.window_handles[0])
    # 關閉戶關係管理系統 End
    # 補要保人 Begin
    GroupTable = pandas.DataFrame(list(PolicyTable[PolicyTable.OwnerName != ""].groupby(["Company", "PolicyNumber", "OwnerName"]).indices.keys()), columns=["Company", "PolicyNumber", "OwnerName"])
    CheckTable = GroupTable.groupby(["Company", "PolicyNumber"]).count()
    if not CheckTable[CheckTable.OwnerName > 1].empty:
        raise Exception()
    for IndexGroup in GroupTable.index:
        [Company, PolicyNumber, OwnerName] = GroupTable.loc[IndexGroup]
        PolicyTable.loc[(PolicyTable.Company == Company) & (PolicyTable.PolicyNumber == PolicyNumber), "OwnerName"] = OwnerName
    # 補要保人 End
    CustomerTable.sort_index(inplace=True)
    CustomerTable.to_csv("TableCustomer.csv", encoding="utf_8_sig", quoting=1)
    PolicyTable.sort_values(by=["InsuredID", "Company", "PolicyNumber"], inplace=True)
    PolicyTable.to_csv("TablePolicy.csv", index=False, encoding="utf_8_sig", quoting=1)
# 讀取保單健檢 End
print("讀取保單健檢 ", datetime.datetime.now())
# 讀取核心契約情報 Begin
if 1 == 1:
    PolicyTable = pandas.read_csv("TablePolicy.csv", encoding="utf_8_sig", quoting=1, dtype=str).fillna("")
    for IndexCount, IndexPolicy in enumerate(PolicyTable.index):
        if IndexCount < 0:
            continue
        # 清 Catch Begin
        if IndexCount % 200 == 0:
            if len(ChromeDriver.window_handles) > 2:
                raise Exception()
            elif len(ChromeDriver.window_handles) == 2:
                ChromeDriver.close()
                ChromeDriver.switch_to.window(ChromeDriver.window_handles[0])
            ChromeDriver.get("chrome://settings/clearBrowserData")
            time.sleep(0.1)
            ClearButton = ChromeDriver.execute_script("return arguments[0].shadowRoot", ChromeDriver.find_element(WebBy.By.XPATH, "/html/body/settings-ui"))
            ClearButton = ChromeDriver.execute_script("return arguments[0].shadowRoot", ClearButton.find_element(WebBy.By.ID, "main"))
            ClearButton = ChromeDriver.execute_script("return arguments[0].shadowRoot", ClearButton.find_element(WebBy.By.CSS_SELECTOR, "settings-basic-page"))
            ClearButton = ChromeDriver.execute_script("return arguments[0].shadowRoot", ClearButton.find_element(WebBy.By.CSS_SELECTOR, "settings-privacy-page"))
            ClearButton = ChromeDriver.execute_script("return arguments[0].shadowRoot", ClearButton.find_element(WebBy.By.CSS_SELECTOR, "settings-clear-browsing-data-dialog"))
            ClearButton = ClearButton.find_element(WebBy.By.ID, "clearBrowsingDataConfirm")
            ClearButton.click()
            time.sleep(10)
            ChromeDriver.get("https://w3.cathaylife.com.tw")
            ChromeDriver.find_element(WebBy.By.ID, "UID").send_keys("A120923934")
            ChromeDriver.find_element(WebBy.By.ID, "KEY").send_keys("0938196099")
            ChromeDriver.find_element(WebBy.By.ID, "btnLogin").click()
            time.sleep(0.1)
            ChromeDriver.execute_script(ChromeDriver.find_element(WebBy.By.XPATH, "//*[@id='idx_menu']/ul[2]/li[3]/ul/li[2]/a").get_attribute("onclick"))
            time.sleep(1)
            ChromeDriver.switch_to.window(ChromeDriver.window_handles[1])
            ChromeDriver.switch_to.default_content()
            ChromeDriver.switch_to.frame(ChromeDriver.find_element(WebBy.By.XPATH, "/html/frameset/frameset/frame"))
            ChromeDriver.find_element(WebBy.By.XPATH, "//*[@id='showMainMenu']/li[1]").click()
            time.sleep(0.1)
            ChromeDriver.switch_to.default_content()
            ChromeDriver.switch_to.frame(ChromeDriver.find_element(WebBy.By.XPATH, "/html/frameset/frame[2]"))
            if ChromeDriver.find_element(WebBy.By.XPATH, "//*[@id='menutree']/ul/li[2]").text != "整合內容查詢":
                raise Exception()
            ChromeDriver.find_element(WebBy.By.XPATH, "//*[@id='menutree']/ul/li[3]").click()
            time.sleep(0.1)
        # 清 Catch End
        [Company, PolicyNumber, ProductName] = PolicyTable.loc[IndexPolicy, ["Company", "PolicyNumber", "ProductName"]]
        if Company != "國泰人壽" or ProductName.startswith("(附約)"):
            continue
        ChromeDriver.switch_to.default_content()
        ChromeDriver.switch_to.frame(ChromeDriver.find_element(WebBy.By.XPATH, "/html/frameset/frame[2]"))
        ChromeDriver.find_element(WebBy.By.XPATH, "//*[@id='menutree']/ul/li[3]/ul/li[3]/a").click()
        time.sleep(1)
        ChromeDriver.switch_to.default_content()
        ChromeDriver.switch_to.frame(ChromeDriver.find_element(WebBy.By.XPATH, "/html/frameset/frame[2]"))
        ChromeDriver.switch_to.frame(ChromeDriver.find_element(WebBy.By.ID, "mainFrame"))
        ChromeDriver.find_element(WebBy.By.ID, "POLICY_NO").send_keys(PolicyNumber)
        time.sleep(0.1)
        ChromeDriver.find_element(WebBy.By.ID, "btnTrial").click()
        time.sleep(0.1)
        ChromeDriver.switch_to.default_content()
        ChromeDriver.switch_to.frame(ChromeDriver.find_element(WebBy.By.XPATH, "/html/frameset/frame[2]"))
        ChromeDriver.switch_to.frame(ChromeDriver.find_element(WebBy.By.ID, "mainFrame"))
        CellList = ChromeDriver.find_elements(WebBy.By.XPATH, "//*[@id='form1']/table[6]/tbody/tr/td")
        [Bonus, PolicyValue, CashValue, Loan] = [0, 0, 0, 0]
        for IndexCell in range(len(CellList) - 1):
            CellField = CellList[IndexCell].text
            CellValue = CellList[IndexCell + 1].text.replace(",", "")
            if CellField == "保單價值準備金":
                PolicyValue = float(CellValue)
            elif CellField == "主約解約金額":
                CashValue = float(CellValue)
            elif CellField == "紅利金額":
                Bonus = float(CellValue)
            elif CellField == "借款金額":
                Loan += float(CellValue)
            elif CellField == "借款利息":
                Loan += float(CellValue)
            elif CellField == "借款延滯息":
                Loan += float(CellValue)
            elif CellField == "墊繳金額":
                Loan += float(CellValue)
            elif CellField == "墊繳利息	":
                Loan += float(CellValue)
        PolicyTable.loc[IndexPolicy, ["Bonus", "PolicyValue", "CashValue", "Loan"]] = [Bonus, PolicyValue, CashValue, Loan]
        PolicyTable.to_csv("TablePolicy.csv", index=False, encoding="utf_8_sig", quoting=1)
    PolicyTable.sort_values(by=["InsuredID", "Company", "PolicyNumber"], inplace=True)
    PolicyTable.to_csv("TablePolicy.csv", index=False, encoding="utf_8_sig", quoting=1)
    if len(ChromeDriver.window_handles) > 2:
        raise Exception()
    elif len(ChromeDriver.window_handles) == 2:
        ChromeDriver.close()
        ChromeDriver.switch_to.window(ChromeDriver.window_handles[0])
# 讀取核心契約情報 End
print("讀取核心契約情報 ", datetime.datetime.now())
ChromeDriver.quit()
print("程式結束 ", datetime.datetime.now())

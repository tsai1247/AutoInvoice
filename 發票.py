# about webdriver
import chromedriver_autoinstaller
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# about qt
import sys
from PyQt5.QtWidgets import QApplication, QWidget, QLabel

# others
from time import sleep
from os import getenv
from datetime import datetime as dt

BUYERID = '42524612'
ITEMID = 'A000001'
ITEMNAME = '商品名稱'
ITEMQTY = '1'
ITEMPRICE = '100'


class Element():
    def __init__(self):
        chromedriver_autoinstaller.install()
        self.driver: WebDriver = None
    
    def init(self):
        if self.driver is None:
            self.driver = webdriver.Chrome()
            self.driver.maximize_window()

    mainURL = r'https://www.einvoice.nat.gov.tw/'

    def get(self, url):
        self.driver.get(url)

    def find_element(self, by, value, timeout=10):
        WebDriverWait(self.driver, timeout).until(
            EC.presence_of_element_located((by, value))
        )
        return self.driver.find_element(by, value)
    
    def find_elements(self, by, value, timeout=10):
        WebDriverWait(self.driver, timeout).until(
            EC.presence_of_element_located((by, value))
        )
        return self.driver.find_elements(by, value)
    
    @property
    def alert(self):
        return self.driver.switch_to.alert

    @property
    def current_url(self):
        return self.driver.current_url

    @property
    def select_Identity(self):
        return Select(self.find_element(By.ID, 'loginSelect'))

    @property
    def btn_login(self):
        return self.find_element(By.XPATH, '//*[@id="loginBtn"]')

    @property
    def input_GUINumber(self):
        return self.find_element(By.XPATH, '//*[@id="l1_ban"]')

    @property
    def input_Account(self):
        return self.find_element(By.XPATH, '//*[@id="l1_userID"]')

    @property
    def input_Password(self):
        return self.find_element(By.XPATH, '//*[@id="l1_password"]')
    @property
    def input_CheckCode(self):
        return self.find_element(By.XPATH, '//*[@id="checkPicIndex"]')
    
    @property
    def btn_submit(self):
        return self.find_element(By.XPATH, '//*[@id="button"]')

    @property
    def button_Date(self):
        ret = self.find_elements(By.CSS_SELECTOR, 'img.ui-datepicker-trigger')
        if len(ret) == 1:
            return ret[0]
        else:
            return ret

    @property
    def button_Date_Today(self):
        today = dt.now().strftime('%d')
        return self.find_element(By.XPATH, f'//a[contains(@class, "ui-state-default")][contains(text(), "{today}")]')
    @property
    def button_ChooseReceiptHead(self):
        return self.find_element(By.XPATH, '//input[@id="invNoPrefix"]/../input[@type="button"]')

    @property
    def button_ReceiptHead(self):
        return self.find_element(By.CSS_SELECTOR, 'button')

    @property
    def button_ReceiptNumber(self):
        return self.find_element(By.XPATH, '//div[@class="ac_results"]')
    @property
    def input_ReceiptNumber(self):
        return self.find_element(By.XPATH, '//*[@id="invNoPrefix"]')

    @property
    def button_BuyerId(self):
        return self.find_element(By.XPATH, '//input[@id="targetBuyerId"]/../input[@type="button"]')

    @property
    def input_SearchBuyerId(self):
        return self.find_element(By.XPATH, '//*[@type="search"][@aria-controls="dataTablesBuyerCompany"]')

    def button_CandidateBuyerId(self, id: str):
        return self.find_element(By.XPATH, f'//a[contains(text(), "{id}")]')

    @property
    def input_ItemId(self):
        return self.find_element(By.XPATH, '//*[@id="productId1"]')
    
    @property
    def input_ItemName(self):
        return self.find_element(By.XPATH, '//*[@id="productName1"]')
    
    @property
    def input_ItemQty(self):
        return self.find_element(By.XPATH, '//*[@id="qty1"]')

    @property
    def input_ItemPrice(self):
        return self.find_element(By.XPATH, '//*[@id="unitPrice1"]')

    @property
    def button_SearchReceipt(self):
        return self.find_element(By.XPATH, '//a[@title="查詢"]')

    @property
    def button_ReceiptTitle(self):
        return self.find_elements(By.XPATH, '//input[not(@disabled)]/../../td[3]/a')[0]

    @property
    def button_SelectAll(self):
        return self.find_element(By.XPATH, '//a[@title="全選"]')
    
    @property
    def button_ToSendPage(self):
        return self.find_element(By.XPATH, '//a[@title="寄送"]')
    
    @property
    def input_ReceiptSendPassword(self):
        return self.find_element(By.XPATH, '//*[@id="password"]')
    
    @property
    def checkbox_ReceiptSendAgree(self):
        return self.find_element(By.XPATH, '//*[@id="cht"]')

    @property
    def button_ReceiptSend(self):
        return self.find_element(By.XPATH, '//a[@id="hyperlink"]')


class Crawler():
    def __init__(self, element: Element):
        self.element = element

    def run(self):
        self.element.init()

class Login(Crawler):
    def run(self):
        super().run()
        self.element.get(self.element.mainURL)
        self.element.btn_login.click()
        self.element.select_Identity.select_by_visible_text('營業人')
        self.element.input_GUINumber.send_keys(getenv('GUINumber'))
        self.element.input_Account.send_keys(getenv('Account'))
        self.element.input_Password.send_keys(getenv('Password'))
        self.element.input_CheckCode.click()
        sleep(1)
        
        while True:
            while len(self.element.input_CheckCode.get_attribute('value')) != 5:
                sleep(.2)
            
            self.element.btn_submit.click()
            
            try:
                while 'Main' not in self.element.current_url:
                    sleep(.2)
                break
            except:
                try:
                    self.element.alert.accept()
                except:
                    pass
                self.element.input_CheckCode.clear()
                self.element.input_CheckCode.click()
                continue

        self.element.url = r'/'.join(self.element.current_url.split(r'/')[:-1])
        

class SaveReceipt(Crawler):
    def run(self):
        super().run()
        self.element.get(self.element.url + '/MultiDelivery/invSave/InvSaveCreate!insert')
        
        self.element.button_Date.click()
        self.element.button_Date_Today.click()

        # 選擇發票號碼
        self.element.button_ChooseReceiptHead.click()
        current_window = self.element.driver.current_window_handle
        window_after = self.element.driver.window_handles[1]
        self.element.driver.switch_to.window(window_after)
        self.element.button_ReceiptHead.click()
        self.element.driver.switch_to.window(current_window)
        
        sleep(1)
        self.element.button_ReceiptNumber.click()
        # ActionChains(self.element.driver).move_to_element_with_offset(self.element.button_ReceiptNumber, 0, 0).click().perform()

        # 選擇買受人
        self.element.button_BuyerId.click()
        self.element.input_SearchBuyerId.send_keys(BUYERID)
        sleep(.5)
        self.element.button_CandidateBuyerId(BUYERID).click()

        # 輸入商品資訊
        self.element.input_ItemId.send_keys(ITEMID)
        self.element.input_ItemName.send_keys(ITEMNAME)
        self.element.input_ItemQty.send_keys(ITEMQTY)
        self.element.input_ItemPrice.send_keys(ITEMPRICE, Keys.TAB)

        # 等待確認
        while self.element.current_url != self.element.url + '/MultiDelivery/invSave/InvSaveCreate!insert':
            sleep(.2)

        sleep(1)

class SendReceipt(Crawler):
    def run(self):
        super().run()
        self.element.get(self.element.url + '/MultiDelivery/SelfDelivery')
        self.element.button_Date[0].click()
        self.element.button_Date_Today.click()
        
        self.element.button_Date[1].click()
        self.element.button_Date_Today.click()
        self.element.button_SearchReceipt.click()
        self.element.button_ReceiptTitle.click()
        self.element.button_SelectAll.click()
        self.element.button_ToSendPage.click()
        self.element.alert.accept()
        self.element.input_ReceiptSendPassword.send_keys(getenv('Account'))
        self.element.checkbox_ReceiptSendAgree.click()
        # self.element.button_ReceiptSend
        self.element.button_ReceiptSend.click()

        sleep(1000)

class View():
    def __init__(self):
        self.app = QApplication(sys.argv)
        self.widget = QWidget()

    def show(self):
        
        textLabel = QLabel(self.widget)
        textLabel.setText("Hello World!")
        textLabel.move(110,85)

        self.widget.setGeometry(50,50,320,200)
        self.widget.setWindowTitle("PyQt5 Example")
        self.widget.show()
        sys.exit(self.app.exec_())

class Main():
    def __init__(self):
        element = Element()
        login = Login(element)
        savereceipt = SaveReceipt(element)
        #sendreceipt = SendReceipt(element).run()

        view = View()
        view.show()
        


if __name__ == '__main__':
    Main()
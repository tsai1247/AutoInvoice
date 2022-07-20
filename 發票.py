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
from PyQt5.QtCore import Qt, QDateTime, QDate, QTime
from PyQt5.QtWidgets import QDateEdit, QApplication, QWidget, QLabel, QLineEdit, QPushButton, QGridLayout, QHBoxLayout, QVBoxLayout, QComboBox, QMessageBox
from PyQt5.QtGui import QFont

# others
from time import sleep
from os import getenv
from datetime import datetime as dt
from functools import partial

# BUYERID = '42524612'
# ITEMID = 'A000001'
# ITEMNAME = '商品名稱'
# ITEMQTY = '1'
# ITEMPRICE = '100'


class Element():
    def __init__(self):
        chromedriver_autoinstaller.install()
        self.driver: WebDriver = None
    
    def init(self):
        if self.driver is None:
            self.driver = webdriver.Chrome()
        self.maximize()

    mainURL = r'https://www.einvoice.nat.gov.tw/'

    def get(self, url):
        self.driver.get(url)
    
    def minimize(self):
        self.driver.minimize_window()
    def maximize(self):
        self.driver.maximize_window()

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
    def button_DateChooser(self):
        ret = self.find_elements(By.CSS_SELECTOR, 'img.ui-datepicker-trigger')
        if len(ret) == 1:
            return ret[0]
        else:
            return ret

    @property
    def button_Date_Today(self):
        today = dt.now().strftime('%d')
        return self.find_element(By.XPATH, f'//a[contains(@class, "ui-state-default")][contains(text(), "{today}")]')

    def button_Date(self, date: int):
        return self.find_element(By.XPATH, f'//a[contains(@class, "ui-state-default")][contains(text(), "{date}")]')


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
        self.element.minimize()
        

class SaveReceipt(Crawler):
    def run(self, date: int, buyerId: QLineEdit, itemId: QLineEdit, itemName: QLineEdit, itemQty: QLineEdit, itemPrice: QLineEdit):
        super().run()

        buyerId = buyerId.text()
        itemId = itemId.text()
        itemName = itemName.text()
        itemQty = itemQty.text()
        itemPrice = itemPrice.text()

        self.element.get(self.element.url + '/MultiDelivery/invSave/InvSaveCreate!insert#')
        
        self.element.button_DateChooser.click()
        self.element.button_Date(date).click()

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
        self.element.input_SearchBuyerId.send_keys(buyerId)
        sleep(.5)
        self.element.button_CandidateBuyerId(buyerId).click()

        # 輸入商品資訊
        self.element.input_ItemId.send_keys(itemId)
        self.element.input_ItemName.send_keys(itemName)
        self.element.input_ItemQty.send_keys(itemQty)
        self.element.input_ItemPrice.send_keys(itemPrice, Keys.TAB)

        origin_url = self.element.current_url
        # 等待確認
        while self.element.current_url == origin_url:
            sleep(.2)

        sleep(1)

class SendReceipt(Crawler):
    def run(self, date):
        super().run()
        self.element.get(self.element.url + '/MultiDelivery/SelfDelivery')
        self.element.button_DateChooser[0].click()
        self.element.button_Date(date).click()
        
        self.element.button_DateChooser[1].click()
        self.element.button_Date(date).click()
        self.element.button_SearchReceipt.click()
        self.element.button_ReceiptTitle.click()
        try:
            self.element.alert.accept()
        except:
            pass
        
        self.element.button_SelectAll.click()
        self.element.button_ToSendPage.click()
        self.element.alert.accept()
        self.element.input_ReceiptSendPassword.send_keys(getenv('Account'))
        self.element.checkbox_ReceiptSendAgree.click()
        # self.element.button_ReceiptSend
        self.element.button_ReceiptSend.click()

        self.element.minimize()

class View():
    def __init__(self):
        self.app = QApplication(sys.argv)
        self.widget = QWidget()
        # set always on top
        self.widget.setWindowFlags(Qt.WindowStaysOnTopHint)

        # BUYERID = '42524612'
        # ITEMID = 'A000001'
        # ITEMNAME = '商品名稱'
        # ITEMQTY = '1'
        # ITEMPRICE = '100'

        self.defaultdata = {
            'BuyerId': getenv('DefaultBuyerId'),
            'ItemId': getenv('DefaultItemId'),
            'ItemName': getenv('DefaultItemName'),
            'ItemQty': getenv('DefaultItemQty'),
            'ItemPrice': getenv('DefaultItemPrice')
        }
        self.set_window()

    def set_window(self):
        # title
        self.widget.setWindowTitle('電子發票寄送')

        # grid 7 row * 2 column
        ''' the layout is like:
        日期, 2022/02/02
        統一編號, '42524612'
        商品ID, 'A000001'
        名稱 = '商品名稱'
        數量 = '1'
        價格 = '100'
        Send Button(column span 2 row span 1)
        '''
        # all fontsize is 14
        font = QFont('微軟正黑體', 14)

        grid = QGridLayout()
        grid.setSpacing(10)
        self.widget.setLayout(grid)

        # date chooser
        self.label_DateChooser = QLabel('日期')
        self.label_DateChooser.setAlignment(Qt.AlignCenter)
        self.label_DateChooser.setFont(font)
        grid.addWidget(self.label_DateChooser, 0, 0)

        self.date = QDateEdit()
        self.date.setDate(QDate.currentDate())
        self.date.setCalendarPopup(True)
        self.date.setDisplayFormat('yyyy/MM/dd')
        self.date.setMinimumDate(QDate.currentDate().addDays(-QDate.currentDate().day() + 1))
        self.date.setMaximumDate(QDate.currentDate().addDays(3))
        self.date.setDate(QDate.currentDate())
        self.date.setFont(font)
        grid.addWidget(self.date, 0, 1)

        # 統一編號
        self.label_GUINumber = QLabel('統一編號')
        self.label_GUINumber.setAlignment(Qt.AlignCenter)
        self.label_GUINumber.setFont(font)
        grid.addWidget(self.label_GUINumber, 1, 0)


        self.label_GUINumber_Value = QLineEdit(self.defaultdata['BuyerId'])
        self.label_GUINumber_Value.setFont(font)
        grid.addWidget(self.label_GUINumber_Value, 1, 1)

        # 物品ID
        self.label_ItemId = QLabel('商品ID')
        self.label_ItemId.setAlignment(Qt.AlignCenter)
        self.label_ItemId.setFont(font)
        grid.addWidget(self.label_ItemId, 2, 0)
        
        self.label_ItemId_Value = QLineEdit(self.defaultdata['ItemId'])
        self.label_ItemId_Value.setFont(font)
        grid.addWidget(self.label_ItemId_Value, 2, 1)

        # 名稱
        self.label_ItemName = QLabel('名稱')
        self.label_ItemName.setAlignment(Qt.AlignCenter)
        self.label_ItemName.setFont(font)
        grid.addWidget(self.label_ItemName, 3, 0)

        self.label_ItemName_Value = QLineEdit(self.defaultdata['ItemName'])
        self.label_ItemName_Value.setFont(font)
        grid.addWidget(self.label_ItemName_Value, 3, 1)

        # 數量
        self.label_ItemQty = QLabel('數量')
        self.label_ItemQty.setAlignment(Qt.AlignCenter)
        self.label_ItemQty.setFont(font)
        grid.addWidget(self.label_ItemQty, 4, 0)

        self.label_ItemQty_Value = QLineEdit(self.defaultdata['ItemQty'])
        self.label_ItemQty_Value.setFont(font)
        grid.addWidget(self.label_ItemQty_Value, 4, 1)

        # 價格
        self.label_ItemPrice = QLabel('價格')
        self.label_ItemPrice.setAlignment(Qt.AlignCenter)
        self.label_ItemPrice.setFont(font)
        grid.addWidget(self.label_ItemPrice, 5, 0)

        self.label_ItemPrice_Value = QLineEdit(self.defaultdata['ItemPrice'])
        self.label_ItemPrice_Value.setFont(font)
        grid.addWidget(self.label_ItemPrice_Value, 5, 1)

        # Send Button
        self.button_Send = QPushButton('Send')
        self.button_Send.setFont(font)
        grid.addWidget(self.button_Send, 6, 0, 1, 2)


    def show(self):
        # show window
        self.widget.show()

        # 還原視窗
        self.widget.showNormal()
        

        sys.exit(self.app.exec_())

class Main():
    def __init__(self):
        element = Element()
        self.login = Login(element)

        self.savereceipt = SaveReceipt(element)
        self.sendreceipt = SendReceipt(element)

        view = View()
        view.button_Send.clicked.connect(partial(self.executeCrawler, 
            view.label_GUINumber_Value, 
            view.label_ItemId_Value, 
            view.label_ItemName_Value, 
            view.label_ItemQty_Value, 
            view.label_ItemPrice_Value
        ))
        
        # processing...
        self.login.run()
        view.show()

    def executeCrawler(self, date, GUINumber, ItemId, ItemName, ItemQty, ItemPrice):
           self.savereceipt.run(date, GUINumber, ItemId, ItemName, ItemQty, ItemPrice)
           self.sendreceipt.run(date)



if __name__ == '__main__':
    Main()
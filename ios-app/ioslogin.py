from appium.webdriver.common.appiumby import AppiumBy
from time import sleep
from appium.webdriver.common.mobileby import MobileBy



class ioslogin:
	def __init__(driver):
            sleep(5)
            #關閉通知
            el1 = driver.find_element(by=AppiumBy.ACCESSIBILITY_ID, value="不允許")
            el1.click()
            #帳號密碼
            sleep(3)
            
            '''action = TouchAction(driver)
            # 在螢幕上的特定位置模擬點擊
            x = 70
            y = 199
            action.tap(None, x, y).perform()'''
            

            element = driver.find_element(by=MobileBy.IOS_CLASS_CHAIN, value='**/XCUIElementTypeTextField[`value == "請輸入身分證字號"`]')
            element.click()
            element.send_keys("F123310943")
            sleep(3)
            el3 = driver.find_element(by=MobileBy.IOS_CLASS_CHAIN, value='**/XCUIElementTypeToolbar[`label == "工具列"`]/XCUIElementTypeOther/XCUIElementTypeOther/XCUIElementTypeButton')
            el3.click()
            
            sleep(3)
            el4 = driver.find_element(by=MobileBy.IOS_CLASS_CHAIN, value='**/XCUIElementTypeSecureTextField[`value == "請輸入電子交易密碼"`]')
            el4.click()
            el4.send_keys("abc123")
            sleep(3)
            el5 = driver.find_element(by=MobileBy.IOS_CLASS_CHAIN, value='**/XCUIElementTypeToolbar[`label == "工具列"`]/XCUIElementTypeOther/XCUIElementTypeOther/XCUIElementTypeButton')
            el5.click()
            #關閉雙因子登入
            el6 = driver.find_element(by=AppiumBy.ACCESSIBILITY_ID, value="version")
            el6.click()
            el6 = driver.find_element(by=AppiumBy.ACCESSIBILITY_ID, value="version")
            el6.click()
            el6 = driver.find_element(by=AppiumBy.ACCESSIBILITY_ID, value="version")
            el6.click()
            el6 = driver.find_element(by=AppiumBy.ACCESSIBILITY_ID, value="version")
            el6.click()
            el6 = driver.find_element(by=AppiumBy.ACCESSIBILITY_ID, value="version")
            el6.click()
            el6 = driver.find_element(by=AppiumBy.ACCESSIBILITY_ID, value="version")
            el6.click()
            el6 = driver.find_element(by=AppiumBy.ACCESSIBILITY_ID, value="version")
            el6.click()
            el6 = driver.find_element(by=AppiumBy.ACCESSIBILITY_ID, value="version")
            el6.click()
            el6 = driver.find_element(by=AppiumBy.ACCESSIBILITY_ID, value="version")
            el6.click()
            el6 = driver.find_element(by=AppiumBy.ACCESSIBILITY_ID, value="version")
            el6.click()
            el6 = driver.find_element(by=AppiumBy.ACCESSIBILITY_ID, value="version")
            el6.click()
            el6 = driver.find_element(by=AppiumBy.ACCESSIBILITY_ID, value="version")
            el6.click()

            el7 = driver.find_element(by=AppiumBy.ACCESSIBILITY_ID, value="Button_Login")
            el7.click()

            sleep(35)
            el8 = driver.find_element(by=MobileBy.IOS_CLASS_CHAIN, value='**/XCUIElementTypeButton[`label == "暫時不用"`]')
            el8.click()
            el9 = driver.find_element(by=AppiumBy.ACCESSIBILITY_ID, value="CathayIntroduction close")
            el9.click()


            








from appium.webdriver.common.appiumby import AppiumBy

class ioslogin:
	def __init__(driver):

            el1 = driver.find_element(by=AppiumBy.XPATH, value="//XCUIElementTypeApplication[@name=\"樹精靈(測)\"]/XCUIElementTypeWindow/XCUIElementTypeOther[1]/XCUIElementTypeOther/XCUIElementTypeOther/XCUIElementTypeOther/XCUIElementTypeOther/XCUIElementTypeOther/XCUIElementTypeOther[1]/XCUIElementTypeOther/XCUIElementTypeOther/XCUIElementTypeOther[2]/XCUIElementTypeOther[1]/XCUIElementTypeOther[1]/XCUIElementTypeTextField")
            el1.click()
            el1.send_keys("F123310943")
            el1.click()
        
            el3 = driver.find_element(by=AppiumBy.XPATH, value="//XCUIElementTypeApplication[@name=\"樹精靈(測)\"]/XCUIElementTypeWindow[1]/XCUIElementTypeOther[2]/XCUIElementTypeOther/XCUIElementTypeOther/XCUIElementTypeOther/XCUIElementTypeOther/XCUIElementTypeOther/XCUIElementTypeOther[1]/XCUIElementTypeOther/XCUIElementTypeOther/XCUIElementTypeOther[2]/XCUIElementTypeOther[1]/XCUIElementTypeOther[2]/XCUIElementTypeSecureTextField")
            el3.send_keys("abc123")
            el3.click()
            el4 = driver.find_element(by=AppiumBy.XPATH, value="//XCUIElementTypeToolbar[@name=\"Toolbar\"]/XCUIElementTypeOther/XCUIElementTypeOther/XCUIElementTypeButton")
            el4.click()

            #關閉雙因子
            el4 = driver.find_element(by=AppiumBy.ACCESSIBILITY_ID, value="3.3.1 R03T")
            el4.click()
            el4 = driver.find_element(by=AppiumBy.ACCESSIBILITY_ID, value="3.3.1 R03T")
            el4.click()
            el4 = driver.find_element(by=AppiumBy.ACCESSIBILITY_ID, value="3.3.1 R03T")
            el4.click()
            el4 = driver.find_element(by=AppiumBy.ACCESSIBILITY_ID, value="3.3.1 R03T")
            el4.click()
            el4 = driver.find_element(by=AppiumBy.ACCESSIBILITY_ID, value="3.3.1 R03T")
            el4.click()
            el4 = driver.find_element(by=AppiumBy.ACCESSIBILITY_ID, value="3.3.1 R03T")
            el4.click()
            el4 = driver.find_element(by=AppiumBy.ACCESSIBILITY_ID, value="3.3.1 R03T")
            el4.click()
            el4 = driver.find_element(by=AppiumBy.ACCESSIBILITY_ID, value="3.3.1 R03T")
            el4.click()
            el4 = driver.find_element(by=AppiumBy.ACCESSIBILITY_ID, value="3.3.1 R03T")
            el4.click()
            el4 = driver.find_element(by=AppiumBy.ACCESSIBILITY_ID, value="3.3.1 R03T")
            el4.click()
            el4 = driver.find_element(by=AppiumBy.ACCESSIBILITY_ID, value="3.3.1 R03T")
            el4.click()
            el4 = driver.find_element(by=AppiumBy.ACCESSIBILITY_ID, value="3.3.1 R03T")
            el4.click()








from appium.webdriver.common.appiumby import AppiumBy

class Cancel:
	def __init__(driver):
                el1 = driver.find_element(by=AppiumBy.ID, value="com.cathay.securities.mBroker:id/ivLeft")
                el1.click()
                print("關閉推播")

    
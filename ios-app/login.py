from appium.webdriver.common.appiumby import AppiumBy

class Login:
	def __init__(driver):
		driver.implicitly_wait(10)
		#輸入登入帳密
		el1 = driver.find_element(by=AppiumBy.ID, value="com.cathay.securities.mBroker:id/editText_ID")
		el1.send_keys("F123310943")
		el2 = driver.find_element(by=AppiumBy.ID, value="com.cathay.securities.mBroker:id/editText_Password")
		el2.send_keys("abc123")
		#雙因子關閉
		el3 = driver.find_element(by=AppiumBy.ID, value="com.cathay.securities.mBroker:id/textView_VersionName")
		el3.click()
		el3.click()
		el3.click()
		el3.click()
		el3.click()
		el3.click()
		el3.click()
		el3.click()
		el3.click()
		el3.click()
		el3.click()

		#點擊登入
		el4 = driver.find_element(by=AppiumBy.ID, value="com.cathay.securities.mBroker:id/button_Login")
		el4.click()

		driver.implicitly_wait(20)
		#關閉生物辨識
		el5 = driver.find_element(by=AppiumBy.ID, value="com.cathay.securities.mBroker:id/btn_fingerprint_guide_dialog_cancel")
		el5.click()
		#美股延遲報價
		try:
			el6 = driver.find_element(by=AppiumBy.ID, value="com.cathay.securities.mBroker:id/ivLeft")
			el6.click()
		except:
			print("沒有美股延遲報價")	
		#截圖提示確定
		el7 = driver.find_element(by=AppiumBy.ID, value="com.cathay.securities.mBroker:id/button_No")
		el7.click()
		#關閉存股助你發
		el8 = driver.find_element(by=AppiumBy.ID, value="com.cathay.securities.mBroker:id/ivSkip")
		el8.click()

		#進入自選庫存畫面
		#el9 = driver.find_element(by=AppiumBy.ACCESSIBILITY_ID, value="樹精靈(測)")
		#el9.click()
		#el10 = driver.find_element(by=AppiumBy.ID, value="com.cathay.securities.mBroker:id/relativeLayout_Portfolio")
		#el10.click()
		

	
from appium.webdriver.common.appiumby import AppiumBy
from compare import CompareName ,Leave ,KLine 
from excelTest import PrintExcel
from word import CompanyName,BuyAndSell

class SearchAapl:
	def __init__(driver):
			
			el2 = driver.find_element(by=AppiumBy.XPATH, value="/hierarchy/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.RelativeLayout/androidx.drawerlayout.widget.DrawerLayout/android.widget.RelativeLayout/android.widget.FrameLayout/android.widget.LinearLayout/androidx.viewpager.widget.ViewPager/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.GridView/android.widget.FrameLayout[2]/android.widget.RelativeLayout/android.widget.RelativeLayout/android.widget.LinearLayout")
			el2.click()
			aapl = driver.find_element(by=AppiumBy.XPATH, value="/hierarchy/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.RelativeLayout/androidx.viewpager.widget.ViewPager/android.widget.LinearLayout/android.view.ViewGroup/android.widget.LinearLayout/android.widget.TextView")
			companyName = aapl.text

			testName = companyName+" 名稱測試"
			print(testName)

			result="正確"
			compareName=CompareName.name(driver)

			if(CompareName.name(driver)!="蘋果"):
				print("名稱有誤")
				result="錯誤,名稱不同"
			else:
				print("兩者相同")
			if(companyName != "AAPL"):
				print("不同")
				result="錯誤,名稱不同"
				PrintExcel.testCompany(driver,testName,result) 
				CompanyName.companyName(driver,testName,result,"","")
			else:
				print("兩者相同")	
				PrintExcel.testCompany(driver,testName,result) 	
				CompanyName.companyName(driver,testName,result,"","")

			#檢查K線
			KLine.kLine(driver)
			KLine.manual(driver)
			KLine.getKLineName(driver)
			KLine.compareKLineName(driver,companyName)
			KLine.leaveKLine(driver)


			SearchAapl.buyAndSell(driver,testName,compareName,companyName)

	def buyAndSell(driver,testName,compareName,companyName):
		#點擊買進
		enterBuy = driver.find_element(by=AppiumBy.ID, value="com.cathay.securities.mBroker:id/BuyText")
		enterBuy.click()

		driver.implicitly_wait(10)  		
		#關閉說明書
		try:
			el2 = driver.find_element(by=AppiumBy.ID, value="com.cathay.securities.mBroker:id/button_No")
			el2.click()
		except:
			print()	
		el3 = driver.find_element(by=AppiumBy.ID, value="com.cathay.securities.mBroker:id/ivBtn")
		el3.click()
		el4 = driver.find_element(by=AppiumBy.ID, value="com.cathay.securities.mBroker:id/ivBtn")
		el4.click()
		el5 = driver.find_element(by=AppiumBy.ID, value="com.cathay.securities.mBroker:id/ivBtn")
		el5.click()

		#先離開,再進來 關閉說明書後 名字會變成AAPL 之後進來都會是蘋果
		Leave.leave(driver)
		enterBuy = driver.find_element(by=AppiumBy.ID, value="com.cathay.securities.mBroker:id/BuyText")
		enterBuy.click()

		buyName = driver.find_element(by=AppiumBy.ID, value="com.cathay.securities.mBroker:id/tvSymbolName") #AAPL
		buyName.click()
		
		if(compareName != buyName.text):
			result = "錯誤"
			print(buyName.text)
			print(compareName+"名稱有誤")
			PrintExcel.clickBuyAndSell(driver,testName,"(買進)",result)
			BuyAndSell.buyAndSell(driver,companyName,result,"","","買進")

		else:
			result = "正確"
			print(buyName.text)
			print("名稱一樣")	
			PrintExcel.clickBuyAndSell(driver,testName,"(買進)",result)
			BuyAndSell.buyAndSell(driver,companyName,result,"","","買進")

		Leave.leave(driver)

		EnterSell = driver.find_element(by=AppiumBy.ID, value="com.cathay.securities.mBroker:id/btn_order_sell")
		EnterSell.click()

		sellName = driver.find_element(by=AppiumBy.ID, value="com.cathay.securities.mBroker:id/tvSymbolName")
		sellName.click()

		if(compareName != sellName.text):
			result = "錯誤"
			print(sellName.text)
			print(compareName + "名稱有誤")
			PrintExcel.clickBuyAndSell(driver,testName,"(賣出)",result)
			BuyAndSell.buyAndSell(driver,companyName,result,"","","賣出")

		else:	
			result = "正確"
			print(sellName.text)
			print("名稱一樣")	
			PrintExcel.clickBuyAndSell(driver,testName,"(賣出)",result)
			BuyAndSell.buyAndSell(driver,companyName,result,"","","賣出")

		Leave.leave(driver)


 
		el2 = driver.find_element(by=AppiumBy.XPATH, value="/hierarchy/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.RelativeLayout/android.widget.LinearLayout/android.widget.LinearLayout[1]/android.widget.RelativeLayout/android.widget.TextView")
		el2.click()


from appium.webdriver.common.appiumby import AppiumBy
from compare import CompareName , OddLot ,Leave ,KLine , CompareNumber, ClickBuyAndSell
from excelTest import PrintExcel
from word import CompanyName

class SearchTsmc:
	def __init__(driver):

			tsmc = driver.find_element(by=AppiumBy.XPATH, value="/hierarchy/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.RelativeLayout/androidx.drawerlayout.widget.DrawerLayout/android.widget.RelativeLayout/android.widget.FrameLayout/android.widget.LinearLayout/androidx.viewpager.widget.ViewPager/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.GridView/android.widget.FrameLayout[3]/android.widget.RelativeLayout/android.widget.RelativeLayout/android.widget.TextView")
			companyName = tsmc.text
			tsmc.click()

			testName = companyName+"名稱測試" 
			print(testName)

			result="正確"
			#檢測點擊後"名稱"、"零股"、"K線" 名稱是否一樣
			if(companyName != CompareName.name(driver)):
				print("不同")
				result="錯誤,名稱不同"
				PrintExcel.testCompany(driver,testName,result) 
				CompanyName.companyName(driver,testName,result,"","")
			else:
				print("兩者相同")	
				PrintExcel.testCompany(driver,testName,result) 
				CompanyName.companyName(driver,testName,result,"","")

			#檢查零股
			OddLot.oddLot(driver)
			OddLot.getOddLotName(driver)
			OddLot.compareOddLotName(driver,companyName)
			Leave.leave(driver)

			#檢查K線
			KLine.kLine(driver)
			KLine.manual(driver)
			KLine.getKLineName(driver)
			KLine.compareKLineName(driver,companyName)
			KLine.leaveKLine(driver)

			CompareNumber.doList(driver,companyName)	#檢查買入 買出 買賣超

			#檢查買進、賣出
			
			#買進
			buy = "買進"
			ClickBuyAndSell.clickBuy(driver)
			ClickBuyAndSell.buyAndSellManual(driver)
			#先離開 原因：第一次進去，預設在中間 看不到上面的字
			Leave.leave(driver)
			#再進來
			ClickBuyAndSell.clickBuy(driver)
			
			ClickBuyAndSell.getName(driver)
			ClickBuyAndSell.clickName(driver,companyName,buy)
			Leave.leave(driver)

			#賣出
			sell = "賣出"
			ClickBuyAndSell.clickSell(driver)
			#關閉說明書
			ClickBuyAndSell.buyAndSellManual(driver)
			#先離開 原因：第一次進去，預設在中間 看不到上面的字
			Leave.leave(driver)
			#再進來
			ClickBuyAndSell.clickSell(driver)
			#取得買進賣出名稱及檢查
			ClickBuyAndSell.getName(driver)
			ClickBuyAndSell.clickName(driver,companyName,sell)
			#離開
			Leave.leave(driver)

			#離開此公司畫面
			el2 = driver.find_element(by=AppiumBy.XPATH, value="/hierarchy/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.RelativeLayout/android.widget.LinearLayout/android.widget.LinearLayout[1]/android.widget.RelativeLayout/android.widget.TextView")
			el2.click()	
			print("離開"+ testName)
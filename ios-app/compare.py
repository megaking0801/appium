from appium.webdriver.common.appiumby import AppiumBy
import openpyxl
from word import TestKLine,Error
from time import sleep
from selenium.webdriver.common.action_chains import ActionChains,ActionBuilder
from selenium.webdriver.common.actions.pointer_input import PointerInput
from selenium.webdriver.common.actions import interaction
#確定名稱
class CompareName:
	def name(driver):
		sleep(3)
		getName = driver.find_element(by=AppiumBy.XPATH, value="//XCUIElementTypeApplication[@name=\"樹精靈(測)\"]/XCUIElementTypeWindow[1]/XCUIElementTypeOther/XCUIElementTypeOther/XCUIElementTypeOther/XCUIElementTypeOther/XCUIElementTypeOther/XCUIElementTypeOther/XCUIElementTypeOther[1]/XCUIElementTypeOther/XCUIElementTypeOther/XCUIElementTypeOther/XCUIElementTypeOther/XCUIElementTypeScrollView/XCUIElementTypeOther/XCUIElementTypeOther/XCUIElementTypeOther/XCUIElementTypeOther[2]/XCUIElementTypeOther[1]")
		setName = getName.text
		return setName 

#進入零股畫面
class OddLot:
	def oddLot(driver):
		enterOddLot = driver.find_element(by=AppiumBy.XPATH, value="//XCUIElementTypeStaticText[@name=\"零股\"]") #零股
		enterOddLot.click()
		return enterOddLot

#零股名稱
	'''def getOddLotName(driver):
		oddLot = driver.find_element(by=AppiumBy.ACCESSIBILITY_ID, value="精誠 6214")
		oddLot.click
		oddLotName = oddLot.text
		return oddLotName'''

#檢查零股名稱
	'''def compareOddLotName(driver,companyName):
		result="正確"	
		if(companyName != OddLot.getOddLotName(driver)):
			print("零股不相同")
			result="錯誤,零股名稱不同"
			PrintExcel.testOddLot(driver,companyName,result)
			TestOddLet.oddLet(driver,companyName,result,"測試圖片","這邊打備註")
		else:
			print("零股相同")

			PrintExcel.testOddLot(driver,companyName,result)	
			TestOddLet.oddLet(driver,companyName,result,"測試圖片","這邊打備註")	
		return result'''

#離開畫面	
class Leave:
	def leave(driver):	
		#防止畫面跳轉太慢，強制睡眠3秒
		sleep(3)
		el9 = driver.find_element(by=AppiumBy.XPATH, value="//XCUIElementTypeButton[@name=\"零股\"]")
		sleep(3)
		el9.click()
			


		
class KLine:
	#進入K線畫面
	def kLine(driver):
		el2 = driver.find_element(by=AppiumBy.XPATH, value="//XCUIElementTypeStaticText[@name=\"K線\"]")
		el2.click()

#偵測是否有說明書	有的話關閉
	def manual(driver):
		try:
			el7 = driver.find_element(by=AppiumBy.ACCESSIBILITY_ID, value="MBkResourcesImage Introduction")
			sleep(3)
			el7.click()
			sleep(3)
			el8 = driver.find_element(by=AppiumBy.ACCESSIBILITY_ID, value="MBkResourcesImage Introduction")
			sleep(3)
			el8.click()
			sleep(3)
		except:
			print("沒有說明書")		
			PrintExcel.error(driver)
			Error.error(driver)
			

#取得K線名稱
	def getKLineName(driver):
		try:
			getName = driver.find_element(by=AppiumBy.XPATH, value="//XCUIElementTypeApplication[@name=\"樹精靈(測)\"]/XCUIElementTypeWindow[1]/XCUIElementTypeOther/XCUIElementTypeOther/XCUIElementTypeOther/XCUIElementTypeOther/XCUIElementTypeOther/XCUIElementTypeOther/XCUIElementTypeOther[1]/XCUIElementTypeOther/XCUIElementTypeOther[1]")
			kLineName = (getName.text +" ").strip()
			print("KLineName:"+kLineName)
		except:
			kLineName != getName.text
			PrintExcel.error(driver)
			Error.error(driver)
		return kLineName

#檢查K線
	def compareKLineName(driver,companyName):
		result="正確"
		if(companyName != KLine.getKLineName(driver)):
			print("K線不相同")
			result="錯誤,K線名稱不同"
			PrintExcel.testKLine(driver,companyName,result)
			TestKLine.kLine(driver,companyName,result,"測試圖片","這邊打備註")
		else:
			print("K線相同")	
			PrintExcel.testKLine(driver,companyName,result)
			TestKLine.kLine(driver,companyName,result,"測試圖片","這邊打備註")

	

#離開K線
	def leaveKLine(driver):
		el13 = driver.find_element(by=AppiumBy.XPATH, value="//XCUIElementTypeStaticText[@name=\"X |\"]")
		sleep(3)
		el13.click()
		sleep(3)

#取得外資資料
class ForeignCapital:
	#外資 Foreign capital
	def twentiethDayNumber(driver):
		#點擊更多選項
		actions = ActionChains(driver)
		actions.w3c_actions = ActionBuilder(driver, mouse=PointerInput(interaction.POINTER_TOUCH, "touch"))
		actions.w3c_actions.pointer_action.move_to_location(348, 66)
		actions.w3c_actions.pointer_action.pointer_down()
		actions.w3c_actions.pointer_action.pause(0.1)
		actions.w3c_actions.pointer_action.release()
		actions.perform()
		
	
		#進入外資
		el2 = driver.find_element(by=AppiumBy.XPATH, value="//XCUIElementTypeApplication[@name=\"樹精靈(測)\"]/XCUIElementTypeWindow[1]/XCUIElementTypeOther[2]/XCUIElementTypeOther/XCUIElementTypeTable/XCUIElementTypeCell[9]/XCUIElementTypeOther[1]/XCUIElementTypeOther")
		sleep(3)
		el2.click()
		print("這是外資")
		list = GetNumber.getTwentiethDayNumber(driver)
		
		return list

#取得投信資料
class InvestmentTrust:
	#投信 Investment Trust
	def twentiethDayNumber(driver):
		#點擊更多選項
		enterOptions = driver.find_element(by=AppiumBy.XPATH, value="//XCUIElementTypeNavigationBar[@name=\"TWCTStockOverviewVC\"]/XCUIElementTypeButton[5]")
		enterOptions.click()
		#進入投信
		investmentTrust = driver.find_element(by=AppiumBy.XPATH, value="//XCUIElementTypeApplication[@name=\"樹精靈(測)\"]/XCUIElementTypeWindow[1]/XCUIElementTypeOther[2]/XCUIElementTypeOther/XCUIElementTypeTable/XCUIElementTypeCell[10]/XCUIElementTypeOther[1]/XCUIElementTypeOther")
		investmentTrust.click()
		print("這是投信")
		list =GetNumber.getTwentiethDayNumber(driver)
		
		return list

#取得自營
class Proprietary:
	#自營 Proprietary
	def twentiethDayNumber(driver):
		enterOptions = driver.find_element(by=AppiumBy.XPATH, value="//XCUIElementTypeNavigationBar[@name=\"TWCTStockOverviewVC\"]/XCUIElementTypeButton[5]")
		enterOptions.click()
		#進入自營
		proprietary = driver.find_element(by=AppiumBy.XPATH, value="//XCUIElementTypeApplication[@name=\"樹精靈(測)\"]/XCUIElementTypeWindow[1]/XCUIElementTypeOther[2]/XCUIElementTypeOther/XCUIElementTypeTable/XCUIElementTypeCell[11]/XCUIElementTypeOther[1]/XCUIElementTypeOther")
		proprietary.click()
		print("這是自營")
		list =GetNumber.getTwentiethDayNumber(driver)
		
		return list


class MarginTrading:
#融資 Margin Trading	
	def twentiethDayNumber(driver):
		enterOptions = driver.find_element(by=AppiumBy.XPATH, value="//XCUIElementTypeNavigationBar[@name=\"TWCTStockOverviewVC\"]/XCUIElementTypeButton[5]")
		enterOptions.click()	
		#進入融資
		marginTrading = driver.find_element(by=AppiumBy.XPATH, value="//XCUIElementTypeApplication[@name=\"樹精靈(測)\"]/XCUIElementTypeWindow[1]/XCUIElementTypeOther[2]/XCUIElementTypeOther/XCUIElementTypeTable/XCUIElementTypeCell[12]/XCUIElementTypeOther[1]/XCUIElementTypeOther")
		marginTrading.click()
		print("這是融資")
		list =GetNumber.getTwentiethDayNumber(driver)
		return list


class ShortSelling:
#融券 Short Selling
	def twentiethDayNumber(driver):
		
		#進入融券
		shortSelling = driver.find_element(by=AppiumBy.ACCESSIBILITY_ID, value=" 融券 ")
		shortSelling.click()
		print("這是融券")
		list = GetNumber.getTwentiethDayNumber(driver)
		return list

#券商(買)
class SecuritiesDealerToBuy:
	def twentiethDayNumber(driver):
		
		#進入融券
		el21 = driver.find_element(by=AppiumBy.ACCESSIBILITY_ID, value="券商(買)")
		el21.click()
		print("這是券商(買)")
		list = GetSecuritiesDealerNumber.first(driver)
		return list


#券商(賣)
class SecuritiesDealerToSell:
	def twentiethDayNumber(driver):
		#進入融券
		el22 = driver.find_element(by=AppiumBy.ACCESSIBILITY_ID, value="券商(賣)")
		el22.click()
		print("這是券商(賣)")
		list = GetSecuritiesDealerNumber.first(driver)
		return list

#讀取外資、投信...等表格數字		
class GetNumber:		
	def getTwentiethDayNumber(driver):
		#外資，，投信，自營，融資，融券近20日的 買張 賣張 買賣超
		try:
			lastTwentiethDay = driver.find_element(by=AppiumBy.ACCESSIBILITY_ID, value="近20日")
			print(lastTwentiethDay.text)
		except:
			print("查無資料")
			PrintExcel.error(driver)
			Error.error(driver)
		#try:
			#近20日 買張
		#buyLastTwentiethDay = driver.find_element(by=AppiumBy.XPATH, value="//XCUIElementTypeApplication[@name=\"樹精靈(測)\"]/XCUIElementTypeWindow[1]/XCUIElementTypeOther/XCUIElementTypeOther/XCUIElementTypeOther/XCUIElementTypeOther/XCUIElementTypeOther/XCUIElementTypeOther/XCUIElementTypeOther[3]/XCUIElementTypeOther/XCUIElementTypeOther/XCUIElementTypeOther/XCUIElementTypeOther[2]/XCUIElementTypeScrollView/XCUIElementTypeOther/XCUIElementTypeOther/XCUIElementTypeTable/XCUIElementTypeCell[1]")
			#注意不確定是否會抓到正確數值，無法抓元件判斷數值
		#print(buyLastTwentiethDay.text)
	
	def totalList(driver):
		print("測試資料：")
		totalList = [GetNumber.List(driver)]	
		return totalList
'''#近20日 賣張
			sellLastTwentiethDay = driver.find_element(by=AppiumBy.XPATH, value="/hierarchy/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.RelativeLayout/android.widget.ScrollView/android.widget.LinearLayout/androidx.viewpager.widget.ViewPager/android.widget.LinearLayout/android.widget.ListView/android.widget.LinearLayout[1]/android.widget.TextView[3]")
			print(sellLastTwentiethDay.text)
		
			#近20日 買賣超
			overBoughtAndSoldLastTwentiethDay = driver.find_element(by=AppiumBy.XPATH, value="/hierarchy/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.RelativeLayout/android.widget.ScrollView/android.widget.LinearLayout/androidx.viewpager.widget.ViewPager/android.widget.LinearLayout/android.widget.ListView/android.widget.LinearLayout[1]/android.widget.TextView[4]")	
			print(overBoughtAndSoldLastTwentiethDay.text)

			#寫入陣列,建議每個天數的"買張、賣張、買賣超"寫一個陣列
			twentiethDay=[buyLastTwentiethDay.text,sellLastTwentiethDay.text,overBoughtAndSoldLastTwentiethDay.text]

			#將 "-" 轉換成0
			for i in range(0,len(twentiethDay)):
				try:	
					if(twentiethDay[i]=="-"):
						twentiethDay[i]="0"
						#print("有非數字’-‘")
				except:
					print("")	

		except:
			print("沒有資料")	
			PrintExcel.error(driver)
			Error.error(driver)

		return twentiethDay '''



	


class GetSecuritiesDealerNumber:		
	def first(driver):
		#券商（買）近20日的 買張 賣張 買賣超
		#抓不到元素
		try:
			securitiesDealer= driver.find_element(by=AppiumBy.XPATH, value="(//XCUIElementTypeStaticText[@name=\"1\"])[5]")
			print(securitiesDealer.text)
		except Exception as e:
			print("查無資料")
			PrintExcel.error(driver)
			Error.error(driver)
		#try:
		sleep(5)
		#first = driver.find_element(by=AppiumBy.XPATH, value="//XCUIElementTypeApplication[@name=\"樹精靈(測)\"]/XCUIElementTypeWindow[1]/XCUIElementTypeOther/XCUIElementTypeOther/XCUIElementTypeOther/XCUIElementTypeOther/XCUIElementTypeOther/XCUIElementTypeOther/XCUIElementTypeOther[3]/XCUIElementTypeOther/XCUIElementTypeOther/XCUIElementTypeOther/XCUIElementTypeOther[2]/XCUIElementTypeScrollView/XCUIElementTypeOther/XCUIElementTypeOther/XCUIElementTypeTable/XCUIElementTypeCell[1]")
		#print(first.text)
	
	def totalList(driver):
		print("測試資料：")
		totalList = [GetSecuritiesDealerNumber.first(driver)]	
		return totalList	
'''			second = driver.find_element(by=AppiumBy.XPATH, value="/hierarchy/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.RelativeLayout/android.widget.ScrollView/android.widget.LinearLayout/androidx.viewpager.widget.ViewPager/android.widget.LinearLayout/android.widget.ListView/android.widget.LinearLayout[1]/android.widget.LinearLayout/android.widget.TextView[3]")
			print(second.text)#目前沒辦法抓元件名稱只能抓數字

			third = driver.find_element(by=AppiumBy.XPATH, value="/hierarchy/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.RelativeLayout/android.widget.ScrollView/android.widget.LinearLayout/androidx.viewpager.widget.ViewPager/android.widget.LinearLayout/android.widget.ListView/android.widget.LinearLayout[1]/android.widget.LinearLayout/android.widget.TextView[4]")
			print(third.text)#目前沒辦法抓元件名稱只能抓數字

			#寫入陣列,建議每個天數的"買張、賣張、買賣超"寫一個陣列
			firstData=[first.text,second.text,third.text]

			#將 "-" 轉換成0
			for i in range(0,len(firstData)):
				try:	
					if(firstData[i]=="-"):
						firstData[i]="0"
						#print("有非數字’-‘")
				except:
					print("")	
		except:
			print("沒有資料")	
			PrintExcel.error(driver)
			Error.error(driver)

		return firstData	'''


	#有其他要比較的天數要放進這
		
		

#比較外資、投信...等表格數字
class CompareNumber:
	def doList(driver,company):
		#載入儲存外資、投信 買入 賣出 買賣超 等數字的陣列
		foreignCapital = ForeignCapital.twentiethDayNumber(driver) 
		investmentTrust = InvestmentTrust.twentiethDayNumber(driver)
		proprietary = Proprietary.twentiethDayNumber(driver)
		marginTrading = MarginTrading.twentiethDayNumber(driver)
		shortSelling = ShortSelling.twentiethDayNumber(driver)
		securitiesDealerToBuy = SecuritiesDealerToBuy.twentiethDayNumber(driver)
		securitiesDealerToSell = SecuritiesDealerToSell.twentiethDayNumber(driver)

		#totalLiest = GetNumber.totalList(driver)
		securitiesDealerList = GetSecuritiesDealerNumber.totalList(driver)

		#比較 買入 賣出 買賣超 數字是否一樣
		buyAndSellList = ["買入","賣出","買賣超"] 
		day = ["近20天","昨天","4天前","5天前"]
		result = "正確"

		foreignCapitalAndInvestmentTrust = "外資、投信比較"
		investmentTrustAndProprietary = "投信、自營比較"
		marginTradingAndShortSelling = "融資、融券比較"
		securitiesDealerToBuyAndSecuritiesDealerToSell = "券商(買)、(賣)比較"
		#比較 買入 賣出 買賣超 數字是否一樣
'''		for i in range(len(securitiesDealerList)): #依total的數量跑幾次迴圈
			for i in range(len(buyAndSellList)): #每個交易日都要跑3次(買入、賣出、買賣超)

				#檢查外資、投信
				if(foreignCapital[i].__eq__(investmentTrust[i]) | int(foreignCapital[i])!=0 & int(investmentTrust[i])!=0):
					print(buyAndSellList[i]+":數字一樣，錯誤")
					result = "錯誤"
					PrintExcel.testNumber(driver,buyAndSellList[i],result,company,foreignCapitalAndInvestmentTrust)
					CompareNumberList.compareNumberList(driver,company,result,"","",foreignCapitalAndInvestmentTrust,buyAndSellList[i])
				else:
					print(buyAndSellList[i]+":數字正常")	
					PrintExcel.testNumber(driver,buyAndSellList[i],result,company,foreignCapitalAndInvestmentTrust)
					CompareNumberList.compareNumberList(driver,company,result,"","",foreignCapitalAndInvestmentTrust,buyAndSellList[i])
			
				#檢查投信、自營
				if(proprietary[i].__eq__(investmentTrust[i]) | int(proprietary[i])!=0 & int(investmentTrust[i])!=0):
					print(buyAndSellList[i]+":數字好像一樣，可能怪怪的")
					result = "錯誤"
					PrintExcel.testNumber(driver,buyAndSellList[i],result,company,investmentTrustAndProprietary)
					CompareNumberList.compareNumberList(driver,company,result,"","",foreignCapitalAndInvestmentTrust,buyAndSellList[i])
				else:
					print(buyAndSellList[i]+":數字正常")	
					PrintExcel.testNumber(driver,buyAndSellList[i],result,company,investmentTrustAndProprietary)	
					CompareNumberList.compareNumberList(driver,company,result,"","",foreignCapitalAndInvestmentTrust,buyAndSellList[i])

				#檢查融資、融券
				if(marginTrading[i].__eq__(shortSelling[i]) | int(marginTrading[i])!=0 & int(shortSelling[i])!=0):
					print(buyAndSellList[i]+":數字好像一樣，可能怪怪的")
					result = "錯誤"
					PrintExcel.testNumber(driver,buyAndSellList[i],result,company,marginTradingAndShortSelling)
					CompareNumberList.compareNumberList(driver,company,result,"","",marginTradingAndShortSelling,buyAndSellList[i])
				else:
					print(buyAndSellList[i]+":數字正常")	
					PrintExcel.testNumber(driver,buyAndSellList[i],result,company,marginTradingAndShortSelling)	
					CompareNumberList.compareNumberList(driver,company,result,"","",marginTradingAndShortSelling,buyAndSellList[i])


		for i in range(len(securitiesDealerList)): #跑幾次迴圈
			for i in range(len(buyAndSellList)): #每個交易日都要跑3次(買入、賣出、買賣超)

				#檢查券商(買)、(賣)
				if(securitiesDealerToBuy[i].__eq__(securitiesDealerToSell[i]) | int(securitiesDealerToBuy[i])!=0 & int(securitiesDealerToSell[i])!=0):
					print(buyAndSellList[i]+":數字好像一樣，可能怪怪的")
					result = "錯誤"
					PrintExcel.testNumber(driver,buyAndSellList[i],result,company,securitiesDealerToBuyAndSecuritiesDealerToSell)
					CompareNumberList.compareNumberList(driver,company,result,"","",securitiesDealerToBuyAndSecuritiesDealerToSell,buyAndSellList[i])
				else:
					print(buyAndSellList[i]+":數字正常")	
					PrintExcel.testNumber(driver,buyAndSellList[i],result,company,securitiesDealerToBuyAndSecuritiesDealerToSell)	
					CompareNumberList.compareNumberList(driver,company,result,"","",securitiesDealerToBuyAndSecuritiesDealerToSell,buyAndSellList[i])'''



#檢查各公司買進、賣出名稱
class ClickBuyAndSell:
	def clickBuy(driver):
		sleep(3)
		buy = driver.find_element(by=AppiumBy.XPATH, value="//XCUIElementTypeApplication[@name=\"樹精靈(測)\"]/XCUIElementTypeWindow[1]/XCUIElementTypeOther/XCUIElementTypeOther/XCUIElementTypeOther/XCUIElementTypeOther/XCUIElementTypeOther/XCUIElementTypeOther/XCUIElementTypeOther[1]/XCUIElementTypeOther/XCUIElementTypeOther/XCUIElementTypeOther/XCUIElementTypeOther/XCUIElementTypeScrollView/XCUIElementTypeOther/XCUIElementTypeOther/XCUIElementTypeOther/XCUIElementTypeOther[3]/XCUIElementTypeOther[3]/XCUIElementTypeButton")
		sleep(2)
		buy.click()

	def clickSell(driver):
		sleep(3)
		sell = driver.find_element(by=AppiumBy.XPATH, value="//XCUIElementTypeApplication[@name=\"樹精靈(測)\"]/XCUIElementTypeWindow[1]/XCUIElementTypeOther/XCUIElementTypeOther/XCUIElementTypeOther/XCUIElementTypeOther/XCUIElementTypeOther/XCUIElementTypeOther/XCUIElementTypeOther[1]/XCUIElementTypeOther/XCUIElementTypeOther/XCUIElementTypeOther/XCUIElementTypeOther/XCUIElementTypeScrollView/XCUIElementTypeOther/XCUIElementTypeOther/XCUIElementTypeOther/XCUIElementTypeOther[3]/XCUIElementTypeOther[5]/XCUIElementTypeButton")
		sleep(2)
		sell.click()

	def buyAndSellManual(driver):
		try:
			el8 = driver.find_element(by=AppiumBy.ACCESSIBILITY_ID, value="MBkResourcesImage Introduction")
			sleep(3)
			el8.click()
			sleep(3)
			el9 = driver.find_element(by=AppiumBy.ACCESSIBILITY_ID, value="MBkResourcesImage Introduction")
			sleep(3)
			el9.click()
			sleep(3)
			el10 = driver.find_element(by=AppiumBy.XPATH, value="//XCUIElementTypeStaticText[@name=\"暫不申請\"]")
			sleep(3)
			el10.click()
		except:
			print("沒有說明書")
			PrintExcel.error(driver)	
			Error.error(driver)
'''	def getName(driver):
		getName = driver.find_element(by=AppiumBy.ID, value="com.cathay.securities.mBroker:id/tv_simple_qoute_view_SymbolName")
		getName.click()
		setName = getName.text	
		return setName

	def clickName(driver,compareTest,buyOrSell):
		#將數字及空格刪掉
		
		testName = buyOrSell +" 公司名稱測試"
		companyName = compareTest.rstrip('0123456789').strip()
		print(companyName)
		result = "正確"
		if(ClickBuyAndSell.getName(driver) !=  companyName ):
			print(ClickBuyAndSell.getName(driver)+"名稱有誤")
			result = "錯誤"
			PrintExcel.clickBuyAndSell(driver,companyName,testName,result)
			BuyAndSell.buyAndSell(driver,companyName,result,"","",testName)

		else:	
			print(ClickBuyAndSell.getName(driver)+"名稱正確")		
			PrintExcel.clickBuyAndSell(driver,ClickBuyAndSell.getName(driver),testName,result)
			BuyAndSell.buyAndSell(driver,companyName,result,"","",testName)'''

	






		
		
		



		



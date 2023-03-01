from appium.webdriver.common.appiumby import AppiumBy
from excelTest import PrintExcel
from word import TestOddLet,TestKLine,CompareNumberList,BuyAndSell,Error

#from systex import SearchSystex

#確定名稱
class CompareName:
	def name(driver):
		driver.implicitly_wait(3)
		getName = driver.find_element(by=AppiumBy.XPATH, value="/hierarchy/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.RelativeLayout/androidx.viewpager.widget.ViewPager/android.widget.LinearLayout/android.view.ViewGroup/android.widget.LinearLayout/android.widget.TextView")
		setName = getName.text
		return setName 

#進入零股畫面
class OddLot:
	def oddLot(driver):
		enterOddLot = driver.find_element(by=AppiumBy.ID, value="com.cathay.securities.mBroker:id/btn_OddLot") #零股
		enterOddLot.click()
		return enterOddLot

#零股名稱
	def getOddLotName(driver):
		oddLot = driver.find_element(by=AppiumBy.ID, value="com.cathay.securities.mBroker:id/textView_name")
		oddLot.click
		oddLotName = oddLot.text
		return oddLotName

#檢查零股名稱
	def compareOddLotName(driver,companyName):
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
		return result

#離開畫面	
class Leave:
	def leave(driver):	
		#防止畫面跳轉太慢，強制睡眠3秒
		el1 = driver.find_element(by=AppiumBy.ACCESSIBILITY_ID, value="向上瀏覽")
		el1.click()
			


		
class KLine:
	#進入K線畫面
	def kLine(driver):
		el2 = driver.find_element(by=AppiumBy.ID, value="com.cathay.securities.mBroker:id/btn_TA")
		el2.click()

#偵測是否有說明書	有的話關閉
	def manual(driver):
		try:
			el1 = driver.find_element(by=AppiumBy.ID, value="com.cathay.securities.mBroker:id/ivBtn")
			el1.click()
			el2 = driver.find_element(by=AppiumBy.ID, value="com.cathay.securities.mBroker:id/ivBtn")
			el2.click()
		except:
			print("沒有說明書")		
			PrintExcel.error(driver)
			Error.error(driver)
			

#取得K線名稱
	def getKLineName(driver):
		getName = driver.find_element(by=AppiumBy.ID, value="com.cathay.securities.mBroker:id/tv_symbol_name") 
		try:
			getId = driver.find_element(by=AppiumBy.ID, value="com.cathay.securities.mBroker:id/tv_symbol_id") 
			kLineName = (getName.text +" "+ getId.text).strip()
		except:
			kLineName = getName.text
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
		leaveKLine = driver.find_element(by=AppiumBy.ID, value="com.cathay.securities.mBroker:id/imbtnCancel")
		leaveKLine.click()
		return

#取得外資資料
class ForeignCapital:
	#外資 Foreign capital
	def twentiethDayNumber(driver):
		#點擊更多選項
		enterOptions = driver.find_element(by=AppiumBy.ACCESSIBILITY_ID, value="更多選項")
		enterOptions.click()
		#進入外資
		foreignCapital = driver.find_element(by=AppiumBy.XPATH, value="/hierarchy/android.widget.FrameLayout/android.widget.FrameLayout/android.widget.ListView/android.widget.LinearLayout[9]/android.widget.LinearLayout/android.widget.RelativeLayout/android.widget.TextView")
		driver.implicitly_wait(3)
		foreignCapital.click()
		print("這是外資")
		list = GetNumber.getTwentiethDayNumber(driver)
		
		return list

#取得投信資料
class InvestmentTrust:
	#投信 Investment Trust
	def twentiethDayNumber(driver):
		#點擊更多選項
		enterOptions = driver.find_element(by=AppiumBy.ACCESSIBILITY_ID, value="更多選項")
		enterOptions.click()
		#進入投信
		investmentTrust = driver.find_element(by=AppiumBy.XPATH, value="/hierarchy/android.widget.FrameLayout/android.widget.FrameLayout/android.widget.ListView/android.widget.LinearLayout[10]/android.widget.LinearLayout/android.widget.RelativeLayout/android.widget.TextView")
		investmentTrust.click()
		print("這是投信")
		list =GetNumber.getTwentiethDayNumber(driver)
		
		return list

#取得自營
class Proprietary:
	#自營 Proprietary
	def twentiethDayNumber(driver):
		enterOptions = driver.find_element(by=AppiumBy.ACCESSIBILITY_ID, value="更多選項")
		enterOptions.click()
		#進入自營
		proprietary =driver.find_element(by=AppiumBy.XPATH, value="/hierarchy/android.widget.FrameLayout/android.widget.FrameLayout/android.widget.ListView/android.widget.LinearLayout[11]/android.widget.LinearLayout/android.widget.RelativeLayout/android.widget.TextView")
		proprietary.click()
		print("這是自營")
		list =GetNumber.getTwentiethDayNumber(driver)
		
		return list

#
class MarginTrading:
#融資 Margin Trading	
	def twentiethDayNumber(driver):
		enterOptions = driver.find_element(by=AppiumBy.ACCESSIBILITY_ID, value="更多選項")
		enterOptions.click()	
		#進入融資
		marginTrading =driver.find_element(by=AppiumBy.XPATH, value="/hierarchy/android.widget.FrameLayout/android.widget.FrameLayout/android.widget.ListView/android.widget.LinearLayout[12]/android.widget.LinearLayout/android.widget.RelativeLayout/android.widget.TextView")
		marginTrading.click()
		print("這是融資")
		list =GetNumber.getTwentiethDayNumber(driver)
		return list

#
class ShortSelling:
#融券 Short Selling
	def twentiethDayNumber(driver):
		enterOptions = driver.find_element(by=AppiumBy.ACCESSIBILITY_ID, value="更多選項")
		enterOptions.click()
		#進入融券
		shortSelling =driver.find_element(by=AppiumBy.XPATH, value="/hierarchy/android.widget.FrameLayout/android.widget.FrameLayout/android.widget.ListView/android.widget.LinearLayout[13]/android.widget.LinearLayout/android.widget.RelativeLayout/android.widget.TextView")
		shortSelling.click()
		print("這是融券")
		list =GetNumber.getTwentiethDayNumber(driver)
		return list

#券商(買)
class SecuritiesDealerToBuy:
	def twentiethDayNumber(driver):
		enterOptions = driver.find_element(by=AppiumBy.ACCESSIBILITY_ID, value="更多選項")
		enterOptions.click()
		el1 = driver.find_element(by=AppiumBy.XPATH, value="/hierarchy/android.widget.FrameLayout/android.widget.FrameLayout/android.widget.ListView/android.widget.LinearLayout[13]/android.widget.LinearLayout/android.widget.RelativeLayout/android.widget.TextView")
		el1.click()
		#進入融券
		el2 = driver.find_element(by=AppiumBy.XPATH, value="/hierarchy/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.RelativeLayout/android.widget.HorizontalScrollView/android.widget.LinearLayout/android.widget.LinearLayout[3]/android.widget.TextView")
		el2.click()
		print("這是券商(買)")
		list =GetSecuritiesDealerNumber.first(driver)
		return list


#券商(賣)
class SecuritiesDealerToSell:
	def twentiethDayNumber(driver):
		enterOptions = driver.find_element(by=AppiumBy.ACCESSIBILITY_ID, value="更多選項")
		enterOptions.click()
		el1 = driver.find_element(by=AppiumBy.XPATH, value="/hierarchy/android.widget.FrameLayout/android.widget.FrameLayout/android.widget.ListView/android.widget.LinearLayout[13]/android.widget.LinearLayout/android.widget.RelativeLayout/android.widget.TextView")
		el1.click()
		#進入融券
		el3 = driver.find_element(by=AppiumBy.XPATH, value="/hierarchy/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.RelativeLayout/android.widget.HorizontalScrollView/android.widget.LinearLayout/android.widget.LinearLayout[4]/android.widget.TextView")
		el3.click()
		print("這是券商(賣)")
		list =GetSecuritiesDealerNumber.first(driver)
		return list

#讀取外資、投信...等表格數字		
class GetNumber:		
	def getTwentiethDayNumber(driver):
		#外資近20日的 買張 賣張 買賣超
		try:
			lastTwentiethDay = driver.find_element(by=AppiumBy.XPATH, value="/hierarchy/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.RelativeLayout/android.widget.ScrollView/android.widget.LinearLayout/androidx.viewpager.widget.ViewPager/android.widget.LinearLayout/android.widget.ListView/android.widget.LinearLayout[1]/android.widget.TextView[1]")
			print(lastTwentiethDay.text)
		except:
			print("查無資料")
			PrintExcel.error(driver)
			Error.error(driver)
		try:
			#近20日 買張
			buyLastTwentiethDay = driver.find_element(by=AppiumBy.XPATH, value="/hierarchy/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.RelativeLayout/android.widget.ScrollView/android.widget.LinearLayout/androidx.viewpager.widget.ViewPager/android.widget.LinearLayout/android.widget.ListView/android.widget.LinearLayout[1]/android.widget.TextView[2]")
			print(buyLastTwentiethDay.text)
	
			#近20日 賣張
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

		return twentiethDay 



	def totalList(driver):
		print("測試資料：")
		totalList = [GetNumber.List(driver)]	
		return totalList


class GetSecuritiesDealerNumber:		
	def first(driver):
		#外資近20日的 買張 賣張 買賣超
		try:
			securitiesDealer= driver.find_element(by=AppiumBy.XPATH, value="/hierarchy/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.RelativeLayout/android.widget.ScrollView/android.widget.LinearLayout/androidx.viewpager.widget.ViewPager/android.widget.LinearLayout/android.widget.ListView/android.widget.LinearLayout[1]/android.widget.LinearLayout/android.widget.TextView[1]")
			print(securitiesDealer.text)
		except:
			print("查無資料")
			PrintExcel.error(driver)
			Error.error(driver)
		try:
			first = driver.find_element(by=AppiumBy.XPATH, value="/hierarchy/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.RelativeLayout/android.widget.ScrollView/android.widget.LinearLayout/androidx.viewpager.widget.ViewPager/android.widget.LinearLayout/android.widget.ListView/android.widget.LinearLayout[1]/android.widget.LinearLayout/android.widget.TextView[2]")
			print(first.text)

			second = driver.find_element(by=AppiumBy.XPATH, value="/hierarchy/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.RelativeLayout/android.widget.ScrollView/android.widget.LinearLayout/androidx.viewpager.widget.ViewPager/android.widget.LinearLayout/android.widget.ListView/android.widget.LinearLayout[1]/android.widget.LinearLayout/android.widget.TextView[3]")
			print(second.text)

			third = driver.find_element(by=AppiumBy.XPATH, value="/hierarchy/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.RelativeLayout/android.widget.ScrollView/android.widget.LinearLayout/androidx.viewpager.widget.ViewPager/android.widget.LinearLayout/android.widget.ListView/android.widget.LinearLayout[1]/android.widget.LinearLayout/android.widget.TextView[4]")
			print(third.text)

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

		return firstData	


	#有其他要比較的天數要放進這
	def totalList(driver):
		print("測試資料：")
		totalList = [GetSecuritiesDealerNumber.first(driver)]	
		return totalList	
		

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
		for i in range(len(securitiesDealerList)): #依total的數量跑幾次迴圈
			for i in range(len(buyAndSellList)): #每個交易日都要跑3次(買入、賣出、買賣超)

				#檢查外資、投信
				if(foreignCapital[i].__eq__(investmentTrust[i]) | int(foreignCapital[i])!=0 & int(investmentTrust[i])!=0):
					print(buyAndSellList[i]+":數字好像一樣，可能怪怪的")
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
					CompareNumberList.compareNumberList(driver,company,result,"","",securitiesDealerToBuyAndSecuritiesDealerToSell,buyAndSellList[i])



#檢查各公司買進、賣出名稱
class ClickBuyAndSell:
	def clickBuy(driver):
		buy = driver.find_element(by=AppiumBy.ID, value="com.cathay.securities.mBroker:id/BuyText") 
		buy.click()

	def clickSell(driver):
		sell = driver.find_element(by=AppiumBy.ID, value="com.cathay.securities.mBroker:id/btn_order_sell")
		sell.click()	

	def getName(driver):
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
			BuyAndSell.buyAndSell(driver,companyName,result,"","",testName)

	def buyAndSellManual(driver):
		try:
			manualNo = driver.find_element(by=AppiumBy.ID, value="com.cathay.securities.mBroker:id/button_No")
			manualNo.click()
			manualNext = driver.find_element(by=AppiumBy.ID, value="com.cathay.securities.mBroker:id/ivBtn")
			manualNext.click()
			manualNextTwo = driver.find_element(by=AppiumBy.ID, value="com.cathay.securities.mBroker:id/ivBtn")
			manualNextTwo.click()
		except:
			print("沒有說明書")
			PrintExcel.error(driver)	
			Error.error(driver)






		
		
		



		



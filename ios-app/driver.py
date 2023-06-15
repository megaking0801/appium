from appium import webdriver

class Driver:
	def __init__():
		desired_caps = {
    		'platformName': 'iOS',
    		'platformVersion': '15.4.1', # 實際設備的版本號
    		'deviceName': 'SM iphoneX', # 實際設備名稱
    		'udid': '3dd9ee3b57268ad708380b6ae98348975fd5a5a0', # 實際設備的UDID
		    "app":"/Users/21ms20r/.jenkins/workspace/iOS_MBroker_Cathay/build/樹精靈(測).app",
    		'bundleId': 'com.softmobile.developer.MBroker.CathayTW.test', # 應用程序包ID
    		'automationName': 'XCUITest', # 使用XCUITest框架
    		'xcodeOrgId': '95E44WWFM7', # 開發人員團隊ID
    		'xcodeSigningId': 'iPhone Developer' # 開發人員證書名稱
		}
		return webdriver.Remote("http://127.0.0.1:4723/wd/hub", desired_caps)
	



'''
desired_caps = {
			"appium:deviceName" : "iPhone 8",
			#"appium:udid": "A9B90166-14C4-4579-9657-3BEA78B7C73E",
			"platformName" : "iOS",
			"appium:platformVersion" : "16.2",
			"allowTestPackages": "true",#加了這個3.3才能直接包到手機
			"app":"/Users/21ms20r/.jenkins/workspace/iOS_MBroker_Cathay/build/樹精靈(測).app",
			"appium:automationName": "XCUITest",
			"bundleId":"com.softmobile.developer.MBroker.CathayTW.test",
			"noReset":"true",
			"autoAcceptAlerts":"true"
			#"wdaStartupRetries": "4",
    		#'webdriveragenturl': 'http://localhost:8100'
		}'''
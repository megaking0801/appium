from appium import webdriver

class Driver2:
	def __init__():
		desired_caps = {
    		'platformName': 'iOS',
    		'platformVersion': '15.1', # 實際設備的版本號
    		'deviceName': '21MS20R的iPhone', # 實際設備名稱
    		'udid': '002328fd3bc39081fa480980116ecf8728a8ea95', # 實際設備的UDID
		    "app":"/Users/21ms20r/.jenkins/workspace/iOS_MBroker_Cathay/build/樹精靈(測).app",
    		'bundleId': 'com.softmobile.developer.MBroker.CathayTW.test', # 應用程序包ID
    		'automationName': 'XCUITest', # 使用XCUITest框架
    		'xcodeOrgId': '95E44WWFM7', # 開發人員團隊ID
    		'xcodeSigningId': 'iPhone Developer' # 開發人員證書名稱
		}
		return webdriver.Remote("http://127.0.0.1:4723/wd/hub", desired_caps)
	
	
	
	'''
		    'platformName': 'iOS',
    		'platformVersion': '15.1', # 實際設備的版本號
    		'deviceName': 'SM 6S', # 實際設備名稱
    		'udid': '002328fd3bc39081fa480980116ecf8728a8ea95', # 實際設備的UDID
		    "app":"/Users/21ms20r/.jenkins/workspace/iOS_MBroker_Cathay/build/樹精靈(測).app",
    		'bundleId': 'com.softmobile.developer.MBroker.CathayTW.test', # 應用程序包ID
    		'automationName': 'XCUITest', # 使用XCUITest框架
    		'xcodeOrgId': '95E44WWFM7', # 開發人員團隊ID
    		'xcodeSigningId': 'iPhone Developer' # 開發人員證書名稱
		    
			'platformName': 'iOS',
		    'automationName': 'XCUITest', # 使用XCUITest框架
    		# 'platformVersion': '16.4', # 實際設備的版本號
    		'deviceName': 'iPhone 14 Pro Max', # 實際設備名稱
    		#'udid': '95978FF2-5E7F-42B1-9F03-BB19DB982DE7',# 實際設備的UDID
			"app":"/Users/21ms20r/.jenkins/workspace/iOS_MBroker_Cathay/build/樹精靈(測).app",
			'bundleId': 'com.softmobile.developer.MBroker.CathayTW.test',
            "noReset" : "true",
            "autoAcceptAlerts": "true"
			
		    
		    
		    'platformName': 'iOS',
    		'platformVersion': '16.4', # 實際設備的版本號
    		'deviceName': 'iPhone 14 Pro Max', # 實際設備名稱
    		'udid': '95978FF2-5E7F-42B1-9F03-BB19DB982DE7',# 實際設備的UDID
			'useNewWDA':True,
		    'wdaLaunchTimeout':60000 
	'''
from appium import webdriver

class Driver:
	def __init__():
		desired_caps = {
			"appium:deviceName" : "2817d3d29904",
			"platformName" : "Android",
			"appium:platformVersion" : "7",
			"allowTestPackages": "true",#加了這個3.3才能直接包到手機
			"app":"/Users/linshaoqun/.jenkins/workspace/iwow_Andoid_MBroker_Cathay/MBroker_Cathay/build/outputs/apk/release/MBroker_Cathay-release-unsigned.apk"
		}
		return webdriver.Remote("http://127.0.0.1:4723/wd/hub", desired_caps)

#"appium:appPackage": "com.cathay.securities.mBroker",
#"appium:appActivity": "cathay.activity.Start.StartActivity"
#"app":"/Users/linshaoqun/.jenkins/workspace/1/MBroker_Cathay/build/outputs/apk/release/MBroker_Cathay-release-unsigned.apk"
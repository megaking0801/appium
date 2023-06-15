from appium import webdriver

class Driver2:
	def __init__():
		desired_caps = {
			"appium:deviceName": "iPhone 14 Pro Max",
            "appium:udid": "1437756B-33CA-489A-AFF5-0896379307ED",
            "platformName": "iOS",
            "appium:app": "/Users/linshaoqun/Library/Developer/Xcode/DerivedData/MBroker_Cathay-cjmmgulzpplldwhctlwiykyjmrof/Build/Products/Release-iphonesimulator/1.app",
            "appium:automationName": "XCUITest",
            "appium:platformVersion": "16.2"
		}
		return webdriver.Remote("http://127.0.0.1:4723/wd/hub", desired_caps)
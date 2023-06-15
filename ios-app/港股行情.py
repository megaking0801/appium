from appium.webdriver.common.appiumby import AppiumBy
from time import sleep
from docx import Document
from appium.webdriver.common.mobileby import MobileBy
from appium.webdriver.common.touch_action import TouchAction
from selenium.webdriver.common.action_chains import ActionChains,ActionBuilder
from selenium.webdriver.common.actions.pointer_input import PointerInput
from selenium.webdriver.common.actions import interaction
from docx import Document
from openpyxl import load_workbook
from docx.shared import Cm,Pt
from datetime import datetime
from openpyxl import Workbook

class h_stock:
    def __init__(driver):

        el1 = driver.find_element(by=AppiumBy.XPATH, value="//XCUIElementTypeButton[@name=\"自選庫存\"]")
        sleep(3)
        el1.click()

        actions = ActionChains(driver)
        actions.w3c_actions = ActionBuilder(driver, mouse=PointerInput(interaction.POINTER_TOUCH, "touch"))
        actions.w3c_actions.pointer_action.move_to_location(236, 555)
        actions.w3c_actions.pointer_action.pointer_down()
        actions.w3c_actions.pointer_action.pause(0.1)
        actions.w3c_actions.pointer_action.release()
        actions.perform()
    
        actions = ActionChains(driver)
        actions.w3c_actions = ActionBuilder(driver, mouse=PointerInput(interaction.POINTER_TOUCH, "touch"))
        actions.w3c_actions.pointer_action.move_to_location(25, 392)
        actions.w3c_actions.pointer_action.pointer_down()
        actions.w3c_actions.pointer_action.move_to_location(25, 642)
        actions.w3c_actions.pointer_action.release()
        actions.perform()
    
        el3 = driver.find_element(by=AppiumBy.ACCESSIBILITY_ID, value="港指成分")
        el3.click()
        el4 = driver.find_element(by=AppiumBy.ACCESSIBILITY_ID, value="恆生指數")
        c1 = el4.text
        if(c1 == "恆生指數"):
            print("ok")
        else:
            print('bad')
        
        el5 = driver.find_element(by=AppiumBy.XPATH, value="//XCUIElementTypeButton[@name=\"港指成分(延遲)\"]")
        el5.click()

        el6 = driver.find_element(by=AppiumBy.ACCESSIBILITY_ID, value="港股焦點")
        el6.click()
        el7 = driver.find_element(by=AppiumBy.ACCESSIBILITY_ID, value="藍籌股")
        c2 = el7.text
        if(c2 == "藍籌股"):
            print("ok")
        else:
            print('bad')

        el9 = driver.find_element(by=AppiumBy.XPATH, value="//XCUIElementTypeButton[@name=\"港股焦點(延遲)\"]")
        el9.click()

        el12 = driver.find_element(by=AppiumBy.ACCESSIBILITY_ID, value="港股報價")
        el12.click()
        sleep(3)
        el13 = driver.find_element(by=AppiumBy.ACCESSIBILITY_ID, value="能源類")
        el13.click()
        sleep(3)
        el14 = driver.find_element(by=AppiumBy.ACCESSIBILITY_ID, value="能源類")
        c3 = el14.text
        if(c3 == "能源類"):
            print("ok")
        else:
            print('bad')

        el15 = driver.find_element(by=AppiumBy.XPATH, value="//XCUIElementTypeButton[@name=\"港股報價(延遲)\"]")
        el15.click()
        el15 = driver.find_element(by=AppiumBy.XPATH, value="//XCUIElementTypeButton[@name=\"港股報價(延遲)\"]")
        el15.click()

        el17 = driver.find_element(by=AppiumBy.ACCESSIBILITY_ID, value="港股排行")
        el17.click()
        el18 = driver.find_element(by=AppiumBy.ACCESSIBILITY_ID, value="成交量排行")
        c4 = el18.text

        if(c4 == "成交量排行"):
            print("成交量排行ok")
        else:
            print('成交量排行bad')
        
        el19 = driver.find_element(by=AppiumBy.XPATH, value="//XCUIElementTypeButton[@name=\"港股排行(延遲)\"]")
        el19.click()
        el20 = driver.find_element(by=AppiumBy.ACCESSIBILITY_ID, value="自選")
        el20.click()
        print("離開港股行情測試")

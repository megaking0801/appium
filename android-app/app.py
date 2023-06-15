from driver import Driver
from systex import SearchSystex
from login import Login
from TurnOffScreenShot import TurnOffScreenShot
from systex import SearchSystex 
from tsmc import SearchTsmc
from aapl import SearchAapl
from excelTest import PrintExcel
from word import PrintWord


if __name__=='__main__':

		driver = Driver.__init__()		
		

		Login.__init__(driver)

		PrintExcel.title(driver)
		
		PrintWord.title(driver)
		
		TurnOffScreenShot.__init__(driver)

		SearchSystex.__init__(driver)
		
		SearchTsmc.__init__(driver)
		
		SearchAapl.__init__(driver)
		

		print("結束測試")
		
		
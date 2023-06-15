from ios import Driver2
from ioslogin import ioslogin
from systex import SearchSystex

from tsmc import SearchTsmc
from aapl import SearchAapl
#from excelTest import PrintExcel
from word import PrintWord
from 證券行情 import Securities
from 美股行情 import a_stock
from 港股行情 import h_stock

if __name__=='__main__':
		
        driver = Driver2.__init__()
        
        ioslogin.__init__(driver)
        
        SearchSystex.__init__(driver)
        
        SearchTsmc.__init__(driver)
        
        SearchAapl.__init__(driver)
        
        Securities.__init__(driver)

        a_stock.__init__(driver)
        
        #h_stock.__init__(driver)

        print("結束測試")
		
		
		
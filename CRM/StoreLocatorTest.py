import sys
SertaFolder = 'D:\AutomationTest\Scripts\SertaTA'
if SertaFolder not in sys.path:
    sys.path.append(SertaFolder)
from Util import Util

import unittest, time
from selenium import webdriver
import openpyxl as pyxl
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from CrmUrls import CrmUrls
from CrmLoginTest import CrmLoginTest

class StoreLocatorFilter(object):
    def __init__(self, zip, radius, proCat):
        self.zip = zip
        self.radius = radius
        if (self.radius is None):
            self.radius = 10
        self.proCat = proCat
        if (self.zip is None) :
            raise Exception('Zip code is required field')

class StoreLocatorTest(unittest.TestCase):
    def setUp(self):
        self.driver = webdriver.Firefox()
        self.driver.get(CrmUrls.LOGIN_URL)
        CrmLoginTest.testLoginCorrectAccount(self)

    def readCriteria(self, fileName):
        workBook = pyxl.load_workbook(fileName)
        sheet = workBook.get_sheet_by_name('Criteria')
        arrFilter = []
        for row in range(2, sheet.max_row + 1):
            rowNo = str(row)
            print('*** Filter ', row - 1, ': ***')
            # Zip code
            zipCode = sheet['A' + rowNo].value
            print(' - Zip code: ', zipCode)
            # Radius
            radius = sheet['B' + rowNo].value
            print(' - Lastname: ', radius)
            # Product category
            proCategory = sheet['C' + rowNo].value
            print(' - Product category: ', proCategory)
            storeLocatorFilter = StoreLocatorFilter(zipCode, radius, proCategory)
            arrFilter.append(storeLocatorFilter)
        return arrFilter


    def testSearchStoreLocator(self):
        self.driver.get(CrmUrls.STORE_LOCATOR_URL)
        self.assertEqual(self.driver.title.lower(), 'serta store - search retailer')
        fileName = 'D:\AutomationTest\Scripts\SertaTA\DataTest\StoreLocatorCriteria.xlsx'
        arrFilter = self.readCriteria(fileName)
        for storeLocatorCriteria in arrFilter:
            Util.fillTextByEleId(self.driver, 'retailer-search-zip', storeLocatorCriteria.zip)
            Select(self.driver.find_element_by_id('retailer-search-radius')).select_by_value(str(storeLocatorCriteria.radius))
            if (storeLocatorCriteria.proCat is not None) :
                Select(self.driver.find_element_by_id('retailers-filter')).select_by_value(storeLocatorCriteria.proCat)
            self.driver.find_element_by_class_name('btn').click()
            # Sleep time allows tester check result before move to next criteria
            time.sleep(30)

    def tearDown(self):
        self.driver.close()

if __name__ == '__main__':
    # Testcases are ordered by alphabet
    suite = unittest.TestSuite()
    suite.addTest(StoreLocatorTest('testSearchStoreLocator'))
    unittest.TextTestRunner(verbosity=2).run(suite)

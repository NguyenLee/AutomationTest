import unittest
import os
from selenium import webdriver
from selenium.webdriver.support.ui import Select
import openpyxl as pyxl


class SaveTaxScreenshot(unittest.TestCase):
    url = 'https://portal.vertexsmb.com/'
    lookupTaxUrl = 'https://portal.vertexsmb.com/ratelookup/'

    def setUp(self):
        # self.driver = webdriver.Chrome('C:\Program Files (x86)\Google\Chrome\Application')
        self.driver = webdriver.Firefox()
        self.driver.get(self.url)

    def fillText(self, eleName, data):
        element = self.driver.find_element_by_name(eleName)
        element.clear()
        element.send_keys(data)

    def login(self):
        self.fillText('username', '')
        self.fillText('password', '')
        self.driver.find_element_by_xpath("//button[@type='submit']").click()
        self.assertEqual('Dashboard | Vertex SMB', self.driver.title)

    def readAddresses(self):
        filename = 'D:\AutomationTest\Scripts\SertaTA\RealAddresses.xlsx'
        workBook = pyxl.load_workbook(filename)
        sheet = workBook.get_sheet_by_name('SERTA')
        arrAddresses = []

        for row in range(2, sheet.max_row + 1):
            print('*** Address ', row - 1, ': ***')
            # Address 1
            address1 = sheet['A' + str(row)].value
            print('	- Address 1: ', address1)
            # Address 2
            address2 = sheet['B' + str(row)].value
            print('	- Address 2: ', address2)
            # City
            city = sheet['C' + str(row)].value
            print('	- City: ', city)
            # State code
            stateCode = sheet['D' + str(row)].value
            print('	- State code: ', stateCode)
            # State
            state = sheet['E' + str(row)].value
            print('	- State: ', state)
            # Zip code
            zip = sheet['F' + str(row)].value
            print('	- Zip: ', zip)
            addressObj = Address(address1, address2, city, stateCode, state, zip)
            arrAddresses.append(addressObj)
        return arrAddresses

    def lookupTaxRate(self):
        self.driver.get(self.lookupTaxUrl)
        arrAddresses = self.readAddresses()
        for addressObj in arrAddresses:
            self.fillText('RateLookupRequest.Address.StreetAddress1', addressObj.address1)
            self.fillText('RateLookupRequest.Address.City', addressObj.city)
            # Select state
            states = Select(self.driver.find_element_by_id('statesDropDown'))
            states.select_by_value(addressObj.stateCode)
            self.fillText('RateLookupRequest.Address.PostalCode', addressObj.zip)
            self.driver.find_element_by_id('lookup-rate').click()

            screenshotFolder = 'C:\Tagrem\TestScreenshots\VertexTA\\'
            filePath = screenshotFolder + addressObj.state
            # Create state folder if it doesn't have
            if not os.path.exists(filePath):
                os.makedirs(filePath)
            fileName = filePath + '\\' + addressObj.stateCode + '_TaxRate.png'

            if self.driver.save_screenshot(fileName):
                print('Save screenshot' + fileName + ' successfully')
            else:
                print('Save screenshot' + fileName + ' failed')

    def saveTaxScreenshots(self):
        self.login()
        self.lookupTaxRate()

    def tearDown(self):
        self.driver.close()


if __name__ == '__main__':
    import sys

    SertaFolder = 'D:\AutomationTest\Scripts\SertaTA'
    if SertaFolder not in sys.path:
        sys.path.append(SertaFolder)
    from Util import Util
    from Address import Address

    suite = unittest.TestSuite();
    suite.addTest(SaveTaxScreenshot('saveTaxScreenshots'));
    unittest.TextTestRunner(verbosity=2).run(suite)

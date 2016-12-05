import sys
SertaFolder = 'D:\AutomationTest\Scripts\SertaTA'
if SertaFolder not in sys.path:
    sys.path.append(SertaFolder)
from Util import Util
from Address import Address
import unittest, time
from selenium import webdriver
import openpyxl as pyxl
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from CrmCustomer import CrmCustomer
from CrmLoginTest import CrmLoginTest
from CrmUrls import CrmUrls

class CrmCustomerTest(unittest.TestCase) :
    def setUp(self):
        self.driver = webdriver.Firefox()
        self.driver.get(CrmUrls.LOGIN_URL)
        CrmLoginTest.testLoginCorrectAccount(self)

    def testSearchCustomerEmptyCriterion(self):
        self.driver.get(CrmUrls.SEARCH_CUSTOMER_URL)
        self.driver.find_element_by_name('search_customer').click()
        assert 'At least 1 field(s) are required for performing the search' in self.driver.page_source

    def searchCustomerByEmail(self, email, isGP = False):
        if isGP :
            self.driver.get(CrmUrls.GP_SEARCH_CUSTOMER_URL)
        else:
            self.driver.get(CrmUrls.SEARCH_CUSTOMER_URL)
        Util.fillText(self.driver, 'CustomerContact[email_address]', email)
        self.driver.find_element_by_name('search_customer').click()

    def searchCustomers(self, orderNo = '', firstname = '', lastName = '', email='', address=''):
        print('Define later')

    def loadCustomer(self, pos = 0):
        try :
            selectLinks = WebDriverWait(self.driver, 20).until(EC.presence_of_all_elements_located((By.LINK_TEXT, 'Select')))
            # Active customer is customer at position = pos in array
            selectLinks[pos].click()
            self.assertEqual(self.driver.title, 'Serta Store - Communication Default')
        except TimeoutException:
            print('No customer')

    def checkActiveCustomer(self):
        return Util.is_element_present(self.driver, By.CLASS_NAME, 'customer_contact')

    def endCustomerSession(self):
        if CrmCustomerTest.checkActiveCustomer(self) :
            self.driver.find_element_by_class_name('customer_end').click()
            WebDriverWait(self.driver, 20).until(EC.visibility_of_element_located((By.ID, 'content')))
            #endCustomerPopup = self.driver.find_element_by_id('content')
            #print(endCustomerPopup)
            self.driver.find_element_by_name('callpopup[call_end]').click()

    def testSearchCustomerByEmail(self):
        CrmCustomerTest.searchCustomerByEmail(self, 'robert_waller@indition.net')
        CrmCustomerTest.loadCustomer(self)

    def readData(self, fileName, sheetName):
        workBook = pyxl.load_workbook(fileName)
        sheet = workBook.get_sheet_by_name(sheetName)
        arrCustomers = []
        for row in range(2, sheet.max_row + 1):
            rowNo = str(row)
            print('*** Customer ', row - 1, ': ***')
            # Firstname
            firstname = sheet['A' + rowNo].value
            print(' - Firstname: ', firstname)
            # Lastname
            lastname = sheet['B' + rowNo].value
            print(' - Lastname: ', lastname)

            # Address 1
            print(' - Address: ')
            address1 = sheet['C' + rowNo].value
            print('	    + Address 1: ', address1)
            # Address 2
            address2 = sheet['D' + rowNo].value
            print('	    + Address 2: ', address2)
            # City
            city = sheet['E' + rowNo].value
            print('	    + City: ', city)
            # State code
            stateCode = sheet['F' + rowNo].value
            print('	    + State code: ', stateCode)
            # State
            state = sheet['G' + rowNo].value
            print('	    + State: ', state)
            # Zip code
            zip = sheet['H' + rowNo].value
            print('	    + Zip: ', zip)
            addressObj = Address(address1, address2, city, stateCode, state, zip)
            # Address type
            addressType = sheet['I' + rowNo].value
            print('- Address type: ', addressType)

            # Email
            email = sheet['J' + rowNo].value
            print('- Email: ', email)
            # Email type
            emailType = sheet['K' + rowNo].value
            print('- Email type:', emailType)

            # Phone
            phone = sheet['L' + rowNo].value
            print('- Phone: ', phone)
            # Phone type
            phoneType = sheet['M' + rowNo].value
            print('- Phone type:', phoneType)

            customer = CrmCustomer(firstname, lastname, addressObj, addressType, email, emailType, phone, phoneType)
            arrCustomers.append(customer)
        return arrCustomers

    def newCustomer(self, customerObj, useVerifiedAddress = True):
        if not isinstance(customerObj, CrmCustomer):
            print('Incorrect address')
            return False;

        self.driver.get(CrmUrls.NEW_CUSTOMER_URL)
        try:
            WebDriverWait(self.driver, 30).until(EC.visibility_of_element_located((By.ID, 'customer-contact-form')))
            #Firstname
            firstNameTextbox = self.driver.find_element_by_xpath("//form[@id='customer-contact-form']//input[@id='CustomerContact_firstname']")
            Util.fillText(self.driver, firstNameTextbox, customerObj.firstname)

            #Lastname
            lastNameTextbox = self.driver.find_element_by_xpath("//form[@id='customer-contact-form']//input[@id='CustomerContact_lastname']")
            Util.fillText(self.driver, lastNameTextbox, customerObj.lastname)

            # Address
            delay = 15
            WebDriverWait(self.driver, delay).until(EC.visibility_of_element_located((By.ID, 'addressDetails')))
            Select(self.driver.find_element_by_id('CustomerContactAddress_addresstype')).select_by_value(str(customerObj.addressType))
            Util.fillTextByEleId(self.driver, 'CustomerContactAddress_address1', customerObj.addressObj.address1)
            if not (customerObj.addressObj.address2 is None):
                Util.fillTextByEleId(self.driver, 'CustomerContactAddress_address2', customerObj.addressObj.address2)
            Util.fillTextByEleId(self.driver, 'CustomerContactAddress_city', customerObj.addressObj.city)
            Select(self.driver.find_element_by_id('CustomerContactAddress_state')).select_by_visible_text(customerObj.addressObj.stateCode)
            Util.fillTextByEleId(self.driver, 'CustomerContactAddress_postalcode', customerObj.addressObj.zip)

            #Email
            if not (customerObj.email is None):
                emailTypeXpath = "//div[@id='emailDetails']//div[@class='custom']//p//select[@id='CustomerContactEmail_emailtype']"
                WebDriverWait(self.driver, delay).until(EC.visibility_of_element_located((By.XPATH, emailTypeXpath)))
                Select(self.driver.find_element_by_xpath(emailTypeXpath)).select_by_value('77')
                emailXpath = "//div[@id='emailDetails']//div[@class='custom']//p[@class='cc_email']//input[@id='CustomerContactEmail_email']"
                emailEle = self.driver.find_element_by_xpath(emailXpath)
                Util.fillTextByEleId(self.driver, emailEle, customerObj.email)

            #Phone
            if not (customerObj.phone is None):
                phoneTypeXpath = "//div[@id='phoneDetails']//div[@class='custom']//p//select[@id='CustomerContactPhone_phonetype']"
                WebDriverWait(self.driver, delay).until(EC.visibility_of_element_located((By.XPATH, phoneTypeXpath)))
                Select(self.driver.find_element_by_xpath(phoneTypeXpath)).select_by_value(str(customerObj.phoneType))
                phoneXpath = "//div[@id='phoneDetails']//div[@class='custom']//p//input[@id='CustomerContactPhone_phone']"
                phoneEle = self.driver.find_element_by_xpath(phoneXpath)
                Util.fillTextByEleId(self.driver, phoneEle, customerObj.phone)

            self.driver.find_element_by_name('save_contact').click()
            time.sleep(delay + 10)
            verifyAddressPopup = self.driver.find_element_by_xpath("//div[@id='modal']//div[@id='content']")
            #print(verifyAddressPopup.is_enabled())

            # Handle verify address popup
            cssButtonClass = 'use-address-as-entered'
            if (useVerifiedAddress is True):
                cssButtonClass = 'use-address-as-verified'
            buttonXpath = "//div[@id='modal']//div[@id='content']//div//div//input[@class='" + cssButtonClass + "']"
            #print(buttonXpath)
            buttonEle = self.driver.find_element_by_xpath(buttonXpath)
            #print(buttonEle.is_displayed())
            buttonEle.click()
            WebDriverWait(self.driver, delay).until(lambda driver: driver.current_url != CrmUrls.NEW_CUSTOMER_URL)

            # Check email of created customer
            result = CrmCustomerTest.checkCreatedCustomerEmail(self, customerObj.email)
            if result :
                print('Customer is created with correct email')
            else :
                print('Customer email is incorrect')

        except TimeoutException :
            print(TimeoutException)

    def testNewCustomer(self):
        customerFile = 'D:\AutomationTest\Scripts\SertaTA\DataTest\Customers.xlsx'
        customers = self.readData(customerFile, 'Test')
        for customer in customers:
            self.newCustomer(customer)
            self.endCustomerSession()

    def checkCreatedCustomerEmail(self, cusEmail = None):
        self.driver.find_element_by_id('ui-id-5').click()
        emailXPath = "//div[@id='emailDetails']/div[@class='custom']/p[@class='cc_email']/span"
        try:
            WebDriverWait(self.driver, 15).until(EC.presence_of_element_located((By.XPATH, emailXPath)))
        except TimeoutException as e:
            print(e)
        emailEle = self.driver.find_element_by_xpath(emailXPath)
        createdEmail = emailEle.text
        print('Customer email: ', createdEmail)
        # Customer created with @noemail.com
        result = False
        if (cusEmail is None) :
            if 'contact' in createdEmail :
                if '@noemail.com' in createdEmail:
                    result = True
        else:
            result = (createdEmail == cusEmail)
        print(result)
        return result

    def testCheckCreatedCustomer(self):
        cusEmail = 'veronica_nguyen@indition.net'
        CrmCustomerTest.searchCustomerByEmail(self, cusEmail)
        CrmCustomerTest.loadCustomer(self)
        CrmCustomerTest.checkCreatedCustomerEmail(self, cusEmail)

    def testEndCustomerSession(self):
        cusEmail = 'veronica_nguyen@indition.net'
        CrmCustomerTest.searchCustomerByEmail(self, cusEmail)
        CrmCustomerTest.loadCustomer(self)
        CrmCustomerTest.endCustomerSession(self)

    def tearDown(self):
        self.driver.close()

if __name__ == '__main__':
    # Testcases are ordered by alphabet
    suite = unittest.TestSuite()
    suite.addTest(CrmCustomerTest('testNewCustomer'))
    #suite.addTest(CrmCustomerTest('testEndCustomerSession'))
    #suite.addTest(CrmCustomerTest('testSearchCustomerEmptyCriterion'))
    #suite.addTest(CrmCustomerTest('testSearchCustomerByEmail'))
    #suite.addTest(CrmCustomerTest('testCheckCreatedCustomer'))
    unittest.TextTestRunner(verbosity=2).run(suite)

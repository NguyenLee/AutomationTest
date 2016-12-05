import sys
SertaFolder = 'D:\AutomationTest\Scripts\SertaTA'
if SertaFolder not in sys.path:
    sys.path.append(SertaFolder)
from Util import Util
from Address import Address
import unittest, time, os
from selenium import webdriver
import openpyxl as pyxl
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.action_chains import ActionChains
from CrmLoginTest import CrmLoginTest
from CrmCustomerTest import CrmCustomerTest
from CrmUrls import CrmUrls
from BrowseProductsTest import BrowseProductsTest

class CheckoutTest(unittest.TestCase):
    def setUp(self):
        self.driver = webdriver.Firefox()
        self.driver.get(CrmUrls.LOGIN_URL)
        CrmLoginTest.testLoginCorrectAccount(self)

    def checkoutStep1(self, firstname, lastname, customerEmail, addressObj, phone, useVerifiedAddress = True):
        # Go to checkout page
        self.driver.get(CrmUrls.CHECKOUT_STEP1_URL)
        delay = 10
        time.sleep(2)
        # This customer has items in shopping cart
        if (self.driver.title != 'Serta Store - Step1 CrmCheckout'):
            print(self.driver.title)
            # Add items to cart
            BrowseProductsTest.testAddItemToCart(self)

        # Fill checkout info
        Util.fillTextByEleId(self.driver, 'BillingForm_firstname', firstname)
        Util.fillTextByEleId(self.driver, 'ShippingForm_firstname', firstname)

        Util.fillTextByEleId(self.driver, 'BillingForm_lastname', lastname)
        Util.fillTextByEleId(self.driver, 'ShippingForm_lastname', lastname)

        Util.fillText(self.driver, 'BillingForm[address_1]', addressObj.address1)
        Util.fillText(self.driver, 'ShippingForm[address_1]', addressObj.address1)
        if (addressObj.address2 is not None):
            Util.fillText(self.driver, 'BillingForm[address_2]', addressObj.address2)
            Util.fillText(self.driver, 'ShippingForm[address_2]', addressObj.address2)

        Util.fillText(self.driver, 'BillingForm[city]', addressObj.city)
        Util.fillText(self.driver, 'ShippingForm[city]', addressObj.city)
        Select(self.driver.find_element_by_name('BillingForm[state]')).select_by_value(addressObj.stateCode)
        Select(self.driver.find_element_by_name('ShippingForm[state]')).select_by_value(addressObj.stateCode)
        Util.fillText(self.driver, 'BillingForm[postcode]', addressObj.zip)
        Util.fillText(self.driver, 'ShippingForm[postcode]', addressObj.zip)
        Util.fillText(self.driver, 'BillingForm[email]', customerEmail)
        Util.fillText(self.driver, 'ShippingForm[email]', customerEmail)
        Util.fillText(self.driver, 'BillingForm[phone]', phone)
        Util.fillText(self.driver, 'ShippingForm[phone]', phone)
        self.driver.find_element_by_id('submit-step1').click()


        # Handle verified address popup if it's openned
        #step1_window_handle = self.driver.current_window_handle
        self.driver.find_element_by_id('submit-step1').click()
        time.sleep(20)
        if (self.driver.current_url == CrmUrls.CHECKOUT_STEP1_URL):
            try :
                xpathAddressPopup = "//div[@id='address_modal']"
                verifiedAddressPopupEle = self.driver.find_element_by_xpath(xpathAddressPopup)
                if(verifiedAddressPopupEle.is_displayed()) :
                    buttonCssClassName = 'use-address-as-entered'
                    if (useVerifiedAddress):
                        buttonCssClassName = 'use-address-as-verified'
                    addressOptionButton = self.driver.find_element_by_class_name(buttonCssClassName)
                    #print(addressOptionButton)
                    #print(addressOptionButton.is_displayed())
                    # Must use js to click the address option because the button encounters invisible status everytime even using WebDriverWait
                    self.driver.execute_script("$(arguments[0]).click();", addressOptionButton)
            except Exception as ex:
                print(ex)

        WebDriverWait(self.driver, 15).until(lambda driver: driver.current_url != CrmUrls.CHECKOUT_STEP1_URL)


    def getTaxScreenshot(self, firstName, lastName, customerEmail, addressObj, phone):
        # Checkout
        if not isinstance(addressObj, Address):
            print('Incorrect address')
            print(addressObj)
            return;

        # Checkout step 1
        self.checkoutStep1(firstName, lastName, customerEmail, addressObj, phone)
        WebDriverWait(self.driver, 10).until(lambda driver: driver.current_url != CrmUrls.CHECKOUT_STEP1_URL)

        # Go to checkout step 2
        self.assertEqual(self.driver.title, 'Serta Store - Step2 CrmCheckout')

        # Get tax screenshot
        screenshotFolder = 'C:\Tagrem\TestScreenshots\VertexTA\\'
        filePath = screenshotFolder + addressObj.state
        # Create state folder if it doesn't have
        if not os.path.exists(filePath):
            os.makedirs(filePath)
        fileName = filePath + '\\' + addressObj.stateCode + '_CRM_TaxToolTip.png'

        # Hover Tax to see tax details
        taxEle = self.driver.find_element_by_css_selector(".cart-info-tax-row .i-info")
        hoverAction = ActionChains(self.driver).move_to_element(taxEle)
        hoverAction.perform()
        time.sleep(7)

        if self.driver.save_screenshot(fileName):
            print('Save screenshot ' + fileName + ' successfully')
        else:
            print('Save screenshot ' + fileName + ' failed')

    def checkoutWithExistingCreditCard(self):
        # Checkout step 2
        self.assertEqual(self.driver.current_url, CrmUrls.CHECKOUT_STEP2_URL)
        try:
            existingCreditCardsXpath = "//table[@class='table-payment']"
            # This customer has an exising payment
            WebDriverWait(self.driver, 20).until(EC.visibility_of_element_located((By.XPATH, existingCreditCardsXpath)))
            # Get first credit card payment
            creditXPath = "//table[@class='table-payment']//input[@class='use-gateway']"
            print(self.driver.find_element_by_xpath(creditXPath))
            if (Util.is_element_present(self.driver, By.XPATH, creditXPath)) :
                self.driver.find_element_by_xpath(creditXPath).click()
                self.driver.find_element_by_id('submit-step2').click()
                WebDriverWait(self.driver, 15).until(lambda driver: driver.current_url !=
                                                                    CrmUrls.CHECKOUT_STEP2_URL)
            else :
                print('Cannot find element')
        except TimeoutException :
            print("This customer doesn't have existing credit card")

    def checkoutNewCreditCard(self, cardNumber, securityCode, expiredMon, expiredYear):
        # Checkout step 2
        self.assertEqual(self.driver.current_url, CrmUrls.CHECKOUT_STEP2_URL)
        if (Util.is_element_present(self.driver, By.CSS_SELECTOR, 'section.new-card')) :
            # This customer has ICC card
            if (Util.is_element_present(self.driver, By.ID, 'alliance-payment')) :
                self.driver.find_element_by_id('chkbox-use-icomfort-credit-card').click()

            Util.fillText(self.driver, 'PaytraceDetails[cardnumber]', cardNumber)
            Util.fillText(self.driver, 'PaytraceDetails[cardnumber_confirm]', cardNumber)
            Util.fillText(self.driver, 'PaytraceDetails[securitycode]', securityCode)
            Util.fillText(self.driver, 'PaytraceDetails[expiremonth]', expiredMon)
            Util.fillText(self.driver, 'PaytraceDetails[expireyear]', expiredYear)
            self.driver.find_element_by_id('submit-step2').click()
            WebDriverWait(self.driver, 20).until(lambda driver: driver.current_url != CrmUrls.CHECKOUT_STEP2_URL)
            self.assertEqual(self.driver.title, 'Serta Store - View OrderDetail')
        else :
            print('Missing checkout with new card')


    def readAddresses(self, fileName, sheetName):
        workBook = pyxl.load_workbook(fileName)
        sheet = workBook.get_sheet_by_name(sheetName)
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

    def tearDown(self):
        self.driver.close()
        #print('Tear down')

    def testVerifyAddress(self):
        # Load existing customer
        customerEmail = 'marco_polo@indition.net'
        CrmCustomerTest.searchCustomerByEmail(self, customerEmail)
        CrmCustomerTest.loadCustomer(self)

        addressObj = Address('8210 Byron ave', '24', 'Miami beach', 'FL', 'Florida', '33141')
        self.checkoutStep1('Marco', 'Polo', customerEmail, addressObj, '0909090909')

    def takeCrmCheckoutScreenshot(self):
        # Load existing customer
        CrmCustomerTest.searchCustomerByEmail(self, 'robert_waller@indition.net')
        CrmCustomerTest.loadCustomer(self)

        # addressObj = Address('6716 Quail Ridge Dr', '', 'Montgomery', 'AL', 'Alabama', '36117')
        fileName = 'D:\AutomationTest\Scripts\SertaTA\RealAddresses.xlsx'
        arrAddresses = self.readAddresses(fileName, 'CRM')
        for addressObj in arrAddresses:
            self.getTaxScreenshot('Robert', 'Waller', 'robert_waller@indition.net', addressObj, '0909090909')

    def testCrmCheckoutExistingCC(self):
        customerEmail = 'gloria_mitchell@indition.net'
        # Load existing customer
        CrmCustomerTest.searchCustomerByEmail(self, customerEmail)
        CrmCustomerTest.loadCustomer(self)

        #BrowseProductsTest.browseProductByName(self, 'Guidance', '35', '249')
        #BrowseProductsTest.browseProductByCategory(self, '552')
        addressObj = Address('911 Paradrome St', '', 'Cicinatti', 'OH', 'Ohio', '45202')
        self.checkoutStep1('Gloria', 'Mitchell', customerEmail, addressObj, '1234567890')
        self.checkoutWithExistingCreditCard()
        '''
        fileName = 'D:\AutomationTest\Scripts\SertaTA\RealAddresses.xlsx'
        arrAddresses = self.readAddresses(fileName, 'Order addresses')
        for addressObj in arrAddresses:
            # Browse products
            #BrowseProductsTest.browseProductByName(self, 'Guidance', '35', '249')
            BrowseProductsTest.browseProductByCategory(self, '552')
            self.checkoutStep1('Gloria', 'Mitchell', customerEmail, addressObj, '1234567890')
            self.checkoutExistingCart()
        '''

    def testCrmCheckoutNewCC(self):
        customerEmail = 'adaline_bowman@indition.net'
        # Load existing customer
        CrmCustomerTest.searchCustomerByEmail(self, customerEmail)
        CrmCustomerTest.loadCustomer(self)

        #BrowseProductsTest.browseProductByName(self, 'Guidance', '35', '249')
        #BrowseProductsTest.browseProductByCategory(self, '558')
        addressObj = Address('477 SMITH STREET', '', 'Providence', 'RI', 'Rhode Island', '02904')
        self.checkoutStep1('Adaline', 'Bowman', customerEmail, addressObj, '1234567890')
        self.checkoutNewCreditCard('5481167872840156', '1818', '08', '2020')

if __name__ == '__main__':
    # Testcases are ordered by alphabet
    suite = unittest.TestSuite()
    #suite.addTest(CheckoutTest('takeCrmCheckoutScreenshot'))
    #suite.addTest(CheckoutTest('testVerifyAddress'))
    #suite.addTest(CheckoutTest('testCrmCheckoutExistingCC'))
    suite.addTest(CheckoutTest('testCrmCheckoutNewCC'))
    unittest.TextTestRunner(verbosity=2).run(suite)

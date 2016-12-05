import sys
SertaFolder = 'D:\AutomationTest\Scripts\SertaTA'
if SertaFolder not in sys.path:
    sys.path.append(SertaFolder)
from Util import Util

import unittest, time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from CrmLoginTest import CrmLoginTest
from CrmCustomerTest import CrmCustomerTest
from CrmUrls import CrmUrls

class BrowseProductsTest(unittest.TestCase):
    def setUp(self):
        self.driver = webdriver.Firefox()
        self.driver.get(CrmUrls.LOGIN_URL)
        CrmLoginTest.testLoginCorrectAccount(self)

    def checkActiveCustomer(self):
        if (Util.is_element_present(self.driver, By.CLASS_NAME, 'customer_contact')) :
            return True
        return False

    def browseProductByName(self, productName, size, foundation = None, addItemToCart = True):
        self.driver.get(CrmUrls.BROWSE_PRODUCTS_URL)
        self.assertEqual('Serta Store - Update CrmCart', self.driver.title)

        # Filter products by name
        xpath = "//table[@id='product-container']//tbody//tr//td[@class='colae300']//input"
        nameFilterEle = self.driver.find_element_by_xpath(xpath)
        if nameFilterEle is None :
            return NoSuchElementException
        Util.fillText(self.driver, nameFilterEle, productName)
        nameFilterEle.send_keys(Keys.ENTER)

        delay = 15
        # Wait for search result
        time.sleep(delay)
        # This is a mattress, so, it has foundation
        if (foundation is not None) :
            # Mattress size drop down list
            WebDriverWait(self.driver, delay).until(EC.presence_of_element_located((By.ID, 'Mattress_Size')))
            Select(self.driver.find_element_by_id('Mattress_Size')).select_by_value(size)

            # Foundation drop down list
            WebDriverWait(self.driver, delay).until(EC.presence_of_element_located((By.ID, 'Foundation')))
            Select(self.driver.find_element_by_id('Foundation')).select_by_value(foundation)
        # This is an accesssory
        else:
            WebDriverWait(self.driver, delay).until(EC.presence_of_element_located((By.ID, 'Size')))
            Select(self.driver.find_element_by_id('Size')).select_by_value(size)

        # Add item to cart
        if (addItemToCart) :
            BrowseProductsTest.addItemToCart(self)

    # Function filter product by category & get first product found
    def browseProductByCategory(self, categoryId, addItemToCart = True):
        self.driver.get(CrmUrls.BROWSE_PRODUCTS_URL)
        self.assertEqual('Serta Store - Update CrmCart', self.driver.title)

        # Filter products by name
        xpath = "//table[@id='product-container']//tbody//tr//td[@class='colae160']//select[@id='category']"
        categoryDdl = Select(self.driver.find_element_by_id('category'))
        if categoryDdl is None:
            return NoSuchElementException
        categoryDdl.select_by_value(categoryId)

        delay = 15
        # Wait for search result
        time.sleep(delay)
        try :
            WebDriverWait(self.driver, delay).until(EC.presence_of_element_located((By.ID, 'Mattress_Size')))
            # Select 3rd mattress size
            Select(self.driver.find_element_by_id('Mattress_Size')).select_by_index(0)
            # Select 1st foundation
            WebDriverWait(self.driver, delay).until(EC.presence_of_element_located((By.ID, 'Foundation')))
            Select(self.driver.find_element_by_id('Foundation')).select_by_index(0)
        except :
            # First result is an accessory
            WebDriverWait(self.driver, delay).until(EC.presence_of_element_located((By.ID, 'Size')))
            # Get price of 2nd item
            Select(self.driver.find_element_by_id('Size')).select_by_index(1)

        # Add item to cart
        if (addItemToCart):
            BrowseProductsTest.addItemToCart(self)


    def addItemToCart(self):
        WebDriverWait(self.driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'button.add-to-cart')))
        self.driver.find_element_by_css_selector("button.add-to-cart").click()

    def testAddItemToCart(self):
        # No customer is active, load an existing customer
        if not BrowseProductsTest.checkActiveCustomer(self) :
            CrmCustomerTest.searchCustomerByEmail(self, 'gloria@indition.net')
            CrmCustomerTest.loadCustomer(self)
        BrowseProductsTest.browseProductByName(self, 'Guidance', '35', '249')
        BrowseProductsTest.browseProductByCategory(self, '552')

    def testBrowseProduct(self):
        self.browseProductByName('Guidance', '35', '249', False)


if __name__ == '__main__':
    # Testcases are ordered by alphabet
    suite = unittest.TestSuite()
    #suite.addTest(BrowseProductsTest('testBrowseProduct'))
    suite.addTest(BrowseProductsTest('testAddItemToCart'))
    unittest.TextTestRunner(verbosity=2).run(suite)
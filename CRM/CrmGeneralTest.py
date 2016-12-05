import sys
SertaFolder = 'D:\AutomationTest\Scripts\SertaTA'
if SertaFolder not in sys.path:
    sys.path.append(SertaFolder)
from Util import Util
from Address import Address
import unittest, time, os
from selenium import webdriver
import openpyxl as pyxl
import os, logging
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from CrmLoginTest import CrmLoginTest
from CrmCustomerTest import CrmCustomerTest
from CrmUrls import CrmUrls

class CrmGeneralTest(unittest.TestCase) :
    def setUp(self):
        self.driver = webdriver.Firefox()
        self.driver.get(CrmUrls.LOGIN_URL)
        CrmLoginTest.testLoginCorrectAccount(self)

    def testMenuEleLinks(self, eleNo):
        itemMenuEles = self.driver.find_elements_by_xpath("//div[@class='menu']/ul/li[" + str(eleNo) + "]/ul/li/a")
        links = []
        for itemEle in itemMenuEles:
            menuLink = itemEle.get_attribute('href')
            #print(itemEle.get_attribute('text'), ': ', menuLink)
            links.append(menuLink)

        LOG_FILENAME = 'D:\AutomationTest\Scripts\SertaTA\CrmLogging.txt'
        logging.basicConfig(filename=LOG_FILENAME, level=logging.ERROR)
        for link in links:
            result = 'OK'
            try :
                print('*** Link: ', link)
                print('Page title: ', self.driver.title)
                self.driver.get(link)
                time.sleep(15)
                if (self.driver.title.lower().find('php warning') > -1 or self.driver.title.lower().find('exception') > -1 or self.driver.title.lower().find('error') > -1) :
                    msg = self.driver.find_element_by_class_name('message').get_attribute('innerHTML')
                    logging.error(link + ': ' + msg)
                    result = msg
            except Warning as w:
                print(w)
            except Exception as e:
                logging.exception("message")
            print(result)

    def testListLinks(self):
        self.testMenuEleLinks(4)
        self.testMenuEleLinks(5)
        self.testMenuEleLinks(6)

    def testReports(self):
        self.driver.get(CrmUrls.REPORT_URL)
        reportEles = self.driver.find_elements_by_xpath("//div[@id='mid_cont']/div/div/div/ul/li/a")
        reportLinks = []
        for reportEle in reportEles :
            reportLinks.append(reportEle.get_attribute('href'))

        LOG_FILENAME = 'D:\AutomationTest\Scripts\SertaTA\CrmLogging.txt'
        logging.basicConfig(filename=LOG_FILENAME, level=logging.ERROR)
        for link in reportLinks :
            result='OK'
            try :
                print('*** Link: ', link)
                print('Page title: ', self.driver.title)
                self.driver.get(link)
                time.sleep(10)
                if (self.driver.title.lower().find('php warning') > -1 or self.driver.title.lower().find('exception') > -1 or self.driver.title.lower().find('error') > -1) :
                    msg = self.driver.find_element_by_class_name('message').get_attribute('innerHTML')
                    logging.error(link + ': ' + msg)
                    result = msg
            except Exception as e:
                print(e)
                logging.exception("exception")
            except Warning as w:
                print(w)
                logging.warning("warning")
            print(result)


    def tearDown(self):
        self.driver.close()

if __name__ == '__main__':
    # Testcases are ordered by alphabet
    suite = unittest.TestSuite()
    suite.addTest(CrmGeneralTest('testListLinks'))
    #suite.addTest(CrmGeneralTest('testReports'))
    unittest.TextTestRunner(verbosity=2).run(suite)
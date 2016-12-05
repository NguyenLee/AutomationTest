import sys
SertaFolder = 'D:\AutomationTest\Scripts\SertaTA'
if SertaFolder not in sys.path:
    sys.path.append(SertaFolder)
from Util import Util
import unittest
from selenium import webdriver
from CrmUrls import CrmUrls

class CrmLoginTest(unittest.TestCase):
    def setUp(self):
        self.driver = webdriver.Firefox()
        self.driver.get(CrmUrls.LOGIN_URL)

    def testLoginEmptyUsername(self):
        print('**** Test Login function with empty username ****', end='\n')
        Util.fillText(self.driver, 'LoginForm[login]', '')
        assert 'Login cannot be blank' in self.driver.page_source
        Util.informMsg(self.driver, 'yt0', 'Invalid username or password')

    def testLoginEmptyPassword(self):
        print('**** Test Login function with empty password ****', end='\n')
        Util.fillText(self.driver, 'LoginForm[login]', 'username')
        Util.fillText(self.driver, 'LoginForm[password]', '')
        assert 'Password cannot be blank'
        Util.informMsg(self.driver, 'yt0', 'Invalid username or password')

    def testLoginInvalidAccount(self):
        print('**** Test Login function with invalid account ****', end='\n')
        username = 'username'
        password = 'password'
        print('Username: ', username, ' Password: ', password)
        Util.fillText(self.driver, 'LoginForm[login]', username)
        Util.fillText(self.driver, 'LoginForm[password]', password)
        Util.informMsg(self.driver, 'yt0', 'Invalid username or password')

    def testLoginCorrectAccount(self):
        print('**** Test Login function with invalid account ****', end='\n')
        username = 'nguyen@indition.net'
        password = '12345'
        print('Username: ', username, ' Password: ', password)
        Util.fillText(self.driver, 'LoginForm[login]', username)
        Util.fillText(self.driver, 'LoginForm[password]', password)
        self.driver.find_element_by_name('yt0').click()
        return self.assertEqual(self.driver.title, 'Serta Store - Dashboard')

    def tearDown(self):
        self.driver.close()

if __name__ == '__main__':
    # Add test suite to order test cases
    suite = unittest.TestSuite()
    #suite.addTest(CrmLoginTest('testLoginEmptyUsername'))
    #suite.addTest(CrmLoginTest('testLoginEmptyPassword'))
    #suite.addTest(CrmLoginTest('testLoginInvalidAccount'))
    suite.addTest(CrmLoginTest('testLoginCorrectAccount'))
    unittest.TextTestRunner(verbosity=2).run(suite)

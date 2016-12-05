from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException

class Util:
    def fillText(driver, ele, data):
        # Pass an element name, find element
        if (isinstance(ele, str)) :
            element = driver.find_element_by_name(ele)
        else :
            element = ele
        element.clear()
        element.send_keys(data)

    def fillTextByEleId(driver, eleId, data):
        # Pass an element id, find element
        if (isinstance(eleId, str)) :
            element = driver.find_element_by_id(eleId)
        else :
            element = eleId
        element.clear()
        element.send_keys(data)

    def informMsg(driver, eleName, msg):
        element = driver.find_element_by_name(eleName)
        element.send_keys(Keys.ENTER)
        assert msg in driver.page_source

    def is_element_present(driver, how, what):
        try: driver.find_element(by=how, value=what)
        except NoSuchElementException as e: return False
        return True
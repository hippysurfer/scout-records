# coding: utf-8
from splinter  import Browser
import os


prefs = {"browser.download.folderList": 2,
            "browser.download.manager.showWhenStarting": False,
            "browser.download.dir": os.getcwd(),
            "browser.helperApps.neverAsk.saveToDisk": "application/octet-stream"}

#browser = Browser('firefox', profile_preferences=prefs)
browser = Browser('phantomjs')

browser.visit('https://compass.scouts.org.uk/login/User/Login')
browser.fill('EM', '')
browser.fill('PW', '')
browser.find_by_value('Submit').first.click()

browser.is_element_present_by_name('ctl00$UserTitleMenu$cboUCRoles', wait_time=30)
browser.select('ctl00$UserTitleMenu$cboUCRoles','1253644')

browser.is_text_present('My Scouting', wait_time=30)
browser.click_link_by_text('My Scouting')

def wait_then_click_xpath(xpath, wait_time=30):
    browser.is_element_present_by_xpath(xpath, wait_time=wait_time)
    browser.find_by_xpath(xpath).click()
    
wait_then_click_xpath('//*[@id="TR_HIER7"]/h2')
wait_then_click_xpath('//*[@id="TR_HIER7_TBL"]/tbody/tr[7]/td[4]/a')
wait_then_click_xpath('//*[@id="bnExport"]')
wait_then_click_xpath('//*[@id="tbl_hdv"]/div/table/tbody/tr[2]/td[2]/input')
wait_then_click_xpath('//*[@id="bnOK"]')
wait_then_click_xpath('//*[@id="bnAlertOK"]')


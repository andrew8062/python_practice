from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
import logging
import ConfigParser
import logging
import File as f
from getpass import getpass
import time

class Agile:
	def __init__(self, username, password):
		self.userid = username
		self.password= password

	def click_item_by_id(self, id):
		elem = self.driver.find_element_by_id(id)
		elem.click()

	def download_issue_list(self):
		#initialization
		self.driver = webdriver.Chrome()
		self.driver.get("https://agile.us.dell.com/Agile/default/login-cms.jsp")
		assert "Agile" in self.driver.title

		#enter username and password
		elem = self.driver.find_element_by_name("j_username")
		elem.clear()
		elem.send_keys(self.userid)
		elem = self.driver.find_element_by_name("j_password")
		elem.clear()
		elem.send_keys(self.password)

		#click login button
		self.click_item_by_id( "login")
		("click log-in")

		#switch window
		logger.info("current handle: "+self.driver.current_window_handle)
		self.driver.switch_to_window(self.driver.window_handles[-1])
		logger.info("current handle: "+self.driver.current_window_handle)

		#in case user need to reset password
		#refresh to skip the reset process
		if "Reset Login" in self.driver.title:
			self.driver.refresh()
		#wait page load
		logger.info("wait until element")
		wait = WebDriverWait(self.driver, 60)
		element = wait.until(EC.element_to_be_clickable((By.ID,'MSG_Show_In_Navigator_121span')))

		#click personal search
		logger.info("click personal search")
		self.click_item_by_id( "ygtvt4")
		
		element = wait.until(EC.element_to_be_clickable((By.ID,'ygtvlabelel115')))

		#click report
		logger.info("click Loki AMD")
		self.click_item_by_id( "ygtvlabelel115")

		#wait for issue querying
		wait = WebDriverWait(self.driver, 180)
		element = wait.until(EC.element_to_be_clickable((By.ID,'More_110span')))

		#click more
		logger.info ("click more button")
		self.click_item_by_id( "More_110span")

		#click download report
		logger.info("click export as xls")
		self.click_item_by_id( "yui-gen41")

		#sleep 5 seconds to complete download
		time.sleep(5)
		#safely exit webdriver

	def quit(self):
		self.driver.quit()


		
if __name__ == "__main__":
	
	logging.basicConfig(level = logging.INFO)
	logger = logging.getLogger(__name__)
	logger.info('start')
	config = ConfigParser.ConfigParser()
	config.read('setting')
	proj_name = config.get('Agile', 'proj_name')
	username = config.get('Agile', 'username')
	download_path = config.get('Agile', 'download_path')
	store_path = config.get('Agile', 'store_path')
	
	password = getpass("Password: ")
	agile = Agile(username, password)
	agile.download_issue_list()
	# #rename downloaded issue list to "{proj_name}_issue_{timestamp}.xls"	
	f.rename_issue_list(proj_name, download_path, store_path)
	#Start comparing 2 most recent issue list
	f.compare_excels(proj_name, store_path)
	agile.quit()


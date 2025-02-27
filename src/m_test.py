from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options

chrome_options = Options()
chrome_options.add_argument("--headless")  # Run in headless mode
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")

service = Service(r"C:\webdrivers\chromedriver.exe")  # Update the path
driver = webdriver.Chrome(service=service, options=chrome_options)

driver.get("https://web.whatsapp.com")
print(driver.title)

driver.quit()

import os
import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from collate import combine, combine_all, formatted_combine, rename, rename_all, saveas_all

# Create a Service object with the path to the ChromeDriver executable
service = Service('C:\\Users\\chromedriver_win32\\chromedriver.exe')
download_dir = 'C:\\Users\\HP\\Downloads\\dgft-downloads'

def run_scrape(service, download_dir):
    options = webdriver.ChromeOptions()
    prefs = {'download.default_directory': f"{download_dir}\\2_2",
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
            }
    options.add_experimental_option('prefs', prefs)
    driver = webdriver.Chrome(service=service, chrome_options=options)

    for i in range(100):
        # Navigate to the URL
        driver.get('https://tradestat.commerce.gov.in/eidb/ecomq.asp')

        # Find the input fields and fill them in
        input_field_1 = driver.find_element(By.NAME, 'hscode')
        ip = str(i)
        if len(ip) < 2:
            ip = '0' + ip
        input_field_1.send_keys(ip)

        # Find the button to navigate to the next page and click it
        next_button = driver.find_element(By.ID, 'button1')
        next_button.click()

        # Wait for the page to load and find the button to download the data
        download_button = WebDriverWait(driver, 1000).until(EC.presence_of_element_located((By.ID, 'button1')))
        download_button.click()

        time.sleep(3)

    # Close the web driver
    driver.quit()

    # Download all 4/6/8 digits HS code
    for type in []:
        # Create a ChromeDriver instance with the Service object
        options = webdriver.ChromeOptions()
        prefs = {'download.default_directory': f"{download_dir}\\{type}",
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
                "safebrowsing.enabled": True
                }
        options.add_experimental_option('prefs', prefs)
        driver = webdriver.Chrome(service=service, chrome_options=options)

        for i in range(1, 100):
            try:
                prev_no_of_files = len(os.listdir(download_dir))

                # Navigate to the URL
                driver.get('https://tradestat.commerce.gov.in/eidb/ecomq.asp')

                # Find the input fields and fill them in
                input_field_1 = driver.find_element(By.NAME, 'hscode')
                ip = str(i)
                if len(ip) < 2:
                    ip = '0' + ip
                input_field_1.send_keys(ip)

                # Find the button to navigate to the next page and click it
                next_button = driver.find_element(By.ID, 'button1')
                next_button.click()

                # Wait for the page to load and find the button to download the data
                type_button = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, f'//a[@href="ecom{type}.asp?hs={ip}"]')))
                type_button.click()

                download_button = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'button1')))
                download_button.click()

                time.sleep(1)

                # wait for the download to complete
                # wait = WebDriverWait(driver, 30)
                # wait.until(lambda driver: len(os.listdir(download_dir)) > prev_no_of_files)

                # # get the path of the downloaded file
                # filename = max([download_dir + '\\' + f for f in os.listdir(download_dir)], key=os.path.getctime)

                # # rename the file
                # new_filename = f'eidb{type}_{ip}.xls'
                # os.rename(filename, download_dir + '\\' + new_filename)

            except Exception as e:
                print("".join(str(e)))

        # Close the web driver
        driver.quit()

def run_combine():
    for type in [2,4,6,8]:
        # combine all files into 1
        # rename_all(f"{download_dir}\\{type}")
        # combine_all(f"{download_dir}\\{type}", type)
        # saveas_all(f"{download_dir}\\{type}")
        formatted_combine(f"{download_dir}\\{type}", type)

run_combine()
# run_scrape(service, download_dir)

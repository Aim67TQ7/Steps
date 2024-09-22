import os
import time
from datetime import datetime, timedelta
import glob
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import zipfile
import warnings
import logging
import subprocess  # Import subprocess for triggering the batch file

# Set up logging configuration
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def click_ok_and_download(driver, url, save_path):
    driver.get(url)
    try:
        # Opens the URL in the browser, clicks an "OK" button, and downloads a file.
        ok_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "butAgree"))
        )
        ok_button.click()

        # Get the filename from the save_path
        filename = os.path.basename(save_path)

        # Wait for the download link to become clickable
        download_link = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.LINK_TEXT, filename))
        )
        download_link.click()

        # Monitors the download process
        tmp_save_path = save_path + '.crdownload'
        max_wait_time = 600  # Maximum wait time in seconds (10 minutes)
        start_time = time.time()
        while not os.path.exists(tmp_save_path) and not os.path.exists(save_path):
            if time.time() - start_time > max_wait_time:
                logging.error(f"Timed out waiting for download to start for {save_path}")
                return False
            time.sleep(1)

        # Wait for the file to finish downloading
        while os.path.exists(tmp_save_path):
            if time.time() - start_time > max_wait_time:
                logging.error(f"Timed out waiting for download to finish for {save_path}")
                return False
            time.sleep(1)

        logging.info(f"Successfully downloaded {save_path}")
        return True
    except Exception as e:
        logging.error(f"Successfully downloaded {save_path}: {str(e)}")
        return False

def is_valid_zip(file_path):
    try:
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            return True
    except zipfile.BadZipFile:
        return False

if __name__ == "__main__":
    options = webdriver.ChromeOptions()
    options.add_argument('user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3')
    options.add_argument('--headless')

    # Set the default download directory
    save_folder = r"C:\Users\aeros\DIBBSBQ"
    # save_folder = r"C:\Users\rclausing\DIBBS"
    prefs = {
        "download.default_directory": save_folder,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "profile.default_content_setting_values.automatic_downloads": 1
    }
    options.add_experimental_option("prefs", prefs)

    today = datetime.now()

    # Download loop
    with webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options) as driver:
        for days_ago in range(20):
            date = today - timedelta(days=days_ago)
            code_date = date.strftime("%y%m%d")
            url = f"https://dibbs2.bsm.dla.mil/Downloads/RFQ/Archive/bq{code_date}.zip"
            filename = f"bq{code_date}.zip"
            save_path = os.path.join(save_folder, filename)

            if os.path.exists(save_path):
                logging.info(f"File {save_path} already exists. Skipping download.")
                continue

            if click_ok_and_download(driver, url, save_path):
                logging.info(f"Downloaded {save_path}")
            else:
                logging.error(f"Error downloading {filename}")

    # Suppressing warnings
    warnings.filterwarnings('ignore')

    # Trigger the batch file
    batch_file_path = r"C:\Users\aeros\DIBBSBQ\STEPS\Z_STEP2.bat"  # Replace with the correct path to your batch file
    # batch_file_path = r"C:\Users\rclausing\Dibbs\STEPS\Z_STEP2.bat"  # Replace with the correct path to your batch file
    try:
        subprocess.run([batch_file_path], check=True)
        logging.info(f"Successfully executed {batch_file_path}")
    except subprocess.CalledProcessError as e:
        logging.error(f"Failed to execute {batch_file_path}: {str(e)}")

logging.info("Script execution completed.")

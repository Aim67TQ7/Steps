import os
import pandas as pd
import logging
import subprocess
import re
import csv
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
from webdriver_manager.chrome import ChromeDriverManager

# Configure logging
logging.basicConfig(filename='scraping_errors.log', level=logging.ERROR,
                    format='%(asctime)s - %(levelname)s - %(message)s')

def scrape_data(driver, url):
    try:
        driver.get(url)
        page_source = driver.page_source
        soup = BeautifulSoup(page_source, 'html.parser')
    except Exception as e:
        logging.error(f"Failed to scrape {url}: {e}")
        return None, None, None, None

    # Scraping the NSN information
    try:
        nsn_element = soup.find("div", class_="tcFrameCaptionLbl army")
        if nsn_element:
            nsn_text = nsn_element.text.strip()
            # Extract the 13-digit number after the colon
            nsn_match = re.search(r':\s*(\d{4}-\d{2}-\d{3}-\d{4})', nsn_text)
            nsn_info = nsn_match.group(1) if nsn_match else None
        else:
            nsn_info = None

        value_element = soup.find("td", class_="KeyCellRt")
        description_element = soup.find("td", class_="FieldCellLt", style="white-space:nowrap; overflow:hidden;")

        value = value_element.text.strip() if value_element else None
        description = description_element.text.strip() if description_element else None
    except Exception as e:
        logging.error(f"Error parsing initial data from {url}: {e}")
        return None, None, None, None

    # Scraping the MATT information from the second table
    matt_value = ""
    try:
        table = soup.find("table", class_="KeyTable")
        if table:
            matt_row = table.find("tr", lambda tag: tag.name == "tr" and tag.find("td", string="MATT"))
            if matt_row:
                matt_value = matt_row.find_all("td")[-1].text.strip()
    except Exception as e:
        logging.error(f"Error parsing MATT data from {url}: {e}")

    return nsn_info, value, description, matt_value

def count_lines(file_path):
    with open(file_path, 'r') as file:
        return sum(1 for line in file)

if __name__ == "__main__":
    driver = None
    try:
        options = Options()
        options.add_argument('--headless')
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        options.add_argument('user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3')
        
        print("Setting up ChromeDriver...")
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)
        print("ChromeDriver setup complete.")

        all_data = pd.DataFrame(columns=["NSN Info", "Value", "Description", "MATT"])

        file_path = "C:\\Users\\aeros\\DibbsBQ\\Haystack.txt"
        #file_path = "C:\\Users\\rclausing\\Dibbs\\Haystack.txt"
        base_url = "https://www.lqlite.com/Lq_FLIS.aspx?NSN="

        save_interval = 10  # Save after every 10 pages
        page_count = 0

        if os.path.exists(file_path):
            # Count and display the number of lines
            total_lines = count_lines(file_path)
            print(f"Total number of lines to be searched: {total_lines}")

            with open(file_path, 'r') as file:
                lines = file.readlines()
                for line_number, line in enumerate(lines, 1):
                    suffix = line.strip()  # Remove any leading/trailing whitespace
                    url = base_url + suffix
                    try:
                        print(f"Processing URL: {url}")
                        nsn_info, value, description, matt_value = scrape_data(driver, url)

                        scraped_data = pd.DataFrame({
                            "NSN Info": [nsn_info],
                            "Value": [value],
                            "Description": [description],
                            "MATT": [matt_value]
                        })

                        all_data = pd.concat([all_data, scraped_data], ignore_index=True)

                        page_count += 1

                        # Save to file after every 10 pages
                        if page_count % save_interval == 0:
                            file_name = r"C:\Users\aeros\DibbsBQ\LQ.txt"
                            #file_name = r"C:\Users\rclausing\Dibbs\LQ.txt"
                            all_data.to_csv(file_name, mode='a', header=not os.path.exists(file_name), index=False, quoting=csv.QUOTE_MINIMAL)
                            print(f"Data for {page_count} webpages saved to '{file_name}' (Processed {line_number}/{total_lines})")
                            print(f"Last MATT value saved: {matt_value}")  # Debug print
                            all_data = pd.DataFrame(columns=["NSN Info", "Value", "Description", "MATT"])  # Clear the cache

                    except Exception as e:
                        logging.error(f"Error processing URL {url}: {e}")
                        print(f"Error processing URL {url}: {e}")

            # Final save if there are any remaining records
            if not all_data.empty:
                file_name = r"C:\Users\aeros\DibbsBQ\LQ.txt"
                #file_name = r"C:\Users\rclausing\Dibbs\LQ.txt"
                all_data.to_csv(file_name, mode='a', header=not os.path.exists(file_name), index=False, quoting=csv.QUOTE_MINIMAL)
                print(f"Final data saved to '{file_name}'")
                print(f"Last MATT value in final save: {all_data['MATT'].iloc[-1]}")  # Debug print

        else:
            print(f"File {file_path} not found.")

    except Exception as e:
        logging.error(f"An error occurred: {e}")
        print(f"An error occurred: {e}")

    finally:
        if driver:
            driver.quit()

print("Script execution completed.")

try:
    script_dir = os.path.dirname(os.path.abspath(__file__))
    step6_path = os.path.join(script_dir, "Z_STEP6.vbs")
    
    if os.path.exists(step6_path):
        subprocess.run(["cscript", step6_path], check=True)
    else:
        print(f"Z_STEP6.vbs not found at {step6_path}")
except subprocess.CalledProcessError as e:
    print(f"Error running Z_STEP6.vbs: {e}")
except Exception as e:
    print(f"An unexpected error occurred while trying to run Z_STEP6.vbs: {e}")

print("Script execution completed, including Z_STEP6.vbs.")

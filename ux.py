import time
import re
import requests
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def visit_with_retry(driver, url, retries=3, wait_time=5):
    for attempt in range(retries):
        try:
            driver.get(url)
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.TAG_NAME, 'body')))
            
            # Scroll the page to load all content after it has loaded
            scroll_pause_time = 3
            scroll_height = driver.execute_script("return document.body.scrollHeight")
            
            while True:
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                time.sleep(scroll_pause_time)
                new_scroll_height = driver.execute_script("return document.body.scrollHeight")
                if new_scroll_height == scroll_height:
                    break
                scroll_height = new_scroll_height
            
            return True
        except Exception as e:
            print(f"Attempt {attempt + 1} failed for {url}: {e}")
            time.sleep(wait_time)
    return False

# Initialize WebDriver with options
options = webdriver.ChromeOptions()
options.add_argument("--disable-gpu")
options.add_argument("--disable-software-rasterizer")
driver = webdriver.Chrome(options=options)

# Open the main page
main_url = "https://circumeye.gr/companies/"
driver.get(main_url)

# Wait for the page to load fully
WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.TAG_NAME, 'a')))

# Scroll down the page to load content
scroll_pause_time = 3
scroll_height = driver.execute_script("return document.body.scrollHeight")

while True:
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(scroll_pause_time)
    new_scroll_height = driver.execute_script("return document.body.scrollHeight")
    if new_scroll_height == scroll_height:
        break
    scroll_height = new_scroll_height

print("Scrolling complete. Collecting links...")

# Extract all <a> links within the table
links = WebDriverWait(driver, 20).until(
    EC.presence_of_all_elements_located((By.XPATH, "//*[@id='coming-section']/div/div/div/table/tbody//a"))
)

hrefs = []
for link in links:
    href = link.get_attribute("href")
    company_name = link.text.strip()  # Get the text inside the <a> tag (company name)
    if href:
        hrefs.append({"Company Name": company_name, "URL": href})

print(f"Links collected: {len(hrefs)} links")

# Visit each link and collect additional links and usability info
collected_data = []

# Updated code to extract website links, excluding tel: and mailto: links
for href in hrefs:
    try:
        if not visit_with_retry(driver, href['URL']):
            print(f"Failed to load {href['Company Name']} after retries.")
            continue

        # Extract the company name (adjust the XPath as needed)
        company_name = href['Company Name']  # Use the name collected from the previous step

        # Find all <a> elements with the specified class that might contain website links
        website_links = driver.find_elements(By.XPATH, "//a[contains(@class, 'd-block') and contains(@class, 'text-blue') and contains(@class, 'zona-semi')]")
        valid_hrefs = [link.get_attribute("href") for link in website_links if link.get_attribute("href")]

        # Filter the links further to exclude social media and other unwanted types
        filtered_links = [
            link for link in valid_hrefs
            if "facebook.com" not in link.lower()
            and "instagram.com" not in link.lower()
            and "twitter.com" not in link.lower()
            and "mailto:" not in link.lower()  # Exclude mailto links
            and "tel:" not in link.lower()  # Exclude tel links
            and ("http://" in link.lower() or "https://" in link.lower())  # Only keep valid website URLs
            and not link.startswith("https://circumeye.gr")  # Exclude internal company links
        ]

        # Use the first valid link as the main website, if available
        primary_link = filtered_links[0] if filtered_links else None

        # Append the data with correct company names, website links, and usability fields
        collected_data.append({
            "Company Name": company_name,
            "Website": primary_link,
            "Design Evaluation": "",  # Add your evaluation here
            "Navigation Ease": "",    # Add your evaluation here
            "Speed Evaluation": "",   # Add your evaluation here
            "Checkout Process": "",   # Add your evaluation here
            "Search Functionality": "",  # Add your evaluation here
            "Filter Usability": "",    # Add your evaluation here
        })
        print(f"Website for {company_name}: {primary_link}")

    except Exception as e:
        print(f"Error visiting {href['Company Name']} ({href['URL']}): {e}")

# Convert collected data to DataFrame and save to Excel
df = pd.DataFrame(collected_data)
output_file = "company_websites_with_usability.xlsx"
df.to_excel(output_file, index=False)
print(f"Data saved to '{output_file}'.")

# Close the browser
driver.quit()

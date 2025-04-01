from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd

chrome_options = Options()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")

chrome_driver_path = r"C:\WebDriver\chromedriver-win64\chromedriver.exe"

service = Service(chrome_driver_path)
driver = webdriver.Chrome(service=service, options=chrome_options)

FACULTY_PAGE_URL = "https://eied.thapar.edu/faculty"
driver.get(FACULTY_PAGE_URL)

wait = WebDriverWait(driver, 10)
faculty_elements = wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME, "faculty-box")))

faculty_list = []
for faculty in faculty_elements:
    try:
        name = faculty.find_element(By.TAG_NAME, "strong").text.strip()
        designation = faculty.find_elements(By.TAG_NAME, "p")[1].text.strip()
        email = "N/A"
        p_tags = faculty.find_elements(By.TAG_NAME, "p")
        for i in range(len(p_tags) - 1):
            if "Email" in p_tags[i].text:
                email = p_tags[i + 1].text.strip()
                break
    except Exception as e:
        print(f"Error extracting data: {e}")
        continue

    faculty_list.append({"Name": name, "Designation": designation, "Email": email})

driver.quit()

df = pd.DataFrame(faculty_list)
df.to_excel("faculty_data.xlsx", index=False, engine="openpyxl")

print("✅ Scraping complete! Data saved to faculty_data.xlsx")

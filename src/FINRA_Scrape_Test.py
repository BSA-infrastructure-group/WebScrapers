from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, WebDriverException, TimeoutException

def search_webpage(search_query):
    url = "https://brokercheck.finra.org/search/genericsearch/grid"
    inputbox_xpath = "/html/body/bc-root/div/bc-home-page/div[1]/div[3]/investor-tools-finder/div[2]/form[1]/div/div[1]/input"
    button_arialabel = "button[aria-label='more details']"
    span_class = "text-primary-60 font-semibold"
    
    try:
        driver = webdriver.Chrome()  # Assuming Webdriver location is added to Path
        driver.get(url)

        try:
            search_box = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, inputbox_xpath))
            )
            search_box.send_keys(search_query)
            search_box.send_keys(Keys.RETURN)

            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, button_arialabel))
            )
            button = driver.find_element(By.CSS_SELECTOR, button_arialabel)
            button.click()

            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CLASS_NAME, span_class.replace(" ", ".")))
            )
            span_element = driver.find_element(By.CLASS_NAME, span_class.replace(" ", "."))
            extracted_text = span_element.text

            return "Inactive" if extracted_text == "Additional Information" else extracted_text

        except (NoSuchElementException, TimeoutException):
            return "Error"
        finally:
            driver.quit()  # Close the browser in any case
    except WebDriverException:
        return "Error"

# Example usage for testing

search_queries_list = [
    "7091750", "2372241", "5161392", "5190412", "5003238", "5712321", "2231151", "7562207",
    "5360389", "6065215", "6510605", "6956147", "6796064", "7122409", "6913876", "5214130",
    "4966091", "4639688"
]

results = []

# Loop through each search query
for query in search_queries_list:
    result = search_webpage(query)
    results.append(result)
    
for result in results:
    print(result)

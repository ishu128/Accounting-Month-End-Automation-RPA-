import time

import logging
log = logging.getLogger(__name__)
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC

def click_automated_ame_reports(driver, wait):
    """
    Opens Automated AME Reports and clicks Submit.
    """

    driver.switch_to.default_content()

    log.info("🔍 Searching for Automated AME Reports...")

    search_menu = wait.until(EC.element_to_be_clickable((By.ID, "miSearch")))
    driver.execute_script("arguments[0].click();", search_menu)

    search_box = wait.until(
        EC.visibility_of_element_located(
            (By.XPATH, "//li[@id='miSearch']//input[@type='text']")
        )
    )

    search_box.clear()
    search_box.send_keys("Automated AME Reports")

    result = wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, "//a[normalize-space()='Automated AME Reports']")
        )
    )

    driver.execute_script("arguments[0].click();", result)
    log.info("AME Reports page opened.")

    driver.switch_to.default_content()
    time.sleep(2)

    # Dynamically find iframe with Submit button
    iframes = driver.find_elements(By.TAG_NAME, "iframe")

    for frame in iframes:
        driver.switch_to.default_content()
        driver.switch_to.frame(frame)

        if driver.find_elements(By.XPATH, "//input[@class='filter_submit']"):
            submit = wait.until(
                EC.element_to_be_clickable((By.XPATH, "//input[@class='filter_submit']"))
            )
            driver.execute_script("arguments[0].click();", submit)
            log.info("AME Submit clicked.")
            time.sleep(5)
            break

    driver.switch_to.default_content()

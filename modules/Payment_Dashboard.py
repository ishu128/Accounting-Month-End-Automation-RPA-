import time
from datetime import datetime, timedelta
import logging
log = logging.getLogger(__name__)
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC

def open_payments_dashboard_and_capture(driver, wait, prop):
    """
    Opens Payments Dashboard.
    Enters property.
    Sets From Date = 1st of last month.
    Clicks Go.
    Takes screenshot.
    Creates Word document (not saved).
    """

    driver.switch_to.default_content()

    log.info("🔍 Searching for Payments Dashboard...")

    search_menu = wait.until(EC.element_to_be_clickable((By.ID, "miSearch")))
    driver.execute_script("arguments[0].click();", search_menu)

    search_box = wait.until(
        EC.visibility_of_element_located(
            (By.XPATH, "//li[@id='miSearch']//input[@type='text']")
        )
    )

    search_box.clear()
    search_box.send_keys("Payments Dashboard")

    result = wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, "//a[normalize-space()='Payments Dashboard']")
        )
    )

    driver.execute_script("arguments[0].click();", result)
    log.info("Payments Dashboard opened.")

    # Switch to filter iframe
    driver.switch_to.default_content()
    wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "filter")))

    # Enter property
    prop_input = wait.until(
        EC.visibility_of_element_located((By.ID, "YsiPropLookup_LookupCode"))
    )

    driver.execute_script("arguments[0].value = arguments[1];", prop_input, prop)
    prop_input.send_keys(Keys.TAB)

    # Calculate first of last month
    today = datetime.today()
    first_this_month = today.replace(day=1)
    last_month_last_day = first_this_month - timedelta(days=1)
    first_last_month = last_month_last_day.replace(day=1)

    formatted_date = first_last_month.strftime("%m/%d/%Y")

    from_date = wait.until(
        EC.visibility_of_element_located((By.ID, "FromDate_TextBox"))
    )

    from_date.clear()
    from_date.send_keys(formatted_date)
    from_date.send_keys(Keys.TAB)

    # Click Go
    go_button = wait.until(
        EC.element_to_be_clickable((By.ID, "YsiBtnGo_Button"))
    )

    driver.execute_script("arguments[0].click();", go_button)

    log.info("Waiting for dashboard data...")
    time.sleep(6)

    driver.switch_to.default_content()


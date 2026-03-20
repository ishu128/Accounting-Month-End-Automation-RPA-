from selenium.webdriver.common.by import By
from datetime import datetime, timedelta
import logging
log = logging.getLogger(__name__)
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
def click_datamart_reset(driver, wait):
    """
    Opens Reset Datamart via search.
    Selects required statuses.
    Clicks Submit.
    """

    driver.switch_to.default_content()

    log.info("🔍 Searching for Reset Datamart...")

    search_menu = wait.until(EC.element_to_be_clickable((By.ID, "miSearch")))
    driver.execute_script("arguments[0].click();", search_menu)

    search_box = wait.until(
        EC.visibility_of_element_located(
            (By.XPATH, "//li[@id='miSearch']//input[@type='text']")
        )
    )

    search_box.clear()
    search_box.send_keys("Reset Datamart")

    result = wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, "//a[normalize-space()='Reset Datamart']")
        )
    )

    driver.execute_script("arguments[0].click();", result)
    log.info("Reset Datamart page opened.")

    driver.switch_to.default_content()

    # Try switching into filter iframe if needed
    try:
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "filter")))
    except:
        pass

    # Select status options
    status_select = wait.until(
        EC.presence_of_element_located((By.NAME, "Status"))
    )

    select = Select(status_select)
    select.deselect_all()

    for option in ["Current", "Past", "Future", "Eviction", "Notice", "Vacant"]:
        select.select_by_value(option)

    log.info("Status options selected.")

        # ------------------------------------------------------------
    # Set From Month = Last Month (MM/YYYY)
    # ------------------------------------------------------------

    # Calculate last month
    today = datetime.today()
    first_this_month = today.replace(day=1)
    last_month_last_day = first_this_month - timedelta(days=1)

    last_month_mmyy = last_month_last_day.strftime("%m/%Y")

    # Locate dtPeriod input
    period_box = wait.until(
        EC.presence_of_element_located((By.NAME, "dtPeriod"))
    )

    # Clear and set value using JS (safer in Yardi)
    driver.execute_script(
        "arguments[0].value = arguments[1];",
        period_box,
        last_month_mmyy
    )

    log.info(f"📅 From Month set to: {last_month_mmyy}")

    submit_button = wait.until(
        EC.element_to_be_clickable((By.XPATH, "//input[@value='Submit']"))
    )

    driver.execute_script("arguments[0].click();", submit_button)

    log.info("Reset Datamart submitted.")
    driver.switch_to.default_content()



import time
from datetime import datetime, timedelta
import logging
log = logging.getLogger(__name__)
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from datetime import datetime, timedelta

def open_credit_card_dashboard_and_capture(driver, wait, prop_code):
    """
    Opens Credit Card Dashboard,
    filters by property + last month date,
    selects required status options,
    submits and waits for results.
    """

    driver.switch_to.default_content()

    log.info("🔍 Searching for Credit Card Dashboard...")

    # --------------------------------------------------
    # Open Search
    # --------------------------------------------------
    search_menu = wait.until(
        EC.element_to_be_clickable((By.ID, "miSearch"))
    )
    driver.execute_script("arguments[0].click();", search_menu)

    search_box = wait.until(
        EC.visibility_of_element_located(
            (By.XPATH, "//li[@id='miSearch']//input[@type='text']")
        )
    )

    search_box.clear()
    search_box.send_keys("Credit Card Dashboard")

    result = wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, "//a[normalize-space()='Credit Card Dashboard']")
        )
    )

    driver.execute_script("arguments[0].click();", result)

    log.info("📊 Credit Card Dashboard opened.")

    # --------------------------------------------------
    # 🔥 FIX: Dynamically locate correct iframe
    # (Original logic preserved, only iframe handling improved)
    # --------------------------------------------------
    driver.switch_to.default_content()

    log.info("🧭 Locating Credit Card Dashboard iframe...")

    found = False
    iframes = driver.find_elements(By.TAG_NAME, "iframe")

    for frame in iframes:
        driver.switch_to.default_content()
        driver.switch_to.frame(frame)

        if driver.find_elements(By.ID, "LookupProperty_LookupCode"):
            log.info("✅ Correct iframe found.")
            found = True
            break

    if not found:
        raise RuntimeError("❌ Could not locate Credit Card Dashboard iframe.")

    # --------------------------------------------------
    # Enter Property (ORIGINAL LOGIC)
    # --------------------------------------------------
    prop_input = wait.until(
        EC.visibility_of_element_located(
            (By.ID, "LookupProperty_LookupCode")
        )
    )

    driver.execute_script("arguments[0].value = arguments[1];", prop_input, prop_code)
    prop_input.send_keys(Keys.TAB)

    log.info(f"🏢 Property set in Credit Dashboard: {prop_code}")

    # --------------------------------------------------
    # First Day of Last Month (ORIGINAL LOGIC)
    # --------------------------------------------------
    today = datetime.today()
    first_this_month = today.replace(day=1)
    last_month_last_day = first_this_month - timedelta(days=1)
    first_last_month = last_month_last_day.replace(day=1)

    formatted_date = first_last_month.strftime("%m/%d/%Y")

    from_date = wait.until(
        EC.visibility_of_element_located(
            (By.ID, "txtCreditCardDateFrom_TextBox")
        )
    )

    from_date.clear()
    from_date.send_keys(formatted_date)
    from_date.send_keys(Keys.TAB)

    log.info(f"📅 Date set: {formatted_date}")

    # --------------------------------------------------
    # Select 4 Status Options (ORIGINAL LOGIC KEPT)
    # --------------------------------------------------
    status_select = wait.until(
        EC.presence_of_element_located((By.ID, "cmbStatus_MultiSelect"))
    )

    select = Select(status_select)

    try:
        select.deselect_all()
    except:
        pass

    # Your original selection logic
    select.select_by_value("DISPUTED")
    select.select_by_value("PARTIAL_DISPUTED")
    select.select_by_value("DISPUTED - RECEIPT REVERSED")
    select.select_by_value("DISPUTE WON")

    log.info("✅ 4 status options selected.")

    # --------------------------------------------------
    # Click Submit (ORIGINAL LOGIC)
    # --------------------------------------------------
    submit_button = wait.until(
        EC.element_to_be_clickable((By.ID, "BtnSubmit_Button"))
    )

    driver.execute_script("arguments[0].click();", submit_button)

    log.info("▶ Submit clicked. Waiting for results...")

    driver.switch_to.default_content()

    time.sleep(6)

    log.info("✅ Credit Card Dashboard data loaded.")


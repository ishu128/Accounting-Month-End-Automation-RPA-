from datetime import datetime
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
import logging
log = logging.getLogger(__name__)
def open_nsf_register(driver, wait):
    """
    Opens NSF Register, clicks Submit, waits for data to load.
    """

    driver.switch_to.default_content()
    log.info("🔍 Searching for NSF Register...")

    # ------------------------------------------------------------
    # Open search
    # ------------------------------------------------------------
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
    search_box.send_keys("NSF Register")

    result = wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, "//a[normalize-space()='NSF Register']")
        )
    )

    driver.execute_script("arguments[0].click();", result)
    log.info("📄 NSF Register page opened.")

    # ------------------------------------------------------------
    # 🔥 FIX: Dynamically find correct iframe
    # ------------------------------------------------------------
    driver.switch_to.default_content()
    log.info("🧭 Locating NSF iframe...")

    found = False
    iframes = driver.find_elements(By.TAG_NAME, "iframe")

    for frame in iframes:
        driver.switch_to.default_content()
        driver.switch_to.frame(frame)

        if driver.find_elements(By.XPATH, "//input[@class='filter_submit']"):
            log.info("✅ Correct NSF iframe found.")
            found = True
            break

    if not found:
        raise RuntimeError("❌ Could not locate NSF Register iframe.")


    # ------------------------------------------------------------
    # Enter Current Month in begmonth and endmonth
    # ------------------------------------------------------------
    from datetime import datetime
    current_month = datetime.today().strftime("%m/%Y")

    # Begin Month
    beg_month_box = wait.until(
        EC.presence_of_element_located((By.NAME, "begmonth"))
    )

    driver.execute_script("arguments[0].value = arguments[1];", beg_month_box, current_month)
    beg_month_box.send_keys(Keys.TAB)

    # End Month
    end_month_box = wait.until(
        EC.presence_of_element_located((By.NAME, "endmonth"))
    )

    driver.execute_script("arguments[0].value = arguments[1];", end_month_box, current_month)
    end_month_box.send_keys(Keys.TAB)

    log.info(f"📅 NSF month set to {current_month}")


    # ------------------------------------------------------------
    # Click Submit (UNCHANGED)
    # ------------------------------------------------------------
    submit_btn = wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, "//input[@class='filter_submit']")
        )
    )

    driver.execute_script("arguments[0].click();", submit_btn)

    log.info("▶ NSF Submit clicked. Waiting for report to load...")

    wait.until(EC.presence_of_element_located((By.TAG_NAME, "table")))

    driver.switch_to.default_content()
    log.info("✅ NSF Register data loaded.")


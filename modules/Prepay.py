import time
import logging
import logging
log = logging.getLogger(__name__)
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait

def click_convert_to_prepay(driver, wait, prop, batch_option):
    """
    Searches and processes Convert Credit to Prepay page.

    If batch_option == "yes":
        → Runs full Display + Post loop until no Post button remains.

    If batch_option == "no":
        → Runs only Display step (no Post loop).
    """

    try:
        # ------------------------------------------------------------
        # Safety Check – Browser still alive?
        # ------------------------------------------------------------
        driver.current_url

        driver.switch_to.default_content()
        log.info("🔍 Searching for Convert Credit to Prepay...")

        # ------------------------------------------------------------
        # Open Global Search
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
        search_box.send_keys("Convert Credit to Prepay")

        result = wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, "//a[normalize-space()='Convert Credit to Prepay']")
            )
        )

        driver.execute_script("arguments[0].click();", result)
        log.info("✅ Convert Credit to Prepay page opened.")

        # ------------------------------------------------------------
        # Switch to filter iframe (Yardi loads inside this)
        # ------------------------------------------------------------
        driver.switch_to.default_content()
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "filter")))
        log.info("🧭 Switched to filter iframe.")

        # ------------------------------------------------------------
        # Enter Property Code Again
        # ------------------------------------------------------------
        prop_input = wait.until(
            EC.visibility_of_element_located(
                (By.ID, "PropertyLkp_LookupCode")
            )
        )

        driver.execute_script(
            "arguments[0].value = arguments[1];",
            prop_input,
            prop
        )
        prop_input.send_keys(Keys.TAB)

        log.info(f"🏢 Property {prop} entered on Convert page.")

        # ------------------------------------------------------------
        # Click Display
        # ------------------------------------------------------------
        display_button = wait.until(
            EC.element_to_be_clickable((By.ID, "DisplayBtn_Button"))
        )

        driver.execute_script("arguments[0].click();", display_button)
        log.info("▶ Display clicked. Waiting for grid to load...")
        time.sleep(4)

        # ------------------------------------------------------------
        # Batch Posting Logic
        # ------------------------------------------------------------
        if batch_option.lower() == "yes":

            log.info("🔁 Batch Posting = YES → Processing Post loop.")

            while True:
                try:
                    post_button = WebDriverWait(driver, 5).until(
                        EC.visibility_of_element_located(
                            (By.ID, "PostBtn_Button")
                        )
                    )

                    log.info("🟢 Post button visible. Clicking Post...")
                    driver.execute_script(
                        "arguments[0].click();",
                        post_button
                    )

                    time.sleep(6)

                except TimeoutException:
                    log.info("✅ No more Post buttons. Convert complete.")
                    break

        else:
            log.info("⏭ Batch Posting = NO → Skipping Post loop.")

        # ------------------------------------------------------------
        # Exit to default content
        # ------------------------------------------------------------
        driver.switch_to.default_content()
        log.info("🔙 Exited Convert to Prepay page successfully.")

    except Exception as e:
        log.exception(f"❌ Convert to Prepay failed: {e}")
        driver.switch_to.default_content()
        raise


import time
import logging
log = logging.getLogger(__name__)
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
def set_property(driver, wait, prop):
    """
    Sets property context inside Voyager using filter iframe.
    """

    log.info(f"🏢 Setting property → {prop}")

    # Always reset to top DOM before switching frames
    driver.switch_to.default_content()

    # Switch into filter iframe (property selector lives here)
    wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "filter")))

    # Locate property input field
    prop_input = wait.until(
        EC.visibility_of_element_located((By.ID, "PropLookup_LookupCode"))
    )

    # Use JS to set value safely (Yardi sometimes blocks send_keys)
    driver.execute_script("arguments[0].value = arguments[1];", prop_input, prop)

    # Trigger blur/change event
    prop_input.send_keys(Keys.TAB)

    # Click Go button to reload property context
    wait.until(EC.element_to_be_clickable((By.ID, "YSIGo_Button"))).click()

    log.info("Property Go clicked. Waiting for refresh...")
    time.sleep(5)  # allow reload

    driver.switch_to.default_content()

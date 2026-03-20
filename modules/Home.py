from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC

import logging
log = logging.getLogger(__name__)
def go_home(driver, wait):
    """
    Navigates back to Voyager Home page.
    """

    driver.switch_to.default_content()

    home_button = wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, "//a[normalize-space()='Home']")
        )
    )

    driver.execute_script("arguments[0].click();", home_button)

    wait.until(EC.presence_of_element_located((By.ID, "side-menu-wrap")))

    log.info("Returned to Home.")

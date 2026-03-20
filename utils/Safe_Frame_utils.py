# ============================================================
# IMPORTS - Safe_Frame_utils.py
# ============================================================

import time
import logging
from selenium.webdriver.support import expected_conditions as EC
import logging
log = logging.getLogger(__name__)
def safe_frame_switch(frame_locator, wait, retries=2):

    for attempt in range(retries + 1):
        try:
            driver.switch_to.default_content()
            wait.until(EC.frame_to_be_available_and_switch_to_it(frame_locator))
            return True
        except Exception as e:
            log.warning(f"Iframe switch failed attempt {attempt+1}")
            time.sleep(2)

    raise RuntimeError("Failed to switch iframe after retries")

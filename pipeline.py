import time
import logging
from utils.browser_utils import is_browser_alive

log = logging.getLogger(__name__)

def run_step(driver, step_name, step_function, retries=1, retry_delay=3):

    attempt = 0

    while attempt <= retries:

        if not is_browser_alive(driver):
            log.error("🚨 Browser session lost. Stopping automation.")
            raise RuntimeError("Browser closed manually.")

        try:
            log.info(f"▶ Running Step: {step_name} | Attempt {attempt + 1}")
            step_function()
            log.info(f"✅ Step Completed: {step_name}")
            return True

        except Exception as e:

            if not is_browser_alive(driver):
                log.error("🚨 Browser closed during step.")
                raise RuntimeError("Browser closed manually.")

            log.warning(f"⚠ Step Failed: {step_name} | {e}")

            attempt += 1

            if attempt > retries:
                log.error(f"❌ Step Skipped After Retry: {step_name}")
                return False

            try:
                driver.switch_to.default_content()
            except:
                pass

            time.sleep(retry_delay)


def execute_pipeline(driver, pipeline, steps_to_run=None, steps_to_skip=None):

    if steps_to_run is None:
        steps_to_run = []

    if steps_to_skip is None:
        steps_to_skip = []

    for step_name, step_function in pipeline:

        if step_name in steps_to_skip:
            log.info(f"⏭ Skipping Step: {step_name}")
            continue

        if steps_to_run and step_name not in steps_to_run:
            continue

        run_step(driver, step_name, step_function, retries=2)

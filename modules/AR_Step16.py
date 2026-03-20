import os
import time
from datetime import datetime
import logging
log = logging.getLogger(__name__)
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from utils.word_utils import add_formatted_section_to_word
from selenium.webdriver.support.ui import Select
from datetime import datetime
from utils.download_utils import wait_for_download_and_rename


def step16(driver, wait, prop_code, prop_name, folder_path):
    """
    Step 10 – Residential AR Analytics (Write Offs)
    """

    driver.switch_to.default_content()
    log.info("🔍 Searching for Residential AR Analytics...")

    # ------------------------------------------------------------
    # Open from search
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
    search_box.send_keys("Residential AR Analytics")

    result = wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, "//a[contains(text(),'Residential AR')]")
        )
    )

    driver.execute_script("arguments[0].click();", result)
    log.info("📊 Residential AR Analytics page opened.")


    # ------------------------------------------------------------
    # Switch to filter iframe (wait for full reload)
    # ------------------------------------------------------------
    driver.switch_to.default_content()

    wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "filter")))
    log.info("🧭 Switched to filter iframe.")

    # VERY IMPORTANT: re-locate element after frame switch
    prop_input = wait.until(
        EC.presence_of_element_located((By.ID, "PropertyLookup_LookupCode"))
    )

    # Use JS instead of send_keys (prevents stale after postback)
    driver.execute_script(
        "arguments[0].value = arguments[1];",
        prop_input,
        prop_code
    )
    prop_input.send_keys(Keys.TAB)

    log.info("🏢 Property entered.")

    # Wait for postback reload to finish
    wait.until(
        EC.presence_of_element_located((By.ID, "tenstatus_MultiSelect"))
    )

    # ------------------------------------------------------------
    # Select Status (Not required for Step 16)
    # ------------------------------------------------------------
   

    # ------------------------------------------------------------
    # Select Report Type
    # ------------------------------------------------------------
    report_type = Select(
        wait.until(
            EC.element_to_be_clickable((By.ID, "cmbReportType_DropDownList"))
        )
    )
    report_type.select_by_visible_text("Financial Aged Receivable")
    log.info("📈 Report Type selected.")

    # Wait for reload
    wait.until(
        EC.presence_of_element_located((By.ID, "cmbGroupby_DropDownList"))
    )

    # ------------------------------------------------------------
    # Select Group By
    # ------------------------------------------------------------
    group_by = Select(
        wait.until(
            EC.element_to_be_clickable((By.ID, "cmbGroupby_DropDownList"))
        )
    )
    group_by.select_by_visible_text("Resident by Charge Code")
    log.info("📊 Group By selected.")

    # ------------------------------------------------------------
    # Enter Current Month
    # ------------------------------------------------------------
    current_month = datetime.today().strftime("%m/%Y")

    as_of_box = wait.until(
        EC.presence_of_element_located((By.ID, "txtDate1_TextBox"))
    )

    driver.execute_script(
        "arguments[0].value = arguments[1];",
        as_of_box,
        current_month
    )
    as_of_box.send_keys(Keys.TAB)

    log.info(f"📅 As Of Month set to {current_month}")

    # ------------------------------------------------------------
    # Click Display
    # ------------------------------------------------------------
    display_btn = wait.until(
        EC.element_to_be_clickable((By.ID, "btnDisplay_Button"))
    )

    driver.execute_script("arguments[0].click();", display_btn)
    log.info("▶ Display clicked. Waiting for report to load...")

    download_dir = os.path.join(os.path.expanduser("~"), "Downloads")

    # 🔥 IMPORTANT — capture files BEFORE download
    files_before = set(os.listdir(download_dir))

    wait.until(EC.element_to_be_clickable((By.ID, "btnExcel_Button"))).click()

    log.info("📥 Waiting for Financial Aged Receivable Excel download...")

    return wait_for_download_and_rename(
        download_dir=download_dir,
        files_before=files_before,
        target_folder=folder_path,
        base_filename=f"{prop_name}_Financial Aged Receivable_Step 16"
    )

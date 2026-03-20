import os
import time
from datetime import datetime
import logging
log = logging.getLogger(__name__)
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC

from utils.download_utils import wait_for_download_and_rename

def run_trial_balance_activity(driver, wait, prop_code, prop_name, folder_path):

    driver.switch_to.default_content()
    log.info("🔍 Searching for Financial Analytics...")

    # ---------------------------------------------------------
    # Open Search
    # ---------------------------------------------------------
    wait.until(EC.element_to_be_clickable((By.ID, "miSearch"))).click()

    search_box = wait.until(
        EC.visibility_of_element_located(
            (By.XPATH, "//li[@id='miSearch']//input[@type='text']")
        )
    )
    search_box.clear()
    search_box.send_keys("Financial Analytics")

    wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, "//a[contains(text(),'Financial Analytics')]")
        )
    ).click()

    log.info("📊 Financial Analytics opened.")

    # ---------------------------------------------------------
    # Switch to correct iframe
    # ---------------------------------------------------------
    driver.switch_to.default_content()

    for frame in driver.find_elements(By.TAG_NAME, "iframe"):
        driver.switch_to.default_content()
        driver.switch_to.frame(frame)
        if driver.find_elements(By.ID, "PropertyID_LookupCode"):
            break
    else:
        raise RuntimeError("❌ Financial Analytics iframe not found.")

    # ---------------------------------------------------------
    # Enter Property Code
    # ---------------------------------------------------------
    prop_box = wait.until(
        EC.presence_of_element_located((By.ID, "PropertyID_LookupCode"))
    )
    prop_box.clear()
    driver.execute_script("arguments[0].value = arguments[1];", prop_box, prop_code)
    prop_box.send_keys(Keys.TAB)

    # ---------------------------------------------------------
    # Enter Book = Accrual
    # ---------------------------------------------------------
    book_box = wait.until(
        EC.presence_of_element_located((By.ID, "BookID_LookupCode"))
    )
    book_box.clear()
    driver.execute_script("arguments[0].value = 'Accrual';", book_box)
    book_box.send_keys(Keys.TAB)

    # ---------------------------------------------------------
    # Enter Tree = ysi_tb
    # ---------------------------------------------------------
    tree_box = wait.until(
        EC.presence_of_element_located((By.ID, "TreeID_LookupCode"))
    )
    tree_box.clear()
    driver.execute_script("arguments[0].value = 'ysi_tb';", tree_box)
    tree_box.send_keys(Keys.TAB)

    # ---------------------------------------------------------
    # Select Trial Balance from dropdown
    # ---------------------------------------------------------
    Select(wait.until(
        EC.presence_of_element_located((By.ID, "ReportNum_DropDownList"))
    )).select_by_visible_text("Trial Balance")

    # ---------------------------------------------------------
    # Set From / To Month (MM/YYYY)
    # ---------------------------------------------------------
    current_mmyy = datetime.today().strftime("%m/%Y")

    from_box = wait.until(
        EC.presence_of_element_located((By.ID, "FromMMYY_TextBox"))
    )
    to_box = wait.until(
        EC.presence_of_element_located((By.ID, "ToMMYY_TextBox"))
    )

    driver.execute_script("arguments[0].value = arguments[1];", from_box, current_mmyy)
    driver.execute_script("arguments[0].value = arguments[1];", to_box, current_mmyy)

    # ---------------------------------------------------------
    # Ensure Suppress Zero is Checked
    # ---------------------------------------------------------
    suppress_checkbox = wait.until(
        EC.presence_of_element_located((By.ID, "SupressZero_CheckBox"))
    )

    if not suppress_checkbox.is_selected():
        suppress_checkbox.click()

    # ---------------------------------------------------------
    # Click Display
    # ---------------------------------------------------------
    wait.until(
        EC.element_to_be_clickable((By.ID, "Display_Button"))
    ).click()

    log.info("📊 Waiting for Trial Balance to load...")
    time.sleep(5)  # you can replace with better wait later

    # ---------------------------------------------------------
    # Prepare Download Tracking
    # ---------------------------------------------------------
    download_dir = os.path.join(os.path.expanduser("~"), "Downloads")
    files_before = set(os.listdir(download_dir))

    # ---------------------------------------------------------
    # Click Excel
    # ---------------------------------------------------------
    wait.until(
        EC.element_to_be_clickable((By.ID, "Excel_Button"))
    ).click()

    log.info("📥 Waiting for Trial Balance Excel download...")

    return wait_for_download_and_rename(
        download_dir=download_dir,
        files_before=files_before,
        target_folder=folder_path,
        base_filename=f"{prop_name}_Trial_Balance"
    )

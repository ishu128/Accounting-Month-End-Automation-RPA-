from selenium.webdriver.common.by import By
import logging
log = logging.getLogger(__name__)
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
def open_transaction_register_analytics(driver, wait, prop_code):
    """
    Opens Transaction Register Analytics
    Selects Batch Register
    Filters:
        - Property
        - Batch Status = Open
        - Batch Type = Payable, Receipt, Charge
    Then clicks Display and waits for data.
    """

    driver.switch_to.default_content()
    log.info("🔍 Searching for Transaction Register Analytics...")

    # ---------------------------------------------------------
    # Open Search
    # ---------------------------------------------------------
    search_menu = wait.until(EC.element_to_be_clickable((By.ID, "miSearch")))
    driver.execute_script("arguments[0].click();", search_menu)

    search_box = wait.until(
        EC.visibility_of_element_located(
            (By.XPATH, "//li[@id='miSearch']//input[@type='text']")
        )
    )

    search_box.clear()
    search_box.send_keys("Transaction Register Analytics")

    result = wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, "//a[normalize-space()='Transaction Register Analytics']")
        )
    )

    driver.execute_script("arguments[0].click();", result)
    log.info("✅ Transaction Register Analytics opened.")

    # ---------------------------------------------------------
    # 🔥 FIX: Dynamically locate correct iframe
    # ---------------------------------------------------------
    driver.switch_to.default_content()

    log.info("🧭 Locating correct iframe...")

    found = False
    iframes = driver.find_elements(By.TAG_NAME, "iframe")

    for frame in iframes:
        driver.switch_to.default_content()
        driver.switch_to.frame(frame)

        if driver.find_elements(By.ID, "TranType_DropDownList"):
            log.info("✅ Correct iframe located.")
            found = True
            break

    if not found:
        raise RuntimeError("❌ Could not locate Transaction Register iframe.")

    # ---------------------------------------------------------
    # Select Batch Register (UNCHANGED)
    # ---------------------------------------------------------
    tran_type = wait.until(
        EC.presence_of_element_located((By.ID, "TranType_DropDownList"))
    )

    Select(tran_type).select_by_visible_text("Batch Register")
    log.info("📑 Selected Batch Register.")

    # ---------------------------------------------------------
    # Enter Property Code (UNCHANGED)
    # ---------------------------------------------------------
    prop_input = wait.until(
        EC.visibility_of_element_located((By.ID, "PropertyID_LookupCode"))
    )

    driver.execute_script("arguments[0].value = arguments[1];", prop_input, prop_code)
    prop_input.send_keys(Keys.TAB)

    log.info(f"🏢 Property entered: {prop_code}")

    # ---------------------------------------------------------
    # Select Batch Status = Open (UNCHANGED)
    # ---------------------------------------------------------
    batch_status = wait.until(
        EC.presence_of_element_located((By.ID, "BatchStatus_DropDownList"))
    )

    Select(batch_status).select_by_visible_text("Open")
    log.info("📂 Batch Status set to Open.")

    # ---------------------------------------------------------
    # Select Batch Types (UNCHANGED)
    # ---------------------------------------------------------
    batch_type = wait.until(
        EC.presence_of_element_located((By.ID, "BatchType_MultiSelect"))
    )

    select_batch_type = Select(batch_type)

    try:
        select_batch_type.deselect_all()
    except:
        pass

    select_batch_type.select_by_visible_text("Payable")
    select_batch_type.select_by_visible_text("Receipt")
    select_batch_type.select_by_visible_text("Charge")

    log.info("🧾 Selected Batch Types: Payable, Receipt, Charge.")

    # ---------------------------------------------------------
    # Click Display
    # ---------------------------------------------------------
    display_btn = wait.until(
        EC.element_to_be_clickable((By.ID, "Display_Button"))
    )

    driver.execute_script("arguments[0].click();", display_btn)

    log.info("▶ Display clicked. Waiting for report data...")

    # Wait for table instead of blind sleep
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "table")))

    driver.switch_to.default_content()
    log.info("✅ Transaction Register data loaded.")

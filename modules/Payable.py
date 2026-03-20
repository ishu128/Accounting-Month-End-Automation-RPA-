import time
from datetime import datetime, timedelta
import logging
log = logging.getLogger(__name__)
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
def run_payable_analytics_aging(driver, wait, prop_code, document):
    """
    Step 13 – Payable Analytics (Aging)
    Verifies unclaimed refund process (over 90 days)
    """

    driver.switch_to.default_content()
    log.info("🔍 Searching for Payable Analytics...")

    # ------------------------------------------------------------
    # 1️⃣ Open from Search
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
    search_box.send_keys("Payable Analytics")

    result = wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, "//a[contains(text(),'Payable Analytics')]")
        )
    )

    driver.execute_script("arguments[0].click();", result)
    log.info("📊 Payable Analytics page opened.")

    # ------------------------------------------------------------
    # 2️⃣ Switch to correct iframe dynamically
    # ------------------------------------------------------------
    driver.switch_to.default_content()

    found = False
    iframes = driver.find_elements(By.TAG_NAME, "iframe")

    for frame in iframes:
        driver.switch_to.default_content()
        driver.switch_to.frame(frame)

        if driver.find_elements(By.ID, "ReportType_DropDownList"):
            log.info("✅ Correct iframe located.")
            found = True
            break

    if not found:
        raise RuntimeError("❌ Could not locate Payable Analytics iframe.")

    # ------------------------------------------------------------
    # 3️⃣ Select Aging Report Type
    # ------------------------------------------------------------
    Select(
        wait.until(
            EC.presence_of_element_located((By.ID, "ReportType_DropDownList"))
        )
    ).select_by_visible_text("Aging")

    log.info("📑 Report Type selected: Aging")

    # ------------------------------------------------------------
    # 4️⃣ Enter Property Code (JS safer in Yardi)
    # ------------------------------------------------------------
    property_box = wait.until(
        EC.presence_of_element_located((By.ID, "PropertyLookup_LookupCode"))
    )

    driver.execute_script(
        "arguments[0].value = arguments[1];",
        property_box,
        prop_code
    )
    property_box.send_keys(Keys.TAB)

    log.info(f"🏢 Property entered: {prop_code}")

    # ------------------------------------------------------------
    # 5️⃣ Enter GL Code
    # ------------------------------------------------------------
    account_box = wait.until(
        EC.presence_of_element_located((By.ID, "AccountLookup_LookupCode"))
    )

    driver.execute_script(
        "arguments[0].value = '23040-000';",
        account_box
    )
    account_box.send_keys(Keys.TAB)

    log.info("🏢 GL Code Entered: 23040-000")

    # ------------------------------------------------------------
    # 6️⃣ Set Date Filters
    # ------------------------------------------------------------
    today_str = datetime.today().strftime("%m/%d/%Y")
    next_month_str = (
        datetime.today().replace(day=1) + timedelta(days=32)
    ).strftime("%m/%Y")

    age_as_of = wait.until(
        EC.presence_of_element_located((By.ID, "AgeAsOf_TextBox"))
    )

    driver.execute_script(
        "arguments[0].value = arguments[1];",
        age_as_of,
        today_str
    )

    period_to = wait.until(
        EC.presence_of_element_located((By.ID, "PeriodTo_TextBox"))
    )

    driver.execute_script(
        "arguments[0].value = arguments[1];",
        period_to,
        next_month_str
    )

    log.info("📅 Date filters set.")

    # ------------------------------------------------------------
    # 7️⃣ Clear Optional Filters
    # ------------------------------------------------------------
    log.info("🧹 Clearing optional filters...")

    optional_fields = [
        "APAccountLookup_LookupCode",
        "ControlNoFrom_TextBox",
        "PeriodFrom_TextBox",
        "DueDateFromText_TextBox",
        "BankLookup_LookupCode",
        "CompanyLookup_LookupCode",
        "VendorLookup_LookupCode"
    ]

    for field_id in optional_fields:
        try:
            field = driver.find_element(By.ID, field_id)
            if field.is_enabled():
                driver.execute_script("arguments[0].value = '';", field)
        except:
            pass

    # ------------------------------------------------------------
    # 8️⃣ Click Display
    # ------------------------------------------------------------
    display_btn = wait.until(
        EC.element_to_be_clickable((By.ID, "Display_Button"))
    )

    driver.execute_script("arguments[0].click();", display_btn)

    log.info("▶ Display clicked. Waiting for report to load...")

    driver.switch_to.default_content()

    time.sleep(5)  # allow grid to load

    log.info("✅ Payable Aging report loaded.")


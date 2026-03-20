import time
import logging
import logging
log = logging.getLogger(__name__)
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from io import BytesIO


# ============================================================
# FULL YARDI LOGIN WITH MFA DISPLAY + TAB SWITCH
# ============================================================

def Yardi_Login(driver, username, password):

    log.info("🌐 Navigating to Yardi login page...")
    driver.get("https://greystarus.yardione.com/Account/Login")

    wait = WebDriverWait(driver, 20)

    # --------------------------------------------------------
    # STEP 1: Click External Auth (if visible)
    # --------------------------------------------------------
    try:
        ext_button = wait.until(
            EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "a.btn.btn-primary.btn-block.btn-shadow.mb-4")
            )
        )
        ext_button.click()
        log.info("External Auth clicked.")
    except TimeoutException:
        log.info("External Auth not shown. Continuing...")

    # --------------------------------------------------------
    # STEP 2: Enter Username
    # --------------------------------------------------------
    user_box = wait.until(EC.presence_of_element_located((By.ID, "i0116")))
    user_box.clear()
    time.sleep(2)
    user_box.send_keys(username)
    wait.until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()

    time.sleep(2)

    # --------------------------------------------------------
    # STEP 3: Enter Password
    # --------------------------------------------------------
    time.sleep(1)
    pass_box = wait.until(EC.presence_of_element_located((By.NAME, "passwd")))
    pass_box.clear()
    time.sleep(2)
    pass_box.send_keys(password)
    wait.until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()

    # --------------------------------------------------------
    # STEP 4: MFA HANDLING (VISIBLE CODE DISPLAY)
    # --------------------------------------------------------
    try:
        log.info("⏳ Waiting for MFA prompt...")

        try:
            mfa_code_element = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.ID, "idRichContext_DisplaySign"))
            )
        except TimeoutException:
            approve_prompt = wait.until(
                EC.element_to_be_clickable(
                    (By.XPATH, "//div[contains(text(),'Approve a request')]")
                )
            )
            approve_prompt.click()
            mfa_code_element = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.ID, "idRichContext_DisplaySign"))
            )

        mfa_code = mfa_code_element.text
        log.info("========================================")
        log.info(f"    >>> ENTER THIS CODE: {mfa_code} <<<")
        log.info("========================================")

        WebDriverWait(driver, 90).until(
            lambda d: "SAS/ProcessAuth" in d.current_url or "yardione.com" in d.current_url
        )

        log.info("MFA Approved.")

        # ----------------------------------------------------
        # STEP 5: Handle Stay Signed In
        # ----------------------------------------------------
        try:
            no_button = WebDriverWait(driver, 3).until(
                EC.element_to_be_clickable((By.ID, "idBtn_Back"))
            )
            no_button.click()
            log.info("'Stay signed in?' → Clicked No.")
        except TimeoutException:
            log.info("Stay signed in dialog not shown.")

        # Wait for dashboard
        wait.until(EC.presence_of_element_located((By.CLASS_NAME, "dashboard-tile")))

    except TimeoutException:
        raise RuntimeError("Login failed during MFA.")

    # --------------------------------------------------------
    # STEP 6: Click Voyager Tile
    # --------------------------------------------------------
    voyager_tile = wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, "//a[contains(@href,'yardipcu.com')]")
        )
    )

    driver.execute_script("arguments[0].click();", voyager_tile)

    WebDriverWait(driver, 20).until(lambda d: len(d.window_handles) > 1)

    # Close first tab
    driver.switch_to.window(driver.window_handles[0])
    driver.close()

    # Focus Voyager
    driver.switch_to.window(driver.window_handles[0])

    log.info(f"🌍 Focused on Voyager: {driver.current_url}")


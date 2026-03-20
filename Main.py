import logging
import os
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager

from pipeline import execute_pipeline

# ------------------------------
# UTILS
# ------------------------------
from utils.credentials import read_credentials
from utils.Read_Excel import read_property_data
from utils.Create_Folder import create_property_folder
from utils.Creat_Word import create_word_document
from utils.browser_utils import is_browser_alive
from utils.word_utils import add_formatted_section_to_word
from utils.file_paths import get_tieout_file

# ------------------------------
# MODULES
# ------------------------------
from modules.Login import Yardi_Login
from modules.Set_Property import set_property
from modules.Home import go_home
from modules.daily_activity import check_move_activity_current_ame
from modules.Prepay import click_convert_to_prepay
from modules.Reset_Data import click_datamart_reset
from modules.AME_Reports import click_automated_ame_reports
from modules.Credit_Card import open_credit_card_dashboard_and_capture
from modules.Payment_Dashboard import open_payments_dashboard_and_capture
from modules.TRA import open_transaction_register_analytics
from modules.NSF import open_nsf_register
from modules.Write_Off import run_residential_ar_analytics
from modules.security_deposit import run_security_deposit_full
from modules.Payable import run_payable_analytics_aging
from modules.gpr import run_gpr_full
from modules.gpr import run_gpr_activity
from modules.gpr import review_gpr_report
from modules.trial_balance import run_trial_balance_activity
from modules.AR_Step16 import step16
from modules.Tie_Out import run_tie_out
from modules.Helper import create_helper_sheet
from modules.Formula import write_tieout_formulas



# ============================================================
# MAIN EXECUTION
# ============================================================

if __name__ == "__main__":

    # --------------------------------------------------------
    # Logging Configuration
    # --------------------------------------------------------
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s"
    )

    log = logging.getLogger(__name__)

    log.info("🚀 Starting AME Full Automation...")

    # --------------------------------------------------------
    # STEP CONTROL CONFIGURATION
    # --------------------------------------------------------
    # Leave empty list [] → runs ALL steps
    # Add step names → runs only those
    steps_to_run = []

    # Add step names here to skip specific steps
    steps_to_skip = ["GPR Download", "GPR Check"]



    # Example:
    #steps_to_run = ["Gross Potential Rent"]

    #steps_to_skip = []
    # Example:
    #steps_to_skip = ["Reset Datamart", "Daily Activity AME", "Automated AME Reports", "Payments Dashboard", "Credit Card Dashboard", "Transaction Register", "NSF Register", "Gross Potential Rent", "Convert to Prepay (NO)", "Write Off"]



    # --------------------------------------------------------
    # BROWSER CONFIGURATION
    # --------------------------------------------------------
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")

    # Keep browser open after script ends
    chrome_options.add_experimental_option("detach", True)

    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=chrome_options
    )

    wait = WebDriverWait(driver, 30)

    try:

        # --------------------------------------------------------
        # 1️⃣ Read Credentials
        # --------------------------------------------------------
        username, password = read_credentials()

        # --------------------------------------------------------
        # 2️⃣ Read Property Excel Data
        # --------------------------------------------------------
        properties = read_property_data()

        if not properties:
            log.warning("⚠ No properties found. Exiting.")
            exit()

        # --------------------------------------------------------
        # 3️⃣ Login to Yardi
        # --------------------------------------------------------
        Yardi_Login(driver, username, password)

        # --------------------------------------------------------
        # 4️⃣ Process Each Property
        # --------------------------------------------------------

        for prop_data in properties:

            if not is_browser_alive(driver):
                log.warning("⚠️ Browser closed manually.")
                break

            prop_name = prop_data.get("name")
            prop_code = prop_data.get("code")
            custom_path = prop_data.get("save_path")
            pms_type = prop_data.get("pms", "").strip().lower()
            batch_option = prop_data.get("batch_option", "no").strip().lower()

            # 🔥 NEW — AME DATES
            last_ame_date = prop_data.get("last_ame_date", "")
            current_ame_date = prop_data.get("current_ame_date", "")

            if pms_type != "yardi":
                continue

            log.info("=================================================")
            log.info(f"🏢 Processing: {prop_name} ({prop_code})")
            log.info("=================================================")

            folder_path = create_property_folder(
                prop_name,
                prop_code,
                custom_path
            )

            document, word_path = create_word_document(
                prop_name,
                folder_path
            )


            # ----------------------------------------------------
            # Set Property Context
            # ----------------------------------------------------
            set_property(driver, wait, prop_code)

            # ----------------------------------------------------
            # Build Execution Pipeline
            # ----------------------------------------------------
            pipeline = [


                #("Daily Activity",
                 #lambda: add_formatted_section_to_word(
                     #driver, wait, document, "daily_activity"
                 #)),

                ("Daily Activity AME",
                 lambda: check_move_activity_current_ame(
                     driver, wait, document, current_ame_date
                 )),

                ("Convert to Prepay (YES)",
                 lambda: click_convert_to_prepay(
                     driver, wait, prop_code, batch_option
                 ) if batch_option == "yes" else None),


                ("Reset Datamart",
                 lambda: click_datamart_reset(driver, wait)),

                ("Automated AME Reports",
                 lambda: click_automated_ame_reports(driver, wait)),

                ("Credit Card Dashboard",
                 lambda: [
                     open_credit_card_dashboard_and_capture(driver, wait, prop_code),
                     add_formatted_section_to_word(driver, wait, document, "credit_card_dashboard")
                 ]),

                ("Payments Dashboard",
                 lambda: [
                     open_payments_dashboard_and_capture(driver, wait, prop_code),
                     add_formatted_section_to_word(driver, wait, document, "payment_dashboard")
                 ]),

                ("Transaction Register",
                 lambda: [
                     open_transaction_register_analytics(driver, wait, prop_code),
                     add_formatted_section_to_word(driver, wait, document, "transaction_register")
                 ]),

                ("NSF Register",
                 lambda: [
                     open_nsf_register(driver, wait),
                     add_formatted_section_to_word(driver, wait, document, "nsf_payments")
                 ]),


                ("Convert to Prepay (NO)",
                 lambda: [
                     click_convert_to_prepay(
                         driver, wait, prop_code, batch_option
                     ),
                     add_formatted_section_to_word(
                         driver, wait, document, "apply_open_credits"
                     )
                 ] if batch_option == "no" else None),


                ("Write Off",
                 lambda: run_residential_ar_analytics(
                     driver, wait, prop_code, document
                 )),

                ("Security Deposit",
                 lambda: run_security_deposit_full(
                     driver, wait, prop_code, prop_name, folder_path, document
                 )),

                ("Payable Aging",
                 lambda: [
                     run_payable_analytics_aging(driver, wait, prop_code, document),
                     add_formatted_section_to_word(driver, wait, document, "payable_aging")
                 ]),

                ("Gross Potential Rent",
                 lambda: run_gpr_full(
                     driver, wait, prop_code, prop_name, folder_path, document
                 )),

                ("GPR Download",
                 lambda: run_gpr_activity(driver, wait, prop_code, prop_name, folder_path)),
                
                ("GPR Check",
                 lambda: review_gpr_report(file_path, document)),

                ("Trial Balance",
                 lambda: run_trial_balance_activity(
                     driver, wait, prop_code, prop_name, folder_path
                 )),
                ("Step 16",
                 lambda: step16(
                     driver, wait, prop_code, prop_name, folder_path
                 )),
                
                ("Tie-Out",
                 lambda: run_tie_out(prop_name, folder_path)),
                
                ("Helper",
                 lambda: create_helper_sheet(os.path.join(folder_path, f"{prop_name} Tie-Out.xlsx"))),
            
                ("Formula",
                 lambda: write_tieout_formulas(os.path.join(folder_path, f"{prop_name} Tie-Out.xlsx"))),

                
            ]

            # ----------------------------------------------------
            # Execute Steps
            # ----------------------------------------------------
            execute_pipeline(driver, pipeline, steps_to_run, steps_to_skip)

            # ----------------------------------------------------
            # Save Word Document
            # ----------------------------------------------------
            document.save(word_path)
            log.info(f"💾 Word saved: {word_path}")

            # ----------------------------------------------------
            # Return to Home
            # ----------------------------------------------------
            go_home(driver, wait)

        log.info("🎯 All properties processed successfully.")

    except Exception as e:
        log.exception(f"🚨 Global automation failure: {e}")

    finally:
        # --------------------------------------------------------
        # OPTIONAL: Close browser
        # --------------------------------------------------------

        # If you want browser to close automatically:
        # driver.quit()

        # If you are debugging and want browser to remain open:
        log.info("🏁 Script execution finished.")

import os
import shutil
import logging
import time
import win32com.client as win32

log = logging.getLogger(__name__)


def find_template():

    user_home = os.path.expanduser("~")

    paths = [
        os.path.join(user_home, "Desktop", "Tie-out.xlsx"),
        os.path.join(user_home, "OneDrive - Greystar", "Desktop", "Tie-out.xlsx")
    ]

    for path in paths:
        if os.path.exists(path):
            log.info(f"📄 Tie-Out template found: {path}")
            return path

    log.error("❌ Tie-Out template not found.")
    return None


def run_tie_out(prop_name, folder_path):

    template_path = find_template()

    if not template_path:
        return

    new_file = os.path.join(folder_path, f"{prop_name} Tie-Out.xlsx")

    shutil.copy(template_path, new_file)

    log.info(f"📁 Tie-Out template copied to: {new_file}")

    reports = {
        "GPR": os.path.join(folder_path, f"{prop_name}_GPR.xlsx"),
        "Security Deposit": os.path.join(folder_path, f"{prop_name}_Security_Deposit_Activity.xlsx"),
        "AR": os.path.join(folder_path, f"{prop_name}_Financial Aged Receivable_Step 16.xlsx"),
        "TB": os.path.join(folder_path, f"{prop_name}_Trial_Balance.xlsx")
    }

    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:

        tieout_wb = excel.Workbooks.Open(new_file)

        for sheet_name, file_path in reports.items():

            if not os.path.exists(file_path):
                log.warning(f"⚠ Missing report: {file_path}")
                continue

            log.info(f"📊 Importing {sheet_name}")

            time.sleep(1)  # ensure file fully written

            try:
                source_wb = excel.Workbooks.Open(file_path)

                source_ws = source_wb.Worksheets(1)

                source_ws.Copy(After=tieout_wb.Sheets(tieout_wb.Sheets.Count))

                new_ws = tieout_wb.Sheets(tieout_wb.Sheets.Count)
                new_ws.Name = sheet_name

                source_wb.Close(False)

                log.info(f"✅ {sheet_name} imported")

            except Exception as e:
                log.warning(f"⚠ Failed importing {sheet_name}: {e}")

        tieout_wb.Save()
        tieout_wb.Close()

    finally:

        excel.Quit()

    log.info(f"🎯 Tie-Out file created successfully for {prop_name}")

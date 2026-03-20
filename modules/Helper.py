import logging
import win32com.client as win32

log = logging.getLogger(__name__)


def create_helper_sheet(tieout_file):

    log.info("▶ Creating Helper sheet")

    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:

        wb = excel.Workbooks.Open(tieout_file)

        # -------------------------------------------------
        # Get or create Helper sheet
        # -------------------------------------------------

        try:
            ws_helper = wb.Worksheets("Helper")
            ws_helper.Cells.Clear()
            log.info("♻ Helper sheet cleared")
        except:
            ws_helper = wb.Worksheets.Add()
            ws_helper.Name = "Helper"
            log.info("📄 Helper sheet created")

        ws_ar = wb.Worksheets("AR")
        ws_sec = wb.Worksheets("Security Deposit")
        ws_gpr = wb.Worksheets("GPR")

        # -------------------------------------------------
        # AR block (start checking row 6)
        # -------------------------------------------------

        row = 6
        last_row = row

        while ws_ar.Cells(row, 7).Value not in (None, ""):
            last_row = row
            row += 1

        for i in range(7):

            val = ws_ar.Cells(last_row, 7 + i).Value

            ws_helper.Cells(7 + i, 2).Value = val
            ws_helper.Cells(7 + i, 3).Formula = f"='AR'!{ws_ar.Cells(last_row, 7+i).Address}"

        log.info("✅ AR Helper block written")

        # -------------------------------------------------
        # Security Deposit (start row 6)
        # -------------------------------------------------

        row = ws_sec.Cells(ws_sec.Rows.Count, 9).End(-4162).Row

        ws_helper.Range("B18").Value = ws_sec.Cells(row, 9).Value
        ws_helper.Range("C18").Formula = f"='Security Deposit'!{ws_sec.Cells(row,9).Address}"

        ws_helper.Range("B19").Value = ws_sec.Cells(row, 10).Value
        ws_helper.Range("C19").Formula = f"='Security Deposit'!{ws_sec.Cells(row,10).Address}"

        log.info("✅ Security Deposit Helper block written")

        # -------------------------------------------------
        # GPR block (start row 7)
        # -------------------------------------------------

        row = 7
        while ws_gpr.Cells(row, 10).Value not in (None, ""):
            row += 1

        row -= 1

        ws_helper.Range("B22").Value = ws_gpr.Cells(row, 10).Value
        ws_helper.Range("C22").Formula = f"='GPR'!{ws_gpr.Cells(row,10).Address}"

        ws_helper.Range("B23").Value = ws_gpr.Cells(row, 12).Value
        ws_helper.Range("C23").Formula = f"='GPR'!{ws_gpr.Cells(row,12).Address}"

        log.info("✅ GPR Helper block written")

        wb.Save()
        wb.Close()

        log.info("🎯 Helper sheet created successfully")

    except Exception as e:

        log.error(f"❌ Helper creation failed: {e}")

    finally:

        excel.Quit()

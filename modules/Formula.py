import win32com.client as win32
import logging

log = logging.getLogger(__name__)


def write_tieout_formulas(tieout_file):

    log.info("▶ Writing Tie-Out formulas")

    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:

        wb = excel.Workbooks.Open(tieout_file)

        try:
            ws = wb.Worksheets("Tie-out")
        except:
            log.error("❌ Tie-out sheet not found")
            wb.Close(False)
            return

        formulas = {

            "N6": "=XLOOKUP(A6,TB!A:A,TB!F:F,0)+XLOOKUP(B6,TB!A:A,TB!F:F,0)",
            "N7": "=XLOOKUP(A7,TB!A:A,TB!F:F,0)+XLOOKUP(B7,TB!A:A,TB!F:F,0)",

            "N9": "=Helper!B7",
            "N10": "=Helper!B12",
            "N11": "=N6-N9",
            "N12": "=N7-N10",

            "N15": "=XLOOKUP(A15,TB!A:A,TB!F:F,0)*-1+XLOOKUP(B15,TB!A:A,TB!F:F,0)*-1",
            "N16": "=XLOOKUP(A16,TB!A:A,TB!F:F,0)*-1+XLOOKUP(B16,TB!A:A,TB!F:F,0)*-1",
            "N17": "=XLOOKUP(A17,TB!A:A,TB!F:F,0)*-1+XLOOKUP(B17,TB!A:A,TB!F:F,0)*-1",
            "N18": "=XLOOKUP(A18,TB!A:A,TB!F:F,0)*-1+XLOOKUP(B18,TB!A:A,TB!F:F,0)*-1",
            "N19": "=XLOOKUP(A19,TB!A:A,TB!F:F,0)*-1+XLOOKUP(B19,TB!A:A,TB!F:F,0)*-1",

            
            "N21": "=Helper!B18",
            "N22": "=Helper!B19",
            "N23": "=SUM(N15:N19)-SUM(N21:N22)",

            "N26": "=XLOOKUP(A26,TB!A:A,TB!E:E,0)+XLOOKUP(B26,TB!A:A,TB!E:E,0)",
            "N27": "=XLOOKUP(A27,TB!A:A,TB!E:E,0)+XLOOKUP(B27,TB!A:A,TB!E:E,0)",
            "N28": "=XLOOKUP(A28,TB!A:A,TB!E:E,0)+XLOOKUP(B28,TB!A:A,TB!E:E,0)",
            "N29": "=XLOOKUP(A29,TB!A:A,TB!E:E,0)+XLOOKUP(B29,TB!A:A,TB!E:E,0)",

            "N31": "=Helper!B22",
            "N32": "=SUM(N25:N29)-N31",

            "N38": "=Helper!B23",
            "N39": "=SUM(N35:N36)+N38"
        }

        for cell, formula in formulas.items():
            ws.Range(cell).Formula = formula

        ws.Activate()
        ws.Range("N3").Select()
        wb.Save()
        wb.Close()

        log.info("✅ Tie-Out formulas written successfully")

    except Exception as e:

        log.error(f"❌ Failed to write Tie-Out formulas: {e}")

    finally:

        excel.Quit()

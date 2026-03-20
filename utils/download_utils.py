# ============================================================
# IMPORTS - download_utils.py
# ============================================================

import os
import time
import shutil
import logging
import logging
log = logging.getLogger(__name__)
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.shared import RGBColor

from io import BytesIO
from docx.shared import Inches

def wait_for_download_and_rename(
    download_dir,
    files_before,
    target_folder,
    base_filename,
    extension=".xlsx",
    timeout=90
):

    log.info("⏳ Waiting for Excel file download...")

    end_time = time.time() + timeout

    while time.time() < end_time:

        time.sleep(1)

        files_after = set(os.listdir(download_dir))

        new_files = [
            f for f in files_after - files_before
            if f.lower().endswith((".xlsx", ".xls"))
            and not f.startswith("~$")
        ]

        if new_files:

            downloaded_file = new_files[0]
            full_path = os.path.join(download_dir, downloaded_file)

            # wait until file stable
            size1 = os.path.getsize(full_path)
            time.sleep(2)
            size2 = os.path.getsize(full_path)

            if size1 == size2:

                final_path = os.path.join(
                    target_folder,
                    f"{base_filename}{extension}"
                )

                counter = 1
                while os.path.exists(final_path):
                    final_path = os.path.join(
                        target_folder,
                        f"{base_filename} ({counter}){extension}"
                    )
                    counter += 1

                shutil.move(full_path, final_path)

                log.info(f"📁 File moved to: {final_path}")
                return final_path

    raise TimeoutError("❌ Excel download timed out.")


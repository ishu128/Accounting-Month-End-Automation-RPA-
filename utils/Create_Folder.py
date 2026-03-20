# ============================================================
# IMPORTS - Create_Folder.py
# ============================================================

import os
from datetime import datetime
import logging
log = logging.getLogger(__name__)
def create_property_folder(property_name, property_code, custom_path=""):
    """
    Creates folder:
    If F column path exists → use that
    Else → create in Downloads:
        PropertyName_PropertyCode_TodayDate

    Always unique (adds _1, _2 if exists)
    """

    today_str = datetime.today().strftime("%Y%m%d")

    if custom_path:
        base_folder = custom_path
    else:
        downloads = os.path.join(os.path.expanduser("~"), "Downloads")
        base_folder = os.path.join(
            downloads,
            f"{property_name}_{property_code}_{today_str}"
        )

    final_folder = base_folder
    counter = 1

    while os.path.exists(final_folder):
        final_folder = f"{base_folder}_{counter}"
        counter += 1

    os.makedirs(final_folder)

    log.info(f"📁 Save folder created: {final_folder}")
    return final_folder

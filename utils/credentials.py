# ============================================================
# READ CREDENTIALS FROM JSON FILE
# ============================================================


# ============================================================
# IMPORTS - credentials.py
# ============================================================

import os
import json
import logging
log = logging.getLogger(__name__)
def read_credentials():
    """
    Reads login credentials from secure JSON file.
    """

    config_path = r"C:\ProgramData\Packages_PY\credentials.json"

    if not os.path.exists(config_path):
        raise FileNotFoundError("Credentials JSON not found.")

    with open(config_path, "r") as f:
        data = json.load(f)

    username = data.get("id")
    password = data.get("password")

    if not username or not password:
        raise ValueError("Credentials JSON missing id or password.")

    return username, password


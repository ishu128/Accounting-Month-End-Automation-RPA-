# ============================================================
# IMPORTS - Creat_Word.py
# ============================================================

import os
from docx import Document
import logging
log = logging.getLogger(__name__)

def create_word_document(property_name, folder_path):
    """
    Creates Word file:
    PropertyName_Observations.docx
    """

    file_name = f"{property_name}_Observations.docx"
    full_path = os.path.join(folder_path, file_name)

    document = Document()
    document.add_heading(f"{property_name} - Observations", level=1)

    log.info(f"📝 Word document initialized: {file_name}")

    return document, full_path

import os


def get_tieout_file(prop_name, folder_path):
    """
    Returns full path of Tie-Out workbook
    """
    return os.path.join(folder_path, f"{prop_name} Tie-Out.xlsm")

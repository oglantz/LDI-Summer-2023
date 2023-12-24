from pathlib import Path


def get_xl_files():
    return list(Path(r"C:\Users\aweso\OneDrive\Desktop\test").glob("*.xlsx"))

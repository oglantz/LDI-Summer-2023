from pathlib import Path


def get_xl_files():
    return list(Path(r"C:\Users\aweso\OneDrive\Desktop\test").glob("*.xlsx"))



# if __name__ == "__main__":
#     print(get_xl_files())

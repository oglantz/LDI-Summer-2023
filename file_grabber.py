from pathlib import Path

# def _get_files():
#     files = []
#     p = Path()
#
#     for i in p.iterdir():
#         if i.is_file():
#             files.append(i)
#     return files

def get_xl_files():
    return list(Path().glob("*.xlsx"))

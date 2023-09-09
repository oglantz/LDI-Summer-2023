from pathlib import Path

def _get_files():
    files = []
    p = Path()

    for i in p.iterdir():
        if i.is_file():
            files.append(i)
    return files

def get_xl_files():
    xl_files = []
    files = _get_files()

    for file in files:
        if file.suffix == '.xlsx':
            xl_files.append(file)
    return xl_files

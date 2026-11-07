import os
from pathlib import Path
import sys
import tempfile


def setup(root):
    root = Path(root).resolve()
    temp = root / 'temp';temp.mkdir(parents=True, exist_ok=True)
    os.environ['TEMP'] = os.environ['TMP'] = str(temp)
    os.environ['SQLITE_TMPDIR'] = str(temp)
    tempfile.tempdir = str(temp)
    home = root
    while not (home / 'runtime/python').is_dir() and home.parent != home:
        home = home.parent
    deps = home / 'runtime/python'
    for path in [deps, deps/'win32', deps/'win32/lib', deps/'Pythonwin', deps/'pywin32_system32']:
        sys.path.insert(0, str(path))
    if os.name == 'nt' and hasattr(os, 'add_dll_directory'):
        global dll_directory
        dll_directory = os.add_dll_directory(str(deps/'pywin32_system32'))
    sys.path.insert(0, str(home/'code'))

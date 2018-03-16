import sys
from cx_Freeze import setup, Executable
import itertools
itertools.imap = lambda *args, **kwargs: list(map(*args, **kwargs))

# Dependencies are automatically detected, but it might need fine tuning.
build_exe_options = {"packages": ["os","gensim","docx","konlpy","copy","sys","numpy","idna","html","jpype"],"includes":["xlout"],"excludes":["scipy.spatial.cKDTree"]}


# GUI applications require a different base on Windows (the default is for a
# console application).
base = None
if sys.platform == "win32":
    base = "Win32GUI"


setup(  name = "add parser",
    version = "1.0",
    description = "add parser",
    options = {"build_exe": build_exe_options},
    executables = [Executable("docx_read.py", base = base)])
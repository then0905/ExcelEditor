# -*- mode: python ; coding: utf-8 -*-
import customtkinter
import os

ctk_path = os.path.dirname(customtkinter.__file__)

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[(ctk_path, 'customtkinter/'), ('icon.ico', '.')],
    hiddenimports=['pandas', 'openpyxl', 'numpy', 'PIL'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'scipy', 'matplotlib', 'IPython', 'jupyter',
        'notebook', 'pytest', 'unittest',
        'numpy.f2py', 'numpy.distutils', 'numpy.testing',
        'pandas.tests', 'pandas.io.formats.style',
        'tkinter.test', 'lib2to3',
        'multiprocessing', 'concurrent',
        'email', 'html', 'http', 'xmlrpc',
        'pydoc', 'doctest', 'argparse',
        'logging.handlers', 'logging.config',
    ],
    noarchive=False,
    optimize=0,
)

# 移除不必要的大型 DLL（numpy MKL/OpenBLAS 測試等）
import re
_exclude_patterns = re.compile(
    r'(mkl_|libopenblas|libiomp|libblas|liblapack|libgfortran|libquadmath|vcruntime)'
    r'.*\.(dll|so)', re.IGNORECASE
)
a.binaries = [b for b in a.binaries if not _exclude_patterns.search(b[0])]

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='ExcelEditor',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico',
)

# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec for School Timetable Generator
# Usage: pyinstaller build_windows.spec

block_cipher = None

a = Analysis(
    ['timetable_app.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        'PIL._tkinter_finder',
        'PIL.Image',
        'PIL.ImageTk',
        'openpyxl',
        'openpyxl.styles',
        'openpyxl.utils',
        'openpyxl.drawing.image',
        'docx',
        'docx.shared',
        'docx.enum.text',
        'docx.enum.table',
        'docx.oxml.ns',
        'docx.oxml',
        'reportlab',
        'reportlab.lib.pagesizes',
        'reportlab.lib.colors',
        'reportlab.lib.units',
        'reportlab.platypus',
        'reportlab.lib.styles',
        'reportlab.lib.enums',
        'reportlab.pdfgen',
        'lxml',
        'lxml._elementpath',
        'et_xmlfile',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='SchoolTimetableGenerator',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,      # No console window
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
)

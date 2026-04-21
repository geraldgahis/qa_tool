# -*- mode: python ; coding: utf-8 -*-

a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('credentials.json', '.'), 
        ('favicon.ico', '.')
    ],
    hiddenimports=[
        'gspread', 
        'google.oauth2.service_account',
        'ttkbootstrap',
        'playwright',       # Add this
        'playwright.drivers' # Add this
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='QA Tool', # You can rename your EXE here
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False, # Set to True if you need to see errors during testing
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['favicon.ico'], # File icon for Windows Explorer
)
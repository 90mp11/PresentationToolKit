# -*- mode: python ; coding: utf-8 -*-


block_cipher = None


a = Analysis(
    ['main.py'],
    pathex=['.'],
    binaries=[],
    datas=[
        ('magic.png', '.'),
        ('LOGO.png', '.'),
        ('templates/_template.pptx', 'templates'),
        ('utilities/builder.py', 'utilities'),
        ('utilities/constants.py', 'utilities'),
        ('utilities/data_utils.py', 'utilities'),
        ('utilities/graph_utils.py', 'utilities'),
        ('utilities/presentation_utils.py', 'utilities'),
        ("C:\\Users\\MattProctor\\OneDrive - Cityfibre Limited\\Tools_Resources\\ProjectReportCreator\\venv\\Lib\\site-packages\\tkcalendar", 'tkcalendar'),
        ("C:\\Users\\MattProctor\\OneDrive - Cityfibre Limited\\Tools_Resources\\ProjectReportCreator\\venv\\Lib\\site-packages\\babel", 'babel'),
    ],
    hiddenimports=['tkcalendar'],
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
    name='PRG_v2',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    icon=['test.ico'],
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

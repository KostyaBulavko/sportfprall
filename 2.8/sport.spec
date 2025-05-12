# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['sport.py'],
    pathex=['E:\\2. Басай\\00. Заходи 2025\\main prog'],
    binaries=[],
    datas=[
        ('ШАБЛОН_кошторис.xlsx', '.'),  # додаємо шаблон
        ('icon.ico', '.')               # додаємо іконку
    ],
    hiddenimports=[
        'tkinter.filedialog',
        'tkinter.messagebox',
        'customtkinter',
        'openpyxl',
        'win32com.client',
    ],
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='sport',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # <-- щоб не з'являлось чорне вікно
    icon='icon.ico'  # <-- шлях до іконки
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='sport'
)

# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec file para PD7Lab Extractor PDF

block_cipher = None

# Arquivos de dados que devem ser incluídos no executável
added_files = [
    ('pd7.png', '.'),              # Logo clara (tema light)
    ('pd7lab-dark.jpeg', '.'),     # Logo escura (tema dark)
    ('pd7-escudo.ico', '.'),       # icone
]

a = Analysis(
    ['get_proof.py'],
    pathex=[],
    binaries=[],
    datas=added_files,
    hiddenimports=[
        'PIL._tkinter_finder',     # Necessário para PIL/Pillow com tkinter
        'PIL.Image',
        'PIL.ImageTk',
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
    name='PD7Lab_ExtractorPDF',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,                  # Sem janela de console (GUI apenas)
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='pd7-escudo.ico',                 # Ícone do executável
)

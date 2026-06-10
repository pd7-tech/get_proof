# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec file para PD7Lab Extractor PDF

from PyInstaller.utils.hooks import collect_all

block_cipher = None

# Coletar todos os assets do customtkinter (temas, fontes, imagens)
ctk_datas, ctk_binaries, ctk_hiddenimports = collect_all('customtkinter')

# Coletar assets do pdfplumber/pdfminer (arquivos de charset)
plumber_datas, plumber_binaries, plumber_hiddenimports = collect_all('pdfplumber')

# Arquivos de dados que devem ser incluídos no executável
added_files = [
    ('pd7.png', '.'),              # Logo clara (tema light)
    ('pd7lab-dark.jpeg', '.'),     # Logo escura (tema dark)
    ('pd7-escudo.ico', '.'),       # icone
] + ctk_datas + plumber_datas

a = Analysis(
    ['get_proof.py'],
    pathex=[],
    binaries=ctk_binaries + plumber_binaries,
    datas=added_files,
    hiddenimports=[
        'PIL._tkinter_finder',
        'PIL.Image',
        'PIL.ImageTk',
        'customtkinter',
        'darkdetect',
    ] + ctk_hiddenimports + plumber_hiddenimports,
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
    upx=False,                      # UPX desativado para evitar falsos positivos de antivírus
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,                  # Sem janela de console (GUI apenas)
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch='x86_64',           # Forçar 64-bit (compatível com Windows 10/11 x64)
    codesign_identity=None,
    entitlements_file=None,
    icon='pd7-escudo.ico',                 # Ícone do executável
)

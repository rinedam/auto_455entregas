# -*- mode: python ; coding: utf-8 -*-

import os
import customtkinter

block_cipher = None

# --- Análise do Projeto ---
a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[
        # Adiciona o arquivo de credenciais para a pasta raiz do executável.
        ('credenciais.env', '.'),
        
        # Adiciona os arquivos de tema do CustomTkinter.
        (os.path.join(os.path.dirname(customtkinter.__file__), "assets"), "customtkinter/assets")
    ],
    hiddenimports=[
        # Bibliotecas que o PyInstaller pode não encontrar sozinho.
        # A referência ao 'webdriver_manager' foi REMOVIDA.
        'schedule',
        'pandas',
        'selenium',
        'dotenv'
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

# --- Criação do Pacote Python (PYZ) ---
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# --- Criação do Executável (EXE) ---
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='PainelDeControle',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,  # Mantém a janela do console oculta.
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

# --- Coleta dos Arquivos Finais (Modo --onedir) ---
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='PainelDeControle',
)
# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller 打包配置：生成 Windows 可执行程序。"""

from PyInstaller.utils.hooks import collect_all

block_cipher = None

# 收集 customtkinter 的主题与资源文件
ctk_datas, ctk_binaries, ctk_hiddenimports = collect_all("customtkinter")

a = Analysis(
    ["src/main.py"],
    pathex=["src"],
    binaries=ctk_binaries,
    datas=ctk_datas,
    hiddenimports=[
        "openpyxl",
        "pandas",
        *ctk_hiddenimports,
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
    [],
    exclude_binaries=True,
    name="客户产品Excel汇总工具",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="客户产品Excel汇总工具",
)

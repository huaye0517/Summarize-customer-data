# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller 打包配置：在 macOS 上生成 .app 应用程序。"""

from PyInstaller.utils.hooks import collect_all

block_cipher = None

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
    argv_emulation=True,
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

app = BUNDLE(
    coll,
    name="客户产品Excel汇总工具.app",
    icon=None,
    bundle_identifier="com.summarize.customer-excel",
)

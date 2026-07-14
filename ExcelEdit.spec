# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller 打包配置。

这里只描述打包时需要收集的资源和最终 EXE 名称；不会在运行程序时被导入。
"""

from pathlib import Path

from PyInstaller.utils.hooks import collect_data_files

# 图标不是必需资源，缺少时仍允许正常打包。
icon_path = Path('assets/nailong_search_master.ico')
datas = []

# CustomTkinter 依赖主题和资源文件，打包时需要一起收集。
datas += collect_data_files('customtkinter')
if icon_path.exists():
    datas += [(str(icon_path), 'assets')]


# Analysis 定义入口脚本和需要一起打进包里的数据文件。
a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)

# PYZ 保存 Python 字节码，EXE 定义最终窗口程序。
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='奶龙检索大师',
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
    icon=str(icon_path) if icon_path.exists() else None,
)

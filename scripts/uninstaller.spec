# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller 打包配置：生成独立自包含卸载器 uninstaller.exe。

用法（项目根目录执行，需先打包 SuJian.spec 产出 dist/SuJian.exe）：
    pyinstaller scripts/uninstaller.spec --noconfirm

产物：dist\\uninstaller.exe（单文件、无控制台、仅标准库 + winreg，约 1MB 级）。
    不依赖用户机器是否安装 Python；由 setup.exe 内嵌并在安装时复制到安装目录，
    注册表 UninstallString 指向它，保证“应用和功能”卸载入口正确走卸载流程。
"""
import os
ROOT = os.path.dirname(os.path.abspath(SPECPATH))

a = Analysis(
    [os.path.join(ROOT, "scripts", "uninstaller.py")],
    pathex=[ROOT],
    binaries=[],
    datas=[],
    hiddenimports=["winreg"],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["PySide6"],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="uninstaller",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    console=False,               # 无控制台窗口（卸载器内部仍打印日志，GUI 隐藏）
    icon=[os.path.join(ROOT, "build_assets", "icon.ico")],
)

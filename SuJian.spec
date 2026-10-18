# -*- mode: python ; coding: utf-8 -*-
"""SuJian（素笺）PyInstaller 打包配置。

打包命令：
    pyinstaller SuJian.spec --clean --noconfirm

产物：
    dist/SuJian.exe      —— 单文件可执行程序（无控制台窗口）

说明：
    - onefile 单文件模式：所有代码、资源与 Qt 插件打包进单个 exe；
    - console=False：GUI 程序不显示黑色控制台窗口；
    - hiddenimports 覆盖办公库的动态子模块（openpyxl / python-docx / xlrd /
      pypdfium2 / PIL）；PDF 已裁剪为只读查看，不再打包 reportlab（创建/导出
      已移除），体积更小；
    - datas 内嵌 qtbase_zh_CN.qm / qt_zh_CN.qm 到 PySide6/translations，
      供 main.py 启动时加载实现全界面中文；
    - icon 使用 build_assets/icon.ico（绿色图标）。
"""

a = Analysis(
    ["main.py"],
    pathex=["."],
    binaries=[],
    datas=[],
    hiddenimports=[
        # GUI
        "PySide6.QtCore",
        "PySide6.QtGui",
        "PySide6.QtWidgets",
        # Excel（openpyxl 动态加载部分子模块）
        "openpyxl",
        "openpyxl.cell._writer",
        "openpyxl.styles",
        "openpyxl.utils.datetime",
        # Word（python-docx 包名 docx）
        "docx",
        "docx.opc",
        "docx.document",
        # 老版 .xls 读取
        "xlrd",
        "xlrd.biffh",
        "xlrd.formula",
        # PDF 渲染（只读查看）
        "pypdfium2",
        # 图像辅助（图标 / 插入图片）
        "PIL",
        "PIL.Image",
        "PIL.ImageDraw",
        "PIL.ImageFont",
        "PIL.ImageQt",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

import os

# 修复：PyInstaller 的 PySide6 hook 会自动把整个 translations 目录收集进 a.datas
#（约 200 个语言 qm）。此前代码在 Analysis 后把单个 qm 以 dest="PySide6/translations"
#（目录名，缺少文件名）追加进 TOC，运行时 onefile 解压时该条目与真实目录冲突，
# 导致 "Failed to extract PySide6\\translations\\qt_ar.qm: failed to open target file!"。
# 现改为：先过滤多余语言 qm，仅保留中文化所需两个文件；缺失时再以完整目标文件名追加。
_KEEP_QM = {"qtbase_zh_CN.qm", "qt_zh_CN.qm"}
_FILTERED = []
for _item in a.datas:
    _dest = _item[0].replace("\\", "/")
    if _dest.startswith("PySide6/translations/"):
        if os.path.basename(_dest) in _KEEP_QM:
            _FILTERED.append(_item)  # 保留中文翻译（hook 已收集）
    else:
        _FILTERED.append(_item)
a.datas = _FILTERED

_P6_T = None
try:
    import site
    for _p in site.getsitepackages():
        _c = os.path.join(_p, "PySide6", "translations")
        if os.path.isdir(_c):
            _P6_T = _c
            break
except Exception:
    pass
if _P6_T:
    _existing = {_i[0].replace("\\", "/") for _i in a.datas}
    for _f in ("qtbase_zh_CN.qm", "qt_zh_CN.qm"):
        _dest = "PySide6/translations/" + _f
        if _dest not in _existing:
            _fp = os.path.join(_P6_T, _f)
            if os.path.exists(_fp):
                # TOC 三元组：(归档内目标完整路径, 源文件, typecode)
                a.datas.append((_dest, _fp, "DATA"))

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="SuJian",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    console=False,
    icon=["build_assets/icon.ico"],
)

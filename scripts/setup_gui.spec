# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller 打包配置：生成单文件图形化安装程序 setup.exe。

用法（项目根目录执行）：
    pyinstaller scripts/setup_gui.spec --clean --noconfirm

产物：dist\\setup.exe（单文件、无控制台、带绿色图标）。
    内嵌数据：dist/SuJian.exe（主程序，62MB 级）、build_assets/icon.ico（窗口图标）、
    dist/uninstaller.exe（自包含卸载器，安装时复制到安装目录并写入 UninstallString，
    避免卸载入口错误依赖用户机器上的 Python 而“打开的是安装程序”）。
    运行时 PyInstaller 会把内嵌数据解压到临时目录（sys._MEIPASS），
    setup_gui.py 从中读取 dist/SuJian.exe 并复制到用户选择的安装目录。
"""
import site
from PyInstaller.utils.hooks import collect_data_files, collect_dynamic_libs, collect_submodules

# 体积优化：不再 collect_all 整个 PySide6（旧方式会把 QML/WebEngine/多媒体等全部打进安装向导，
# 仅向导部分就约 260MB）。改为只收集 QtCore/QtGui/QtWidgets 三个实际用到的模块，
# 手动补齐运行必需的 Qt 平台/样式/ico 插件与中文翻译，并在 Analysis.excludes 排除无关模块。
datas = []
binaries = []
hiddenimports = []
for _mod in ("PySide6.QtCore", "PySide6.QtGui", "PySide6.QtWidgets"):
    for _d in collect_data_files(_mod):
        if "translations" not in _d[0].replace("\\", "/"):
            datas.append(_d)
    binaries += collect_dynamic_libs(_mod)
    hiddenimports += collect_submodules(_mod)

# 定位 PySide6 安装目录，手动补充运行必需的 Qt 插件
_P6 = None
for _p in site.getsitepackages():
    _cand = os.path.join(_p, "PySide6")
    if os.path.isdir(_cand):
        _P6 = _cand
        break
if _P6:
    for _rel in ("plugins/platforms/qwindows.dll",
                 "plugins/styles/qmodernwindowsstyle.dll",
                 "plugins/imageformats/qico.dll"):
        _f = os.path.join(_P6, _rel)
        if os.path.exists(_f):
            datas.append((_f, "PySide6/" + os.path.dirname(_rel).replace("\\", "/")))
    # 仅保留中文界面翻译（两个小文件），替代整包几百个 .qm
    for _f in ("qtbase_zh_CN.qm", "qt_zh_CN.qm"):
        _fp = os.path.join(_P6, "translations", _f)
        if os.path.exists(_fp):
            datas.append((_fp, "PySide6/translations"))

import os
ROOT = os.path.dirname(os.path.abspath(SPECPATH))

a = Analysis(
    [os.path.join(ROOT, "scripts", "setup_gui.py")],
    pathex=[ROOT],
    binaries=binaries,
    datas=datas + [
        (os.path.join(ROOT, "dist", "SuJian.exe"), "dist"),       # 内嵌主程序，随 setup.exe 分发
        (os.path.join(ROOT, "dist", "uninstaller.exe"), "dist"),  # 内嵌自包含卸载器，安装时复制到安装目录
        (os.path.join(ROOT, "build_assets", "icon.ico"), "build_assets"),  # 向导窗口图标
    ],
    hiddenimports=hiddenimports + [
        "winreg",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # 安装向导用不到的 Qt 模块，全部排除以大幅缩小体积
        "PySide6.QtQml", "PySide6.QtQuick", "PySide6.QtQuick3D",
        "PySide6.QtQuickControls2", "PySide6.QtQuickWidgets",
        "PySide6.QtWebEngineCore", "PySide6.QtWebEngineWidgets", "PySide6.QtWebEngineQuick",
        "PySide6.QtWebChannel", "PySide6.QtWebSockets", "PySide6.QtWebView",
        "PySide6.QtMultimedia", "PySide6.QtMultimediaWidgets", "PySide6.QtSpatialAudio",
        "PySide6.Qt3DCore", "PySide6.Qt3DRender", "PySide6.Qt3DInput",
        "PySide6.Qt3DLogic", "PySide6.Qt3DAnimation", "PySide6.Qt3DExtras",
        "PySide6.QtCharts", "PySide6.QtDataVisualization", "PySide6.QtGraphs",
        "PySide6.QtNetwork", "PySide6.QtSql", "PySide6.QtSvg", "PySide6.QtSvgWidgets",
        "PySide6.QtPdf", "PySide6.QtPdfWidgets",
        "PySide6.QtTextToSpeech", "PySide6.QtSensors", "PySide6.QtSerialPort",
        "PySide6.QtPositioning", "PySide6.QtLocation", "PySide6.QtRemoteObjects",
        "PySide6.QtScxml", "PySide6.QtStateMachine", "PySide6.QtTest",
        "PySide6.QtDesigner", "PySide6.QtHelp", "PySide6.QtOpenGL",
        "PySide6.QtOpenGLWidgets", "PySide6.QtConcurrent", "PySide6.QtXml",
        "PySide6.QtXmlPatterns", "PySide6.QtDBus", "PySide6.QtBluetooth", "PySide6.QtNfc",
        "PySide6.QtCoreTranslations", "PySide6.QtGuiTranslations", "PySide6.QtWidgetsTranslations",
    ],
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
    name="setup",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    console=False,               # 无控制台窗口（GUI 安装程序）
    icon=[os.path.join(ROOT, "build_assets", "icon.ico")],
)

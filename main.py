# -*- coding: utf-8 -*-
"""SuJian（素笺）程序启动入口。

运行方式：
    python main.py
"""
import os
import sys
import traceback

from PySide6.QtWidgets import QApplication, QMessageBox
from PySide6.QtGui import QIcon
from PySide6.QtCore import Qt

from app.main_window import MainWindow
from app.i18n import load_qt_translation, install as install_i18n


def _app_icon_path():
    """定位程序窗口图标（打包后位于 _MEIPASS/build_assets，开发模式在项目根）。"""
    base = getattr(sys, "_MEIPASS", None)
    if base:
        cand = os.path.join(base, "build_assets", "icon.ico")
        if os.path.isfile(cand):
            return cand
    cand = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                        "build_assets", "icon.ico")
    return cand if os.path.isfile(cand) else None


def _install_excepthook():
    """全局异常兜底：未捕获异常写入 error.log 并弹中文提示，避免静默闪退。"""
    log_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "error.log")

    def hook(exc_type, exc_value, exc_tb):
        try:
            with open(log_path, "a", encoding="utf-8") as f:
                f.write("=" * 60 + "\n")
                f.write("".join(traceback.format_exception(exc_type, exc_value, exc_tb)))
                f.write("\n")
        except Exception:
            pass
        try:
            app = QApplication.instance()
            if app is not None:
                msg = "".join(
                    traceback.format_exception_only(exc_type, exc_value)).strip()
                QMessageBox.critical(
                    None, "程序错误",
                    f"程序遇到未处理的错误，详情已记录到 error.log。\n\n{msg}")
        except Exception:
            pass

    sys.excepthook = hook


def _check_dependencies(app):
    """启动时自检可选依赖库，缺失时弹出中文提示（不阻断启动）。

    各编辑器内置了独立的缺失检测（带 pip 安装提示），这里的自检
    用于启动阶段统一告知，帮助用户一次看清缺了哪些库。
    """
    checks = (
        ("openpyxl", "openpyxl", "Excel 表格的新建 / 保存"),
        ("docx", "python-docx", "Word 文档的新建 / 保存"),
        ("pypdfium2", "pypdfium2", "PDF 的打开查看"),
    )
    missing = []
    for module, pip_name, usage in checks:
        try:
            __import__(module)
        except Exception:
            missing.append((pip_name, usage))

    if not missing:
        return

    lines = "\n".join(f"未安装 {pip_name} → 请运行：pip install {pip_name}"
                      for pip_name, _ in missing)
    box = QMessageBox()
    box.setIcon(QMessageBox.Icon.Warning)
    box.setWindowTitle("依赖库缺失提醒")
    box.setText("检测到以下依赖库未安装：\n\n"
                f"{lines}\n\n"
                "这些功能暂时不可用，但不影响软件启动。\n"
                "请在命令行（CMD / PowerShell）中逐条执行上述命令，\n"
                "安装完成后重启本软件即可正常使用对应功能。")
    box.addButton("知道了", QMessageBox.ButtonRole.AcceptRole)
    box.exec()


def main():
    # 全局异常兜底（写 error.log + 中文提示），防止任何未捕获异常静默闪退
    _install_excepthook()

    # 启用高分屏缩放
    QApplication.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    app = QApplication(sys.argv)
    app.setApplicationName("SuJian")
    app.setApplicationDisplayName("素笺")
    app.setOrganizationName("SuJian")

    # 窗口 / 任务栏图标：素笺绿色图标（开发与打包环境均可定位）
    _icon = QIcon(_app_icon_path())
    if not _icon.isNull():
        app.setWindowIcon(_icon)

    # 全界面中文：加载 Qt 中文翻译（qtbase_zh_CN.qm 等）+ 原生对话框按钮兜底
    load_qt_translation()
    install_i18n(app)

    # 启动时依赖自检（仅提示，不阻断）
    _check_dependencies(app)

    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()

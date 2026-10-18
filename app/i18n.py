# -*- coding: utf-8 -*-
"""全界面中文方案：Qt 中文翻译加载 + 原生对话框按钮兜底。

- load_qt_translation()：依次从 PyInstaller 临时目录、site-packages、
  项目本地目录查找 qtbase_zh_CN.qm / qt_zh_CN.qm 并加载到 QTranslator，
  覆盖 QMessageBox / QFileDialog / QColorDialog / QFontDialog /
  QInputDialog 等 Qt 原生对话框的全部标准按钮与提示文案；
- install(app)：安装全局事件过滤器，任何 QDialog 弹出时再次把标准英文
  按钮（OK / Cancel / Yes / No / Save / Open ...）逐个替换为中文，
  作为翻译文件缺失 / 未覆盖场景的兜底，保证任何界面不残留英文。
"""
import os
import sys

from PySide6.QtCore import QObject, QTranslator, QEvent, QTimer, QCoreApplication
from PySide6.QtWidgets import QDialog, QAbstractButton, QDialogButtonBox

#: Qt 标准按钮英文 -> 中文（含带 & 加速键前缀的形式）
_BTN_MAP = {
    "OK": "确定",
    "&OK": "确定",
    "Cancel": "取消",
    "&Cancel": "取消",
    "Close": "关闭",
    "&Close": "关闭",
    "Yes": "是",
    "&Yes": "是",
    "No": "否",
    "&No": "否",
    "Save": "保存",
    "&Save": "保存",
    "Don't Save": "不保存",
    "Discard": "放弃",
    "&Discard": "放弃",
    "Apply": "应用",
    "&Apply": "应用",
    "Reset": "重置",
    "&Reset": "重置",
    "Retry": "重试",
    "&Retry": "重试",
    "Abort": "中止",
    "&Abort": "中止",
    "Ignore": "忽略",
    "&Ignore": "忽略",
    "Help": "帮助",
    "&Help": "帮助",
    "Open": "打开",
    "&Open": "打开",
    "Choose": "选择",
    "&Choose": "选择",
    "Select Folder": "选择文件夹",
    "&Select Folder": "选择文件夹",
    "Save As": "另存为",
    "&Save As": "另存为",
    "Choose Font": "选择字体",
    "&Choose Font": "选择字体",
}

#: 静态持有 QTranslator 引用，防止被垃圾回收
_translators = []
_localizer = None


def _candidate_paths():
    """返回可能的 Qt 翻译文件目录（按优先级）。"""
    roots = []
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        roots.append(sys._MEIPASS)  # PyInstaller onefile 解压目录
    # 项目根目录（源码运行 / 打包后同目录 dist 场景）
    roots.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    try:
        import site
        roots.extend(site.getsitepackages())
    except Exception:
        pass
    seen = set()
    for r in roots:
        if r and r not in seen:
            seen.add(r)
            yield os.path.join(r, "PySide6", "translations")


def load_qt_translation():
    """加载 Qt 中文翻译，返回成功加载的翻译文件数量。"""
    loaded = 0
    for tdir in _candidate_paths():
        for name in ("qtbase_zh_CN", "qt_zh_CN"):
            path = os.path.join(tdir, name + ".qm")
            if not os.path.exists(path):
                continue
            tr = QTranslator()
            if tr.load(path):
                _translators.append(tr)
                QCoreApplication.installTranslator(tr)
                loaded += 1
    return loaded


def _translate_buttons(dialog):
    """递归替换对话框内所有按钮文案（QAbstractButton + QDialogButtonBox）。"""
    for btn in dialog.findChildren(QAbstractButton):
        text = btn.text().strip()
        target = _BTN_MAP.get(text)
        if target is not None:
            btn.setText(target)
    for box in dialog.findChildren(QDialogButtonBox):
        for btn in box.buttons():
            text = btn.text().strip()
            target = _BTN_MAP.get(text)
            if target is not None:
                btn.setText(target)


class _DialogLocalizer(QObject):
    """全局事件过滤器：对话框 Show 事件后延迟翻译按钮，不干扰用户交互。"""

    def eventFilter(self, obj, event):
        if event.type() == QEvent.Type.Show and isinstance(obj, QDialog):
            QTimer.singleShot(0, lambda d=obj: _translate_buttons(d))
        return False


def install(app):
    """安装对话框按钮本地化过滤器（重复安装自动去重）。"""
    global _localizer
    if _localizer is not None:
        return
    _localizer = _DialogLocalizer()
    app.installEventFilter(_localizer)

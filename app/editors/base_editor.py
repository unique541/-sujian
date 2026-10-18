# -*- coding: utf-8 -*-
"""编辑器基类：定义所有文档编辑器统一的接口。"""
from PySide6.QtCore import Signal
from PySide6.QtWidgets import QWidget


class BaseEditor(QWidget):
    """所有文档编辑器的公共基类。

    子类必须实现：doc_type / load / save / save_as / new_document。
    可选实现（主窗口会按 doc_type 反射调用）：find_next / replace_current / replace_all。
    """

    #: 修改状态变化信号（True 表示文档被修改过）
    modified_changed = Signal(bool)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.file_path = None       # 当前关联的文件路径，None 表示尚未保存
        self._modified = False

    # ---------- 必须实现的接口 ----------
    def doc_type(self):
        """返回文档类型：excel / word / text / pdf"""
        raise NotImplementedError

    def load(self, path):
        raise NotImplementedError

    def save(self):
        """保存到 file_path；成功返回 True，失败返回 False。"""
        raise NotImplementedError

    def save_as(self, path):
        raise NotImplementedError

    def new_document(self):
        """创建空白文档（新建场景）。"""
        raise NotImplementedError

    # ---------- 通用方法 ----------
    def is_modified(self):
        return self._modified

    def set_modified(self, flag):
        if self._modified != flag:
            self._modified = flag
            self.modified_changed.emit(flag)

    def undo(self):
        pass

    def redo(self):
        pass

    def cut(self):
        pass

    def copy(self):
        pass

    def paste(self):
        pass

    # ---------- 查找替换（可选） ----------
    def find_next(self, text, case_sensitive=False):
        """查找下一个匹配；找到返回 True。"""
        return False

    def replace_current(self, find_text, repl_text, case_sensitive=False):
        """替换当前选中的匹配项。"""
        return False

    def replace_all(self, find_text, repl_text, case_sensitive=False):
        """全部替换；返回替换次数。"""
        return 0

    # ---------- 状态栏信息 ----------
    def status_text(self):
        """返回展示在状态栏的文档信息文本。"""
        return ""

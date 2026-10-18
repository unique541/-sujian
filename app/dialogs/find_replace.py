# -*- coding: utf-8 -*-
"""通用查找 / 替换对话框。"""
from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QCheckBox, QPushButton, QGroupBox,
)


class FindReplaceDialog(QDialog):
    """查找替换对话框，通过信号把请求转发给当前编辑器。"""

    find_next_requested = Signal(str, bool)              # (查找文本, 区分大小写)
    replace_requested = Signal(str, str, bool)           # (查找, 替换, 区分大小写)
    replace_all_requested = Signal(str, str, bool)       # (查找, 替换, 区分大小写)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("查找与替换")
        self.setModal(False)
        self.resize(420, 210)

        outer = QVBoxLayout(self)

        box = QGroupBox("查找", self)
        lay = QVBoxLayout(box)
        row = QHBoxLayout()
        row.addWidget(QLabel("查找内容：", box))
        self.edit_find = QLineEdit(box)
        row.addWidget(self.edit_find, 1)
        lay.addLayout(row)

        row2 = QHBoxLayout()
        row2.addWidget(QLabel("替换为：", box))
        self.edit_replace = QLineEdit(box)
        row2.addWidget(self.edit_replace, 1)
        lay.addLayout(row2)

        self.chk_case = QCheckBox("区分大小写", box)
        self.chk_loop = QCheckBox("循环查找", box)
        self.chk_loop.setChecked(True)
        row3 = QHBoxLayout()
        row3.addWidget(self.chk_case)
        row3.addWidget(self.chk_loop)
        row3.addStretch(1)
        lay.addLayout(row3)
        outer.addWidget(box)

        btns = QHBoxLayout()
        btn_find = QPushButton("查找下一个", self)
        btn_replace = QPushButton("替换", self)
        btn_replace_all = QPushButton("全部替换", self)
        btn_close = QPushButton("关闭", self)
        for b in (btn_find, btn_replace, btn_replace_all, btn_close):
            btns.addWidget(b)
        outer.addLayout(btns)

        btn_find.clicked.connect(self._emit_find_next)
        btn_replace.clicked.connect(self._emit_replace)
        btn_replace_all.clicked.connect(self._emit_replace_all)
        btn_close.clicked.connect(self.close)
        self.edit_find.returnPressed.connect(self._emit_find_next)

    def _case(self):
        return self.chk_case.isChecked()

    def _emit_find_next(self):
        text = self.edit_find.text()
        if text:
            self.find_next_requested.emit(text, self._case())

    def _emit_replace(self):
        find = self.edit_find.text()
        repl = self.edit_replace.text()
        if find:
            self.replace_requested.emit(find, repl, self._case())

    def _emit_replace_all(self):
        find = self.edit_find.text()
        repl = self.edit_replace.text()
        if find:
            self.replace_all_requested.emit(find, repl, self._case())

    def show_with_text(self, text=""):
        """弹出并预填选中文本。"""
        if text:
            self.edit_find.setText(text)
            self.edit_find.selectAll()
        self.show()
        self.raise_()
        self.activateWindow()

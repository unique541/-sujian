# -*- coding: utf-8 -*-
"""TXT 纯文本编辑器：多行文本编辑、行号栏、查找替换、编码处理、保存。

升级点：左侧行号栏实时刷新；状态栏显示当前光标行列号；编码探测与切换保留。
"""
import os

from PySide6.QtCore import Qt, Signal, QRect, QSize
from PySide6.QtGui import QTextCursor, QTextDocument, QFont, QColor, QPainter
from PySide6.QtPrintSupport import QPrinter, QPrintPreviewDialog
from PySide6.QtWidgets import QVBoxLayout, QPlainTextEdit, QWidget, QMenu

from .base_editor import BaseEditor

#: 支持的编码顺序（优先使用 UTF-8，其余按通用度排序）
ENCODINGS = ["utf-8", "gbk", "gb18030", "big5", "utf-16", "latin-1"]


class _LineNumberArea(QWidget):
    """QPlainTextEdit 左侧的行号栏。"""

    def __init__(self, editor):
        super().__init__(editor)
        self._editor = editor

    def sizeHint(self):
        return QSize(self._editor.line_number_area_width(), 0)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.fillRect(event.rect(), QColor("#f3f6f4"))
        block = self._editor.firstVisibleBlock()
        block_number = block.blockNumber()
        top = self._editor.blockBoundingGeometry(block).translated(
            self._editor.contentOffset()).top()
        bottom = top + self._editor.blockBoundingRect(block).height()
        while block.isValid() and top <= event.rect().bottom():
            if block.isVisible() and bottom >= event.rect().top():
                painter.setPen(QColor("#9aa6a0"))
                painter.drawText(0, int(top), self.width() - 6,
                                 self._editor.fontMetrics().height(),
                                 Qt.AlignmentFlag.AlignRight,
                                 str(block_number + 1))
            block = block.next()
            top = bottom
            bottom = top + self._editor.blockBoundingRect(block).height()


class TextEditor(BaseEditor):
    """基于 QPlainTextEdit 的纯文本编辑器（带行号栏）。"""

    #: 光标位置变化信号（用于状态栏）
    cursor_changed = Signal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        self.edit = QPlainTextEdit(self)
        self.edit.setFont(QFont("Consolas", 11))
        layout.addWidget(self.edit)

        # 自定义中文右键菜单（覆盖 QPlainTextEdit 默认英文菜单）
        self.edit.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.edit.customContextMenuRequested.connect(self._show_context_menu)

        self._encoding = "utf-8"
        self._original_bytes = b""

        # 行号栏
        self._line_area = _LineNumberArea(self.edit)
        self.edit.blockCountChanged.connect(self.update_line_number_area_width)
        self.edit.updateRequest.connect(self.update_line_number_area)
        self.edit.cursorPositionChanged.connect(self._on_cursor_moved)
        self.update_line_number_area_width(0)

        # 内容变化时标记修改
        self.edit.textChanged.connect(lambda: self.set_modified(True))

    # ---------- 行号栏 ----------
    def line_number_area_width(self):
        digits = len(str(max(1, self.edit.blockCount())))
        return 12 + self.edit.fontMetrics().horizontalAdvance("9") * digits

    def update_line_number_area_width(self, _=None):
        self.edit.setViewportMargins(self.line_number_area_width(), 0, 0, 0)

    def update_line_number_area(self, rect, dy):
        if dy:
            self._line_area.scroll(0, dy)
        else:
            self._line_area.update(0, rect.y(), self._line_area.width(), rect.height())
        if rect.contains(self.edit.viewport().rect()):
            self.update_line_number_area_width()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        cr = self.edit.contentsRect()
        self._line_area.setGeometry(
            QRect(cr.left(), cr.top(), self.line_number_area_width(), cr.height()))

    def _on_cursor_moved(self):
        cursor = self.edit.textCursor()
        line = cursor.blockNumber() + 1
        col = cursor.positionInBlock() + 1
        self.cursor_changed.emit(f"行：{line}  列：{col}")

    # ---------- 接口实现 ----------
    def doc_type(self):
        return "text"

    def new_document(self):
        self.edit.clear()
        self.file_path = None
        self._encoding = "utf-8"
        self.set_modified(False)

    def load(self, path):
        """按编码探测顺序读取文本文件。"""
        with open(path, "rb") as f:
            data = f.read()
        self._original_bytes = data
        decoded = False
        for enc in ENCODINGS:
            try:
                text = data.decode(enc)
                self._encoding = enc
                decoded = True
                break
            except (UnicodeDecodeError, LookupError):
                continue
        if not decoded:
            # 实在解不出来就用 utf-8 带替换字符读入，避免崩溃
            text = data.decode("utf-8", errors="replace")
            self._encoding = "utf-8"
        self.edit.setPlainText(text)
        self.file_path = path
        self.set_modified(False)

    def save(self):
        if not self.file_path:
            return False
        return self.save_as(self.file_path)

    def save_as(self, path):
        text = self.edit.toPlainText()
        with open(path, "w", encoding=self._encoding, newline="") as f:
            f.write(text)
        self.file_path = path
        self.set_modified(False)
        return True

    # ---------- 右键菜单（全中文） ----------
    def _show_context_menu(self, pos):
        """弹出全中文右键菜单，覆盖 QPlainTextEdit 默认英文菜单。"""
        menu = QMenu(self.edit)
        cursor = self.edit.textCursor()
        has_sel = cursor.hasSelection()

        act_undo = menu.addAction("撤销")
        act_undo.setEnabled(self.edit.document().isUndoAvailable())
        act_undo.triggered.connect(self.undo)
        act_redo = menu.addAction("重做")
        act_redo.setEnabled(self.edit.document().isRedoAvailable())
        act_redo.triggered.connect(self.redo)
        menu.addSeparator()

        act_cut = menu.addAction("剪切")
        act_cut.setEnabled(has_sel)
        act_cut.triggered.connect(self.cut)
        act_copy = menu.addAction("复制")
        act_copy.setEnabled(has_sel)
        act_copy.triggered.connect(self.copy)
        act_paste = menu.addAction("粘贴")
        act_paste.setEnabled(self.edit.canPaste())
        act_paste.triggered.connect(self.paste)
        act_del = menu.addAction("删除")
        act_del.setEnabled(has_sel)
        act_del.triggered.connect(lambda: self.edit.textCursor().removeSelectedText())
        act_all = menu.addAction("全选")
        act_all.triggered.connect(lambda: self.edit.selectAll())

        menu.exec(self.edit.mapToGlobal(pos))

    # ---------- 编辑操作 ----------
    def set_encoding(self, encoding):
        """切换保存编码。"""
        if encoding in ENCODINGS:
            self._encoding = encoding
            self.set_modified(True)

    def get_encoding(self):
        return self._encoding

    def undo(self):
        self.edit.undo()
        self.set_modified(True)

    def redo(self):
        self.edit.redo()
        self.set_modified(True)

    def cut(self):
        self.edit.cut()

    def copy(self):
        self.edit.copy()

    def paste(self):
        self.edit.paste()
        self.set_modified(True)

    # ---------- 查找替换 ----------
    def _find_flags(self, case_sensitive):
        if case_sensitive:
            return QTextDocument.FindFlag.FindCaseSensitively
        return QTextDocument.FindFlag(0)

    def find_next(self, text, case_sensitive=False):
        doc = self.edit.document()
        cursor = self.edit.textCursor()
        # 若当前已选中目标文本，则从其后继续找（实现“下一个”）
        start = cursor
        found = doc.find(text, start, self._find_flags(case_sensitive))
        if found.isNull() and self._find_flags(case_sensitive):
            return False
        if found.isNull():
            # 循环查找：从头开始
            found = doc.find(text, 0, self._find_flags(case_sensitive))
        if found.isNull():
            return False
        self.edit.setTextCursor(found)
        self.edit.ensureCursorVisible()
        return True

    def replace_current(self, find_text, repl_text, case_sensitive=False):
        cursor = self.edit.textCursor()
        if cursor.hasSelection() and cursor.selectedText() == find_text:
            cursor.insertText(repl_text)
            self.set_modified(True)
            return True
        return self.find_next(find_text, case_sensitive)

    def replace_all(self, find_text, repl_text, case_sensitive=False):
        doc = self.edit.document()
        flags = self._find_flags(case_sensitive)
        count = 0
        cursor = QTextCursor(doc)
        cursor.beginEditBlock()
        while True:
            found = doc.find(find_text, cursor, flags)
            if found.isNull():
                break
            found.insertText(repl_text)
            cursor = found
            count += 1
        cursor.endEditBlock()
        if count:
            self.set_modified(True)
        return count

    # ---------- 状态栏 ----------
    def print_preview(self):
        """打印预览当前文本。"""
        printer = QPrinter(QPrinter.PrinterMode.HighResolution)
        dialog = QPrintPreviewDialog(printer, self)
        dialog.paintRequested.connect(self._edit.print_)
        dialog.exec()

    def status_text(self):
        text = self.edit.toPlainText()
        lines = len(text.splitlines())
        chars = len(text)
        cursor = self.edit.textCursor()
        line = cursor.blockNumber() + 1
        col = cursor.positionInBlock() + 1
        return f"行：{line}  列：{col}   总行数：{lines}  字符数：{chars}  编码：{self._encoding}"

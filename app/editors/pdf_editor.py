# -*- coding: utf-8 -*-
"""PDF 阅读器（只读查看模式）。

本文件只保留「打开查看」能力，已移除全部创建与编辑功能：
- 不再使用 reportlab，不能新建空白 PDF、不能保存导出 PDF、不能编辑页面；
- 只能通过 pypdfium2 渲染页面为图片进行翻页 / 缩放 / 旋转 / 适合宽度浏览；
- 另存为仅复制原始 PDF 文件（不改变内容）。
"""
import os
import shutil

import pypdfium2 as pdfium
from PIL.ImageQt import ImageQt
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QLabel, QScrollArea, QMessageBox,
)
from PySide6.QtGui import QPixmap, QTransform

from app.editors.base_editor import BaseEditor


class PdfEditor(BaseEditor):
    """PDF 查看器：只读打开 / 翻页 / 缩放 / 旋转 / 适合宽度。"""

    #: 缩放范围
    MIN_SCALE = 0.5
    MAX_SCALE = 4.0

    def __init__(self, parent=None):
        super().__init__(parent)
        self._pdf = None            # pypdfium2 PdfDocument
        self._page_count = 0
        self._current = 0           # 当前页索引（0 起）
        self._scale = 1.0
        self._rotation = 0          # 顺时针旋转角度（0/90/180/270）
        self._source_path = None    # 原始 PDF 路径（另存为时复制它）

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        self._scroll = QScrollArea(self)
        self._scroll.setWidgetResizable(True)
        self._scroll.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._image_label = QLabel("请打开 PDF 文件")
        self._image_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._image_label.setStyleSheet("color: #888888; font-size: 14px;")
        self._scroll.setWidget(self._image_label)
        layout.addWidget(self._scroll)

    # ---------- 文档类型 ----------
    def doc_type(self):
        return "pdf"

    def is_editable_new(self):
        """已移除 PDF 创建/编辑能力，恒为只读。"""
        return False

    # ---------- 打开 / 保存 ----------
    def load(self, path):
        """打开 PDF 文件并渲染第一页。"""
        if not os.path.exists(path):
            raise FileNotFoundError(f"文件不存在：{path}")
        pdf = pdfium.PdfDocument(path)
        if pdf is None or len(pdf) == 0:
            raise ValueError("无法打开 PDF 或文件为空")
        # 先关闭旧文档再替换，避免句柄泄漏
        if self._pdf is not None:
            try:
                self._pdf.close()
            except Exception:
                pass
        self._pdf = pdf
        self._page_count = len(pdf)
        self._current = 0
        self._scale = 1.0
        self._rotation = 0
        self._source_path = path
        self.file_path = path
        self.set_modified(False)
        self._render_current()
        return True

    def save(self):
        """PDF 只读，不允许覆盖保存；返回 False 由上层转另存为。"""
        return False

    def save_as(self, path):
        """另存为：复制原始 PDF 文件到新路径（不改变内容）。"""
        if not self._source_path or not os.path.exists(self._source_path):
            return False
        try:
            shutil.copy2(self._source_path, path)
        except Exception as e:
            QMessageBox.critical(self, "另存为失败", str(e))
            return False
        self.file_path = path
        self.set_modified(False)
        return True

    def new_document(self):
        """PDF 创建功能已移除，仅提示查看。"""
        QMessageBox.information(self, "提示", "PDF 仅支持打开查看，不支持新建编辑。")
        return False

    # ---------- 查看操作（Ribbon 按钮调用） ----------
    def _render_current(self):
        """按当前页 / 缩放 / 旋转渲染页面图片。"""
        if self._pdf is None:
            self._image_label.setPixmap(QPixmap())
            return
        page = self._pdf[self._current]
        try:
            dpr = self.devicePixelRatioF() or 1.0
            bitmap = page.render(scale=self._scale * dpr)
            pil = bitmap.to_pil()
            qimg = ImageQt(pil)
            pix = QPixmap.fromImage(qimg)
            if self._rotation:
                pix = pix.transformed(
                    QTransform().rotate(self._rotation),
                    Qt.TransformationMode.SmoothTransformation)
            self._image_label.setPixmap(pix)
        finally:
            page.close()

    def prev_page(self):
        if self._pdf is None or self._current <= 0:
            return
        self._current -= 1
        self._render_current()

    def next_page(self):
        if self._pdf is None or self._current >= self._page_count - 1:
            return
        self._current += 1
        self._render_current()

    def zoom_in(self):
        self.set_scale(self._scale * 1.2)

    def zoom_out(self):
        self.set_scale(self._scale / 1.2)

    def set_scale(self, scale):
        """设置缩放比例并重新渲染（带上下限钳制）。"""
        if self._pdf is None:
            return
        self._scale = max(self.MIN_SCALE, min(self.MAX_SCALE, scale))
        self._render_current()

    def fit_width(self):
        """适合宽度：按当前视口宽度自动缩放。"""
        if self._pdf is None:
            return
        page = self._pdf[self._current]
        try:
            width, _ = page.get_size()
        finally:
            page.close()
        if not width:
            return
        view_w = max(50, self._scroll.viewport().width())
        dpr = self.devicePixelRatioF() or 1.0
        self.set_scale(view_w / (width * dpr))

    def rotate_clockwise(self):
        """顺时针旋转 90 度。"""
        self._rotation = (self._rotation + 90) % 360
        self._render_current()

    # ---------- 状态栏 ----------
    def status_text(self):
        if self._pdf is None:
            return "PDF 阅读器（只读）"
        return (f"第 {self._current + 1} / {self._page_count} 页 · "
                f"缩放 {int(self._scale * 100)}% · 旋转 {self._rotation}°")

    # ---------- 生命周期 ----------
    def closeEvent(self, event):
        if self._pdf is not None:
            try:
                self._pdf.close()
            except Exception:
                pass
            self._pdf = None
        super().closeEvent(event)

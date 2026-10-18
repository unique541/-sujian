# -*- coding: utf-8 -*-
"""仿 WPS 的启动欢迎界面。

- 顶部大标题 + 绿色主题横幅；
- 「新建文档」区域：Excel / Word / TXT / PDF 四个入口卡片；
- 「最近打开」区域：最近文件列表（点击重新打开）；
- 加载时整页淡入，卡片悬停轻微上浮 + 边框变色（动画过渡）。
"""
import os

from PySide6.QtCore import Qt, QVariantAnimation, QEasingCurve
from PySide6.QtGui import QColor, QFont
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QFrame,
    QGridLayout, QPushButton, QMenu,
)
from app.animation import fade_in


class _HoverCard(QFrame):
    """可点击的入口卡片：悬停上浮 + 变色动画。"""

    def __init__(self, title, subtitle, parent=None):
        super().__init__(parent)
        self.setObjectName("welcomeCard")
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setFixedSize(170, 100)
        self._base_color = QColor("#ffffff")
        self._hover_color = QColor("#e8f6ee")
        self._anim = None

        lay = QVBoxLayout(self)
        lay.setContentsMargins(14, 14, 14, 10)
        lay.setSpacing(4)
        self.title_label = QLabel(title)
        self.title_label.setObjectName("cardTitle")
        self.subtitle_label = QLabel(subtitle)
        self.subtitle_label.setObjectName("cardSubtitle")
        self.subtitle_label.setWordWrap(True)
        lay.addWidget(self.title_label)
        lay.addWidget(self.subtitle_label)
        lay.addStretch(1)

    def _apply_color(self, color):
        self.setStyleSheet(
            "QFrame#welcomeCard { background-color: %s; border: 1px solid #d5e8dc; border-radius: 8px; }"
            "QLabel#cardTitle { font-size: 16px; font-weight: bold; color: #1d7a44; }"
            "QLabel#cardSubtitle { font-size: 12px; color: #777777; }" % color.name()
        )

    def _forget_anim(self, anim):
        """动画结束仅解除引用（对象复用，随父 widget 生命周期管理）。"""
        if self._anim is anim:
            self._anim = None

    def _animate_color(self, target):
        """复用同一个 QVariantAnimation 对象，先停旧再启新，避免叠加闪烁。"""
        if self._anim is None:
            self._anim = QVariantAnimation(self)
            self._anim.valueChanged.connect(self._apply_color)
            # 不依赖 DeleteWhenStopped（运行中随父析构可能二次删除）；
            # 用 finished 回调主动置空，避免 C++ 已删而 Python 仍持有引用
            self._anim.finished.connect(
                lambda a=self._anim: self._forget_anim(a))
        else:
            try:
                self._anim.stop()
            except RuntimeError:
                pass   # 旧动画已被 Qt 删除，忽略即可
        anim = self._anim
        anim.setDuration(140)
        anim.setStartValue(QColor(self._base_color))
        anim.setEndValue(target)
        anim.setEasingCurve(QEasingCurve.Type.OutCubic)
        anim.start()

    def enterEvent(self, event):
        self._animate_color(self._hover_color)
        super().enterEvent(event)

    def leaveEvent(self, event):
        self._animate_color(self._base_color)
        super().leaveEvent(event)


class WelcomePage(QWidget):
    """启动欢迎页：新建入口 + 最近文件。"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("welcomePage")
        self.new_clicked = None   # callable: doc_type
        self.open_clicked = None  # callable: file_path

        outer = QVBoxLayout(self)
        outer.setContentsMargins(60, 40, 60, 40)
        outer.setSpacing(28)

        # 标题区
        title = QLabel("素笺")
        title.setObjectName("welcomeTitle")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        sub = QLabel("今天做点什么好呢")
        sub.setObjectName("welcomeSubtitle")
        sub.setAlignment(Qt.AlignmentFlag.AlignCenter)
        outer.addWidget(title)
        outer.addWidget(sub)

        # 新建区
        new_title = QLabel("新建")
        new_title.setObjectName("welcomeSection")
        outer.addWidget(new_title)

        cards = QGridLayout()
        cards.setSpacing(14)
        self._cards = {}
        specs = [
            ("Excel 表格", "新建空白表格", "excel"),
            ("Word 文档", "新建空白文档", "word"),
            ("TXT 文本", "新建纯文本文档", "text"),
            ("PDF 阅读", "打开 PDF 文档", "pdf"),
        ]
        for idx, (title_, subtitle_, key) in enumerate(specs):
            card = _HoverCard(title_, subtitle_)
            card.mousePressEvent = lambda e, k=key: self._on_new(k)
            cards.addWidget(card, idx // 4, idx % 4)
        cards.setColumnStretch(4, 1)
        outer.addLayout(cards)

        # 最近文件区
        recent_title = QLabel("最近打开")
        recent_title.setObjectName("welcomeSection")
        outer.addWidget(recent_title)

        self.recent_list = QVBoxLayout()
        self.recent_list.setSpacing(4)
        self.recent_list.addStretch(1)
        outer.addLayout(self.recent_list, 1)

        outer.addStretch(1)
        fade_in(self, duration=200, start=0.0, end=1.0)

    # ---------- 交互 ----------
    def _on_new(self, doc_type):
        if self.new_clicked:
            self.new_clicked(doc_type)

    def set_recent_files(self, paths):
        """刷新最近文件列表。"""
        while self.recent_list.count() > 1:
            item = self.recent_list.takeAt(0)
            w = item.widget()
            if w is not None:
                w.deleteLater()

        if not paths:
            label = QLabel("暂无最近打开的文件")
            label.setObjectName("recentEmpty")
            self.recent_list.insertWidget(0, label)
            return

        for path in paths:
            btn = QPushButton(os.path.basename(path))
            btn.setObjectName("recentButton")
            btn.setToolTip(path)
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
            btn.clicked.connect(lambda checked=False, p=path: self._on_open(p))
            # 中文右键菜单：打开 / 复制路径
            btn.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
            btn.customContextMenuRequested.connect(
                lambda pos, p=path, b=btn: self._show_recent_menu(b, pos, p))
            self.recent_list.insertWidget(self.recent_list.count() - 1, btn)

    def _show_recent_menu(self, button, pos, path):
        """最近文件右键菜单（全中文）。"""
        menu = QMenu(button)
        act_open = menu.addAction("打开文件")
        act_open.triggered.connect(lambda: self._on_open(path))
        act_copy = menu.addAction("复制文件路径")
        act_copy.triggered.connect(lambda: self._copy_path(path))
        menu.exec(button.mapToGlobal(pos))

    def _copy_path(self, path):
        from PySide6.QtWidgets import QApplication
        QApplication.clipboard().setText(path)

    def _on_open(self, path):
        if self.open_clicked:
            self.open_clicked(path)

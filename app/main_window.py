# -*- coding: utf-8 -*-
"""主窗口：快速访问工具栏 + Ribbon + 欢迎页 + 多文档标签页 + 状态栏。

功能：
- 仿 WPS 启动欢迎页：最近文件列表 + 新建各类型文档入口 + 打开文件
- 快速访问工具栏：新建 / 打开 / 保存 / 另存为 / 撤销 / 重做 / 深色主题切换
- 精细 Ribbon（开始 / 插入 / 页面 / 引用 / 审阅 / 视图），按钮按文档类型启用/禁用
- 打开按钮弹出 QFileDialog 支持多选，按扩展名自动路由到对应编辑器
- 文档标签切换淡入动画、标签关闭淡出动画
- 深色 / 浅色主题可切换
"""
import os
import sys

from PySide6.QtCore import Qt, QEvent, QSettings
from PySide6.QtGui import QAction, QKeySequence, QColor, QIcon
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QTabWidget, QFileDialog,
    QMessageBox, QLabel, QMenu, QInputDialog, QListWidget, QListWidgetItem,
    QToolButton, QPushButton, QFrame, QToolBar,
)

from app.animation import fade_in, fade_out_then, pop_in, clear_transition
from app.ribbon import WpsRibbon
from app.dialogs.find_replace import FindReplaceDialog
from app.editors import TextEditor, WordEditor, ExcelEditor, PdfEditor

#: 支持的扩展名 -> 文档类型
SUPPORTED_EXT = {
    ".xlsx": "excel", ".xls": "excel", ".xlsm": "excel",
    ".docx": "word",
    ".txt": "text", ".md": "text", ".log": "text", ".py": "text",
    ".json": "text", ".ini": "text", ".csv": "text",
    ".pdf": "pdf",
}

_FILE_FILTER = ("所有支持的文件 (*.xlsx *.xls *.xlsm *.docx *.txt *.md *.log "
                "*.py *.json *.ini *.csv *.pdf);;"
                "Excel 表格 (*.xlsx *.xls *.xlsm);;"
                "Word 文档 (*.docx);;"
                "文本文件 (*.txt *.md *.log *.py *.json *.ini *.csv);;"
                "PDF 文件 (*.pdf);;"
                "所有文件 (*.*)")

#: 新建类型的中文名 -> 内部类型（PDF 仅支持查看，不在新建列表）
NEW_TYPES = {"excel": "Excel 表格", "word": "Word 文档", "text": "文本文件"}

#: 文档类型 -> 状态栏中文名
_DOC_LABEL = {"excel": "Excel 表格", "word": "Word 文档",
              "text": "文本文件", "pdf": "PDF 文档"}


def _style_sheet():
    """生成 WPS 风格浅色绿主题 QSS。"""
    bg, panel, line, text, sub = "#f5f6f7", "#ffffff", "#d9d9d9", "#333333", "#666666"
    sel = "#e1f3e8"
    return f"""
QMainWindow, QWidget#welcomePage {{
    background: {bg};
}}
#titleBar {{
    background: {panel};
    border-bottom: 1px solid {line};
}}
#titleBarTitle {{
    color: {text};
    font-size: 13px;
    font-weight: bold;
}}
QToolButton#titleBarBtn, QToolButton#titleBarBtnClose {{
    border: none;
    background: transparent;
    color: {sub};
    font-size: 14px;
    border-radius: 4px;
    font-family: "Segoe UI Symbol", "Microsoft YaHei", "微软雅黑";
}}
QToolButton#titleBarBtn:hover {{
    background: #d8efe2;
    color: #1d7a44;
}}
QToolButton#titleBarBtn:pressed {{ background: #b9e2cc; }}
QToolButton#titleBarBtnClose:hover {{ background: #e81123; color: white; }}
QToolButton#titleBarBtnClose:pressed {{ background: #c50f1f; }}
QMenuBar {{
    background: {bg};
    padding: 2px 6px;
    font-size: 13px;
}}
QMenuBar::item {{ color: {text}; }}
QMenuBar::item:selected {{ background: {sel}; }}
QMenu {{
    background: {panel};
    border: 1px solid {line};
    font-size: 13px;
    color: {text};
}}
QMenu::item:selected {{ background: {sel}; }}
QToolBar {{
    background: {bg};
    border: none;
    spacing: 2px;
    padding: 2px 4px;
}}
#wpsRibbon {{ background: {panel}; border-bottom: 1px solid {line}; }}
#groupTitle {{ color: {sub}; font-size: 11px; }}
QTabBar::tab {{
    background: transparent;
    color: {text};
    padding: 6px 22px;
    font-size: 13px;
    border: none;
    border-bottom: 3px solid transparent;
}}
QTabBar::tab:hover {{ background: {sel}; }}
QTabBar::tab:selected {{
    color: #1d7a44;
    font-weight: bold;
    border-bottom: 3px solid #2e9e5b;
}}
QTabWidget::pane {{ border-top: 1px solid {line}; }}
QStatusBar {{ background: {panel}; border-top: 1px solid {line}; color: {text}; }}
QListWidget {{
    background: {panel};
    color: {text};
    border: 1px solid {line};
    border-radius: 6px;
}}
QListWidget::item {{ padding: 6px 8px; }}
QListWidget::item:selected {{ background: {sel}; }}
QScrollBar:vertical, QScrollBar:horizontal {{
    background: {bg};
    width: 10px; height: 10px;
}}
QScrollBar::handle:vertical, QScrollBar::handle:horizontal {{
    background: #c4c9ce;
    border-radius: 5px;
    min-height: 24px; min-width: 24px;
}}
QScrollBar::add-line, QScrollBar::sub-line {{ height: 0; width: 0; }}
QScrollBar::add-page, QScrollBar::sub-page {{ background: transparent; }}
QToolTip {{
    background: {panel};
    color: {text};
    border: 1px solid {line};
    padding: 3px 6px;
}}
QTableWidget {{
    background: {panel};
    color: {text};
    gridline-color: #e0e0e0;
}}
"""


class TitleBar(QWidget):
    """自绘标题栏（去系统边框后的窗口拖拽区 + 窗口按钮）。

    仿 WPS 无边框方案：浅色背景 + 素笺绿主题，左侧程序图标与标题，
    右侧最小化 / 最大化(还原) / 关闭三个按钮（悬停高亮）。
    支持按住标题栏拖动窗口（startSystemMove）与双击最大化/还原。
    """

    _BTN_MIN = "—"
    _BTN_MAX = "□"
    _BTN_RESTORE = "❐"
    _BTN_CLOSE = "✕"

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("titleBar")
        self.setFixedHeight(36)
        self._drag_armed = False
        self._press_pos = None

        lay = QHBoxLayout(self)
        lay.setContentsMargins(12, 0, 0, 0)
        lay.setSpacing(6)

        # 左侧：程序图标 + 标题
        self.lbl_icon = QLabel(self)
        _icon = QIcon(os.path.join(
            getattr(sys, "_MEIPASS",
                    os.path.dirname(os.path.dirname(os.path.abspath(__file__)))),
            "build_assets", "icon.ico"))
        if not _icon.isNull():
            self.lbl_icon.setPixmap(_icon.pixmap(18, 18))
        lay.addWidget(self.lbl_icon)

        self.lbl_title = QLabel("素笺", self)
        self.lbl_title.setObjectName("titleBarTitle")
        lay.addWidget(self.lbl_title)
        lay.addStretch(1)

        # 右侧：最小化 / 最大化(还原) / 关闭
        self.btn_min = QToolButton(self)
        self.btn_min.setObjectName("titleBarBtn")
        self.btn_min.setText(self._BTN_MIN)
        self.btn_min.setToolTip("最小化")
        self.btn_min.setFixedSize(44, 30)
        self.btn_min.clicked.connect(self._on_min)
        lay.addWidget(self.btn_min)

        self.btn_max = QToolButton(self)
        self.btn_max.setObjectName("titleBarBtn")
        self.btn_max.setText(self._BTN_MAX)
        self.btn_max.setToolTip("最大化")
        self.btn_max.setFixedSize(44, 30)
        self.btn_max.clicked.connect(self._on_max)
        lay.addWidget(self.btn_max)

        self.btn_close = QToolButton(self)
        self.btn_close.setObjectName("titleBarBtnClose")
        self.btn_close.setText(self._BTN_CLOSE)
        self.btn_close.setToolTip("关闭")
        self.btn_close.setFixedSize(44, 30)
        self.btn_close.clicked.connect(self._on_close)
        lay.addWidget(self.btn_close)

    # ---------- 窗口按钮 ----------
    def _on_min(self):
        self.window().showMinimized()

    def _on_max(self):
        w = self.window()
        if w.isMaximized():
            w.showNormal()
        else:
            w.showMaximized()

    def _on_close(self):
        self.window().close()

    def _sync_max_btn(self):
        """随窗口最大化状态刷新按钮图标（还原/最大化）。"""
        if self.window().isMaximized():
            self.btn_max.setText(self._BTN_RESTORE)
            self.btn_max.setToolTip("还原")
        else:
            self.btn_max.setText(self._BTN_MAX)
            self.btn_max.setToolTip("最大化")

    # ---------- 窗口拖动 ----------
    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self._drag_armed = True
            self._press_pos = event.globalPosition().toPoint()
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if (self._drag_armed and self._press_pos is not None
                and event.buttons() & Qt.MouseButton.LeftButton):
            # 超过阈值后再启动系统级移动，避免单击误触发
            delta = (event.globalPosition().toPoint()
                     - self._press_pos).manhattanLength()
            if delta >= 3:
                win = self.window().windowHandle()
                if win is not None:
                    win.startSystemMove()
                self._drag_armed = False
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        self._drag_armed = False
        self._press_pos = None
        super().mouseReleaseEvent(event)

    def mouseDoubleClickEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self._on_max()
        super().mouseDoubleClickEvent(event)


class WelcomePage(QWidget):
    """仿 WPS 启动欢迎页：新建入口 + 最近文件。"""

    #: 用户动作信号
    new_requested = __import__("PySide6.QtCore", fromlist=["Signal"]).Signal(str)
    open_requested = __import__("PySide6.QtCore", fromlist=["Signal"]).Signal()
    recent_clicked = __import__("PySide6.QtCore", fromlist=["Signal"]).Signal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("welcomePage")

        outer = QVBoxLayout(self)
        outer.setContentsMargins(60, 40, 60, 40)
        outer.setSpacing(10)
        outer.addStretch(1)

        title = QLabel("素笺")
        title.setStyleSheet("font-size: 30px; font-weight: bold; color: #2e9e5b;")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        outer.addWidget(title)

        sub = QLabel("今天做点什么好呢")
        sub.setStyleSheet("font-size: 13px; color: #888888;")
        sub.setAlignment(Qt.AlignmentFlag.AlignCenter)
        outer.addWidget(sub)
        outer.addSpacing(14)

        # 新建入口（4 个卡片按钮）
        row = QHBoxLayout()
        row.setSpacing(14)
        row.addStretch(1)
        self._new_buttons = {}
        for key, label in NEW_TYPES.items():
            btn = QPushButton(label)
            btn.setFixedSize(132, 92)
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
            btn.setStyleSheet(
                "QPushButton { background: #ffffff; border: 1px solid #d9d9d9;"
                " border-radius: 8px; font-size: 15px; color: #333333; }"
                "QPushButton:hover { border-color: #2e9e5b; color: #2e9e5b;"
                " background: #f0faf4; }")
            btn.clicked.connect(lambda checked=False, k=key: self.new_requested.emit(k))
            row.addWidget(btn)
            self._new_buttons[key] = btn
        row.addStretch(1)
        outer.addLayout(row)
        outer.addSpacing(10)

        # 打开按钮
        open_btn = QPushButton("打开文件...")
        open_btn.setFixedSize(180, 40)
        open_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        open_btn.setStyleSheet(
            "QPushButton { background: #2e9e5b; color: #ffffff; border: none;"
            " border-radius: 6px; font-size: 14px; }"
            "QPushButton:hover { background: #269150; }"
            "QPushButton:pressed { background: #1e7a42; }")
        open_btn.clicked.connect(lambda: self.open_requested.emit())
        open_btn.setParent(self)
        outer.addWidget(open_btn, 0, Qt.AlignmentFlag.AlignHCenter)

        outer.addSpacing(24)

        # 最近文件
        recent_label = QLabel("最近使用")
        recent_label.setStyleSheet("font-size: 12px; color: #999999;")
        outer.addWidget(recent_label)

        self.recent_list = QListWidget()
        self.recent_list.setMinimumHeight(150)
        self.recent_list.itemClicked.connect(self._on_recent_clicked)
        outer.addWidget(self.recent_list)
        outer.addStretch(2)

    def _on_recent_clicked(self, item):
        path = item.data(Qt.ItemDataRole.UserRole)
        if path:
            self.recent_clicked.emit(path)

    def set_recent_files(self, paths):
        """刷新最近文件列表；不存在或无记录时显示提示项。"""
        self.recent_list.clear()
        if not paths:
            item = QListWidgetItem("（暂无最近文件）")
            item.setFlags(Qt.ItemFlag.NoItemFlags)
            item.setForeground(QColor("#aaaaaa"))
            self.recent_list.addItem(item)
            return
        for path in paths:
            item = QListWidgetItem(os.path.basename(path))
            item.setToolTip(path)
            item.setData(Qt.ItemDataRole.UserRole, path)
            self.recent_list.addItem(item)


class MainWindow(QMainWindow):
    """仿 WPS 主窗口。"""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("素笺")
        # 去掉系统原生边框：窗口内容自绘标题栏 + 窗口按钮（仿 WPS 无边框）
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint)
        # 窗口 / 任务栏图标（与 QApplication 一致；显式设置保证打包后依旧生效）
        _icon = QIcon(os.path.join(
            getattr(sys, "_MEIPASS",
                    os.path.dirname(os.path.dirname(os.path.abspath(__file__)))),
            "build_assets", "icon.ico"))
        if not _icon.isNull():
            self.setWindowIcon(_icon)
        self.resize(1360, 860)

        self._recent_files = []
        self._settings = QSettings("SuJian", "SuJian")
        self._load_recent_files()

        self._build_ui()
        self._connect_ribbon()
        self.setStyleSheet(_style_sheet())
        # 自绘标题栏替换系统边框（原菜单栏被隐藏但保留，快捷键/动作仍生效）
        self._title_bar = TitleBar(self)
        self.setMenuWidget(self._title_bar)
        self._on_tab_changed(0)

    def changeEvent(self, event):
        """窗口最大化/还原状态变化时同步标题栏按钮图标。"""
        if event.type() == QEvent.Type.WindowStateChange:
            tb = getattr(self, "_title_bar", None)
            if tb is not None:
                tb._sync_max_btn()
        super().changeEvent(event)

    # ---------- UI 构建 ----------
    def _build_ui(self):
        central = QWidget(self)
        lay = QVBoxLayout(central)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(0)

        # 快速访问工具栏
        self.quick = QToolBar("快速访问")
        self.quick.setMovable(False)
        self.addToolBar(Qt.ToolBarArea.TopToolBarArea, self.quick)
        self._build_quick_bar()

        # Ribbon
        self.ribbon = WpsRibbon(central)
        lay.addWidget(self.ribbon)

        # 文档标签页（index 0 固定为欢迎页）
        self.tabs = QTabWidget(central)
        self.tabs.setTabsClosable(True)
        self.tabs.setMovable(True)
        self.tabs.setDocumentMode(True)
        lay.addWidget(self.tabs, 1)
        self.setCentralWidget(central)

        self.welcome = WelcomePage()
        self.tabs.addTab(self.welcome, "欢迎")
        self.tabs.tabCloseRequested.connect(self._close_tab)
        self.tabs.currentChanged.connect(self._on_tab_changed)
        self.welcome.new_requested.connect(self.new_document)
        self.welcome.open_requested.connect(self.open_file)
        self.welcome.recent_clicked.connect(self.open_file)

        # 菜单栏
        self._build_menu()

        # 状态栏
        self._status_doc = QLabel("未打开文档")
        self._status_info = QLabel("")
        self.statusBar().addWidget(self._status_doc, 1)
        self.statusBar().addPermanentWidget(self._status_info)
        # 格式刷模式标志（True 表示已记录格式、等待应用）
        self._format_brush_active = False

        # 查找替换对话框
        self._find_dialog = FindReplaceDialog(self)
        self._find_dialog.find_next_requested.connect(self._on_find_next)
        self._find_dialog.replace_requested.connect(self._on_replace)
        self._find_dialog.replace_all_requested.connect(self._on_replace_all)

    def _build_quick_bar(self):
        """快速访问工具栏：新建 / 打开 / 保存 / 另存为 | 撤销 / 重做。"""
        self.q_new = QToolButton()
        self.q_new.setText("新建")
        self.q_new.setToolTip("新建文档")
        self.q_new.clicked.connect(self._ribbon_new)
        self.quick.addWidget(self.q_new)

        self.q_open = QToolButton()
        self.q_open.setText("打开")
        self.q_open.setToolTip("打开文件 (Ctrl+O)，支持多选")
        self.q_open.clicked.connect(self.open_file)
        self.quick.addWidget(self.q_open)

        self.q_save = QToolButton()
        self.q_save.setText("保存")
        self.q_save.setToolTip("保存 (Ctrl+S)")
        self.q_save.clicked.connect(self.save_file)
        self.quick.addWidget(self.q_save)

        self.q_save_as = QToolButton()
        self.q_save_as.setText("另存为")
        self.q_save_as.setToolTip("另存为 (Ctrl+Shift+S)")
        self.q_save_as.clicked.connect(self.save_file_as)
        self.quick.addWidget(self.q_save_as)

        self.quick.addSeparator()

        self.q_undo = QToolButton()
        self.q_undo.setText("撤销")
        self.q_undo.setToolTip("撤销 (Ctrl+Z)")
        self.q_undo.clicked.connect(self._do_undo)
        self.quick.addWidget(self.q_undo)

        self.q_redo = QToolButton()
        self.q_redo.setText("重做")
        self.q_redo.setToolTip("重做 (Ctrl+Y)")
        self.q_redo.clicked.connect(self._do_redo)
        self.quick.addWidget(self.q_redo)

        # 快速访问栏按钮统一样式（文字完整显示）
        for btn in (self.q_new, self.q_open, self.q_save, self.q_save_as,
                    self.q_undo, self.q_redo):
            btn.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextOnly)
            btn.setMinimumWidth(52)
            btn.setFixedHeight(30)

    def _build_menu(self):
        menubar = self.menuBar()

        # 文件
        menu_file = menubar.addMenu("文件(&F)")
        sub_new = menu_file.addMenu("新建")
        for key, label in NEW_TYPES.items():
            act = QAction(label, self)
            act.triggered.connect(lambda checked=False, k=key: self.new_document(k))
            sub_new.addAction(act)

        act_open = QAction("打开...", self)
        act_open.setShortcut(QKeySequence.StandardKey.Open)
        act_open.triggered.connect(self.open_file)
        menu_file.addAction(act_open)

        act_save = QAction("保存", self)
        act_save.setShortcut(QKeySequence.StandardKey.Save)
        act_save.triggered.connect(self.save_file)
        menu_file.addAction(act_save)

        act_save_as = QAction("另存为...", self)
        act_save_as.setShortcut(QKeySequence("Ctrl+Shift+S"))
        act_save_as.triggered.connect(self.save_file_as)
        menu_file.addAction(act_save_as)

        self.recent_menu = menu_file.addMenu("最近文件")
        menu_file.addSeparator()
        act_exit = QAction("退出", self)
        act_exit.triggered.connect(self.close)
        menu_file.addAction(act_exit)

        # 编辑
        menu_edit = menubar.addMenu("编辑(&E)")
        act_undo = QAction("撤销", self)
        act_undo.setShortcut(QKeySequence.StandardKey.Undo)
        act_undo.triggered.connect(self._do_undo)
        menu_edit.addAction(act_undo)
        act_redo = QAction("重做", self)
        act_redo.setShortcut(QKeySequence.StandardKey.Redo)
        act_redo.triggered.connect(self._do_redo)
        menu_edit.addAction(act_redo)
        menu_edit.addSeparator()
        for label, method in (("剪切", self._do_cut), ("复制", self._do_copy),
                              ("粘贴", self._do_paste)):
            act = QAction(label, self)
            act.triggered.connect(method)
            menu_edit.addAction(act)
        menu_edit.addSeparator()
        act_find = QAction("查找替换", self)
        act_find.setShortcut(QKeySequence.StandardKey.Find)
        act_find.triggered.connect(self.open_find_dialog)
        menu_edit.addAction(act_find)

        # 视图
        menu_view = menubar.addMenu("视图(&V)")
        act_zoom = QAction("放大", self)
        act_zoom.triggered.connect(lambda: self._call_editor("set_font_size", 1))
        menu_view.addAction(act_zoom)
        act_zoom_out = QAction("缩小", self)
        act_zoom_out.triggered.connect(lambda: self._call_editor("set_font_size", -1))
        menu_view.addAction(act_zoom_out)

        # 帮助
        menu_help = menubar.addMenu("帮助(&H)")
        act_about = QAction("关于", self)
        act_about.triggered.connect(self._show_about)
        menu_help.addAction(act_about)

    # ---------- Ribbon 连接 ----------
    def _report_error(self, exc):
        """统一错误记录：追加写入 error.log 并弹中文提示，防止静默闪退。"""
        import traceback
        try:
            log_path = os.path.join(os.path.dirname(os.path.dirname(
                os.path.abspath(__file__))), "error.log")
            with open(log_path, "a", encoding="utf-8") as f:
                f.write("=" * 60 + "\n")
                f.write(traceback.format_exc())
                f.write("\n")
        except Exception:
            pass
        try:
            QMessageBox.critical(
                self, "操作出错",
                f"操作执行失败，详情已记录到 error.log。\n\n{exc}")
        except Exception:
            pass

    def _safe(self, fn):
        """把回调包装为安全调用：任何异常都走 _report_error，绝不导致闪退。"""
        def wrapper(*args, **kwargs):
            try:
                return fn(*args, **kwargs)
            except Exception as e:
                self._report_error(e)
        return wrapper

    def _connect_ribbon(self):
        """连接 Ribbon 全部按钮。c=clicked，tc=toggled，均加安全包装防闪退。"""
        r = self.ribbon
        b = r.button

        def c(name, slot):
            btn = b(name)
            if btn is not None:
                btn.clicked.connect(self._safe(slot))

        def tc(name, slot):
            btn = b(name)
            if btn is not None:
                btn.toggled.connect(self._safe(slot))

        c("new", self._ribbon_new)
        c("open", self.open_file)
        c("save", self.save_file)
        c("save_as", self.save_file_as)
        c("print_preview", lambda: self._call_editor("print_preview"))

        c("cut", self._do_cut)
        c("copy", self._do_copy)
        c("paste", self._do_paste)

        tc("bold", lambda on: self._call_editor("set_bold", on))
        tc("italic", lambda on: self._call_editor("set_italic", on))
        tc("underline", lambda on: self._call_editor("set_underline", on))
        tc("strikethrough", lambda on: self._call_editor("set_strikethrough", on))
        c("font_color", lambda: self._call_editor("set_font_color"))
        c("fill_color", lambda: self._call_editor("set_fill_color"))
        c("border", lambda: self._call_editor("set_border"))
        c("font_up", lambda: self._call_editor("set_font_size", 1))
        c("font_down", lambda: self._call_editor("set_font_size", -1))

        # 字体 / 字号 / 数字格式下拉框（Word / Excel 共用编辑器方法）
        combo = b("font_family")
        if combo is not None:
            combo.currentTextChanged.connect(
                self._safe(lambda text: self._call_editor("set_font_family", text)))
        combo = b("font_size")
        if combo is not None:
            combo.currentTextChanged.connect(self._safe(self._on_font_size_combo))
        combo = b("num_format")
        if combo is not None:
            combo.currentTextChanged.connect(
                self._safe(lambda text: self._call_editor("set_num_format_explicit", text)))

        c("align_left", lambda: self._call_editor("set_align", "left"))
        c("align_center", lambda: self._call_editor("set_align", "center"))
        c("align_right", lambda: self._call_editor("set_align", "right"))
        c("align_justify", lambda: self._call_editor("set_align", "justify"))
        tc("wrap_text", lambda on: self._call_editor("set_wrap_text", on))
        c("merge", lambda: self._call_editor("merge_selection"))
        c("unmerge", lambda: self._call_editor("unmerge_selection"))
        c("rotation", lambda: self._call_editor("set_rotation"))
        c("indent_up", lambda: self._call_editor("increase_indent"))
        c("indent_down", lambda: self._call_editor("decrease_indent"))
        c("decimal_up", lambda: self._call_editor("increase_decimal"))
        c("decimal_down", lambda: self._call_editor("decrease_decimal"))

        c("undo", self._do_undo)
        c("redo", self._do_redo)
        c("find", self.open_find_dialog)
        c("select_all", self._do_select_all)
        c("format_brush", self._do_format_brush)

        # Excel 单元格
        c("bg_color", lambda: self._call_editor("set_fill_color"))
        c("insert_row", lambda: self._call_editor("insert_row"))
        c("delete_row", lambda: self._call_editor("delete_row"))
        c("insert_col", lambda: self._call_editor("insert_column"))
        c("delete_col", lambda: self._call_editor("delete_column"))
        c("auto_sum", lambda: self._call_editor("auto_sum"))
        c("fill", lambda: self._call_editor("fill_dialog"))
        c("clear_content", lambda: self._call_editor("clear_content"))
        c("sort_asc", lambda: self._call_editor("sort_asc"))
        c("sort_desc", lambda: self._call_editor("sort_desc"))
        c("filter", lambda: self._call_editor("filter_by_keyword"))
        c("clear_filter", lambda: self._call_editor("clear_filter"))
        c("hide_row", lambda: self._call_editor("hide_row"))
        c("hide_col", lambda: self._call_editor("hide_column"))
        c("unhide", self._excel_unhide)

        # 插图 / 工作表
        c("insert_image", lambda: self._call_editor("insert_image"))
        c("insert_table", lambda: self._call_editor("insert_table"))
        c("new_sheet", lambda: self._call_editor("add_sheet"))
        c("del_sheet", lambda: self._call_editor("delete_sheet"))
        c("rename_sheet", lambda: self._call_editor("rename_sheet"))

        # 引用
        c("header", lambda: self._call_editor("insert_header"))
        c("footer", lambda: self._call_editor("insert_footer"))
        c("page_number", lambda: self._call_editor("insert_page_number"))
        c("word_count", lambda: self._call_editor("word_count"))

        # PDF（只读查看）
        c("prev_page", lambda: self._call_editor("prev_page"))
        c("next_page", lambda: self._call_editor("next_page"))
        c("zoom_in", lambda: self._call_editor("zoom_in"))
        c("zoom_out", lambda: self._call_editor("zoom_out"))
        c("rotate", lambda: self._call_editor("rotate_clockwise"))
        c("fit_width", lambda: self._call_editor("fit_width"))

        # 审阅 / 视图
        c("find_review", self.open_find_dialog)
        c("encoding", self._choose_encoding)
        c("zoom_in_view", lambda: self._call_editor("set_font_size", 1))
        c("zoom_out_view", lambda: self._call_editor("set_font_size", -1))

    def _excel_unhide(self):
        """Excel：同时恢复所有隐藏行与隐藏列。"""
        self._call_editor("unhide_row")
        self._call_editor("unhide_column")

    # ---------- 文档管理 ----------
    def current_editor(self):
        """当前标签页中的编辑器；欢迎页时返回 None。"""
        w = self.tabs.currentWidget()
        if w is None or w is self.welcome:
            return None
        return w

    def _ribbon_new(self):
        keys = list(NEW_TYPES.keys())
        labels = [NEW_TYPES[k] for k in keys]
        text, ok = QInputDialog.getItem(self, "新建文档", "选择要新建的文档类型：",
                                        labels, 0, False)
        if ok:
            self.new_document(keys[labels.index(text)])

    def new_document(self, doc_type):
        if doc_type == "excel":
            editor = ExcelEditor()
            name = "未命名表格.xlsx"
            try:
                editor.new_document()
            except Exception as e:
                QMessageBox.critical(self, "新建失败", str(e))
                return
        elif doc_type == "word":
            editor = WordEditor()
            name = "未命名文档.docx"
            try:
                editor.new_document()
            except Exception as e:
                QMessageBox.critical(self, "新建失败", str(e))
                return
        else:
            editor = TextEditor()
            name = "未命名文本.txt"
            editor.new_document()
        self._setup_editor_signals(editor)
        idx = self.tabs.addTab(editor, name)
        self.tabs.setCurrentIndex(idx)

    def open_file(self, paths=None):
        """打开文件：支持多选，按扩展名路由到对应编辑器。

        paths 为 None 时弹出 QFileDialog 多选对话框；否则直接打开传入路径列表。
        """
        if paths is None:
            paths, _ = QFileDialog.getOpenFileNames(self, "打开文件", "", _FILE_FILTER,
                                                    options=QFileDialog.Option.DontUseNativeDialog)
        elif isinstance(paths, str):
            paths = [paths]
        if not paths:
            return
        for path in paths:
            ext = os.path.splitext(path)[1].lower()
            if ext not in SUPPORTED_EXT:
                QMessageBox.warning(self, "提示", f"暂不支持打开 {ext or '无扩展名'} 文件：{path}")
                continue
            editor = self._create_editor_by_ext(ext)
            try:
                editor.load(path)
            except Exception as e:
                QMessageBox.critical(self, "打开失败", f"无法打开文件：\n{path}\n\n{str(e)}")
                continue
            self._setup_editor_signals(editor)
            idx = self.tabs.addTab(editor, os.path.basename(path))
            self.tabs.setCurrentIndex(idx)
            self._add_recent(path)

    def _create_editor_by_ext(self, ext):
        doc_type = SUPPORTED_EXT[ext]
        if doc_type == "excel":
            return ExcelEditor()
        if doc_type == "word":
            return WordEditor()
        if doc_type == "pdf":
            return PdfEditor()
        return TextEditor()

    def _setup_editor_signals(self, editor):
        editor.modified_changed.connect(self._on_modified_changed)
        if hasattr(editor, "cursor_changed"):
            editor.cursor_changed.connect(self._on_editor_cursor_changed)

    def save_file(self):
        editor = self.current_editor()
        if editor is None:
            return
        # PDF 为只读查看，保存一律走另存为
        if editor.doc_type() == "pdf":
            self.save_file_as()
            return
        if editor.file_path is None:
            self.save_file_as()
            return
        try:
            ok = editor.save()
            if not ok:
                self.save_file_as()
        except Exception as e:
            QMessageBox.critical(self, "保存失败", str(e))

    def save_file_as(self):
        editor = self.current_editor()
        if editor is None:
            return
        doc_type = editor.doc_type()
        if doc_type == "excel":
            name, filt = "表格.xlsx", "Excel 表格 (*.xlsx *.xlsm)"
        elif doc_type == "word":
            name, filt = "文档.docx", "Word 文档 (*.docx)"
        elif doc_type == "pdf":
            name, filt = "文档.pdf", "PDF 文件 (*.pdf)"
        else:
            name, filt = "文本.txt", "文本文件 (*.txt);;所有文件 (*.*)"
        start = editor.file_path or os.path.join(os.path.expanduser("~"), name)
        path, _ = QFileDialog.getSaveFileName(self, "另存为", start, filt,
                                              options=QFileDialog.Option.DontUseNativeDialog)
        if not path:
            return
        try:
            ok = editor.save_as(path)
            if ok:
                self.tabs.setTabText(self.tabs.currentIndex(), os.path.basename(editor.file_path))
                self._add_recent(editor.file_path)
                self._update_status(editor)
        except Exception as e:
            QMessageBox.critical(self, "保存失败", str(e))

    # ---------- 标签页管理 ----------
    def _ask_save_changes(self, title, name):
        """关闭未保存文档前的中文确认框。

        返回："save" 保存 / "discard" 不保存 / "cancel" 取消。
        使用自定义中文按钮，避免 Qt 默认英文按钮（Save/Discard/Cancel）。
        """
        box = QMessageBox(self)
        box.setWindowTitle(title)
        box.setIcon(QMessageBox.Icon.Question)
        box.setText(f"文档“{name}”已修改，是否保存？")
        btn_save = box.addButton("保存", QMessageBox.ButtonRole.AcceptRole)
        btn_discard = box.addButton("不保存", QMessageBox.ButtonRole.DestructiveRole)
        btn_cancel = box.addButton("取消", QMessageBox.ButtonRole.RejectRole)
        box.setDefaultButton(btn_save)
        box.exec()
        clicked = box.clickedButton()
        if clicked is btn_save:
            return "save"
        if clicked is btn_discard:
            return "discard"
        return "cancel"

    def _close_tab(self, index):
        if index == 0:  # 欢迎页不可关闭
            return
        editor = self.tabs.widget(index)
        if editor is None:
            return
        if editor.is_modified():
            name = self.tabs.tabText(index)
            ret = self._ask_save_changes("确认", name)
            if ret == "cancel":
                return
            if ret == "save":
                self.tabs.setCurrentIndex(index)
                self.save_file()
                if editor.is_modified():
                    return
        # 淡出后移除标签页
        fade_out_then(editor, lambda ed=editor: self._remove_widget(ed), duration=150)

    def _remove_widget(self, editor):
        """按控件定位并移除标签页（淡出动画期间索引可能变化）。"""
        idx = self.tabs.indexOf(editor)
        if idx > 0:
            self.tabs.removeTab(idx)
        try:
            editor.deleteLater()
        except Exception:
            pass

    def _on_tab_changed(self, index):
        # 旧内容完全退出后再显示新内容：切换瞬间清掉所有非目标标签页
        # 的过渡特效/动画（杜绝新旧编辑器文字同帧渲染重叠）
        for i in range(self.tabs.count()):
            w = self.tabs.widget(i)
            if w is not None and i != index:
                clear_transition(w)
        editor = self.current_editor()
        if editor is None:
            self.ribbon.set_mode(None)
            self._status_doc.setText("未打开文档")
            self._status_info.setText("")
            return
        # 标签切换淡入动画（统一时长常量）
        fade_in(editor)
        doc_type = editor.doc_type()
        self.ribbon.set_mode(doc_type)
        self._status_doc.setText(f"{_DOC_LABEL.get(doc_type, doc_type.upper())}：{editor.file_path or '未命名'}")
        self._update_status(editor)

    def _on_modified_changed(self, modified):
        editor = self.sender()
        if editor is None:
            return
        idx = self.tabs.indexOf(editor)
        if idx <= 0:
            return
        text = self.tabs.tabText(idx)
        if modified and not text.endswith(" *"):
            self.tabs.setTabText(idx, text + " *")
        elif not modified and text.endswith(" *"):
            self.tabs.setTabText(idx, text[:-2])

    def _on_editor_cursor_changed(self, text):
        editor = self.sender()
        if editor is not None and self.tabs.currentWidget() is editor:
            self._status_info.setText(text)

    def _update_status(self, editor):
        try:
            text = editor.status_text()
        except Exception:
            text = ""
        if self.tabs.currentWidget() is editor:
            self._status_info.setText(text)

    # ---------- 动作转发 ----------
    def _call_editor(self, method_name, *args):
        editor = self.current_editor()
        if editor is None:
            return
        method = getattr(editor, method_name, None)
        if method is None:
            return
        try:
            method(*args)
        except Exception as e:
            QMessageBox.warning(self, "操作失败", str(e))
        self._update_status(editor)

    def _do_undo(self):
        self._call_editor("undo")

    def _do_redo(self):
        self._call_editor("redo")

    def _do_cut(self):
        self._call_editor("cut")

    def _do_copy(self):
        self._call_editor("copy")

    def _do_paste(self):
        self._call_editor("paste")

    def _do_select_all(self):
        editor = self.current_editor()
        if editor is None:
            return
        edit = getattr(editor, "edit", None)
        if edit is not None and hasattr(edit, "selectAll"):
            edit.selectAll()

    def _on_font_size_combo(self, text):
        """字号下拉框选择/输入：解析为数值后应用到当前编辑器。"""
        try:
            size = float(text.strip())
        except ValueError:
            return
        self._call_editor("set_font_size_value", max(1.0, min(400.0, size)))

    def _do_format_brush(self):
        """格式刷（两次点击模式）：第一次记录格式，第二次应用到选中文字。"""
        editor = self.current_editor()
        if editor is None:
            return
        btn = self.ribbon.button("format_brush")
        if editor.doc_type() not in ("word", "excel"):
            btn.setChecked(False)
            QMessageBox.information(self, "格式刷", "格式刷仅支持 Word / Excel 文档。")
            return
        if not self._format_brush_active:
            # 第一次点击：记录当前光标处格式
            editor.capture_format_brush()
            self._format_brush_active = True
            btn.setChecked(True)
            self._status_info.setText("已记录格式，请选中目标文字后再次点击“格式刷”应用")
        else:
            # 第二次点击：应用格式到当前选区
            if not editor.apply_format_brush():
                QMessageBox.warning(self, "格式刷", "尚未记录格式，请先点击一次“格式刷”。")
            self._format_brush_active = False
            btn.setChecked(False)
            editor.clear_format_brush()
            self._status_info.setText("")

    def open_find_dialog(self):
        editor = self.current_editor()
        if editor is None:
            return
        text = ""
        edit = getattr(editor, "edit", None)
        if edit is not None and hasattr(edit, "textCursor") and edit.textCursor().hasSelection():
            text = edit.textCursor().selectedText()
        self._find_dialog.show_with_text(text)
        pop_in(self._find_dialog)

    def _on_find_next(self, text, case):
        editor = self.current_editor()
        if editor is None:
            return
        if not editor.find_next(text, case):
            QMessageBox.information(self, "查找", f"未找到“{text}”")

    def _on_replace(self, find, repl, case):
        editor = self.current_editor()
        if editor is None:
            return
        editor.replace_current(find, repl, case)
        self._update_status(editor)

    def _on_replace_all(self, find, repl, case):
        editor = self.current_editor()
        if editor is None:
            return
        count = editor.replace_all(find, repl, case)
        QMessageBox.information(self, "全部替换", f"共替换 {count} 处")

    def _choose_encoding(self):
        editor = self.current_editor()
        if editor is None or editor.doc_type() != "text":
            return
        from app.editors.text_editor import ENCODINGS
        current = editor.get_encoding()
        try:
            idx = ENCODINGS.index(current)
        except ValueError:
            idx = 0
        enc, ok = QInputDialog.getItem(self, "选择编码", "文本编码：", ENCODINGS, idx, False)
        if ok and enc:
            editor.set_encoding(enc)
            self._update_status(editor)

    # ---------- 最近文件 ----------
    def _load_recent_files(self):
        values = self._settings.value("recent_files", [])
        self._recent_files = [v for v in (values or []) if os.path.exists(v)]

    def _save_recent_files(self):
        self._settings.setValue("recent_files", self._recent_files[:8])

    def _add_recent(self, path):
        if path in self._recent_files:
            self._recent_files.remove(path)
        self._recent_files.insert(0, path)
        self._recent_files = self._recent_files[:8]
        self._save_recent_files()
        self._refresh_recent_menu()
        self.welcome.set_recent_files(self._recent_files)

    def _refresh_recent_menu(self):
        self.recent_menu.clear()
        if not self._recent_files:
            act = QAction("（无）", self)
            act.setEnabled(False)
            self.recent_menu.addAction(act)
            return
        for path in self._recent_files:
            act = QAction(path, self)
            act.triggered.connect(lambda checked=False, p=path: self.open_file(p))
            self.recent_menu.addAction(act)

    # ---------- 关闭确认 ----------
    def closeEvent(self, event):
        for idx in range(1, self.tabs.count()):
            editor = self.tabs.widget(idx)
            if editor is not None and editor.is_modified():
                name = self.tabs.tabText(idx)
                ret = self._ask_save_changes("确认退出", name)
                if ret == "cancel":
                    event.ignore()
                    return
                if ret == "save":
                    self.tabs.setCurrentIndex(idx)
                    self.save_file()
                    if editor.is_modified():
                        event.ignore()
                        return
        event.accept()

    def _show_about(self):
        QMessageBox.about(
            self, "关于",
            "素笺 (SuJian)\n\n"
            "基于 Python + PySide6 开发的本地办公软件，支持：\n"
            "· Excel 表格编辑（公式 / 合并 / 排序筛选 / 隐藏 / 数字格式 / 多工作表）\n"
            "· Word 文档编辑（富文本 / 删除线 / 高亮 / 页眉页脚 / 页码 / 字数统计）\n"
            "· TXT 文本编辑（多编码 / 行号 / 查找替换）\n"
            "· PDF 查看（翻页 / 缩放 / 旋转 / 适合宽度）\n\n"
            "仅供学习交流使用。")

# -*- coding: utf-8 -*-
"""SuJian（素笺）图形化安装向导（setup.exe）。

功能（五步向导，仿常见 Windows 安装程序，WPS 绿主题，全中文）：
  第 1 页  欢迎页：显示软件名与版本，按钮“下一步”；
  第 2 页  询问“是否安装素笺？”，提供“安装”与“取消”，点取消则退出；
  第 3 页  选择安装路径（默认 %LOCALAPPDATA%\\SuJian，可浏览选择并校验可写）；
  第 4 页  安装进度：复制主程序 exe → 复制卸载组件 → 创建桌面/开始菜单快捷方式
           → 注册 .xlsx/.xls/.docx/.txt/.pdf 文件关联 → 写入系统卸载入口；
  第 5 页  完成页：默认勾选“立即运行素笺”，点“完成”关闭。

复用 scripts/installer.py 中的注册逻辑（make_shortcut / register_file_assoc /
write_uninstall_entry），并支持自定义安装目录。

打包：见 scripts/setup_gui.spec，运行后产物为 dist\\setup.exe（单文件，
内嵌 dist/SuJian.exe，运行时释放到用户选择的安装目录）。
"""
import os
import shutil
import subprocess
import sys

# 保证能 import 同目录下的 installer.py（开发模式直接运行 / 打包后同包释放）
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import installer  # noqa: E402

from PySide6.QtCore import Qt, QThread, Signal  # noqa: E402
from PySide6.QtGui import QIcon  # noqa: E402
from PySide6.QtWidgets import (  # noqa: E402
    QApplication, QCheckBox, QDialog, QFileDialog, QHBoxLayout, QLabel,
    QLineEdit, QMessageBox, QProgressBar, QPushButton, QStackedWidget,
    QVBoxLayout, QWidget,
)

APP_VERSION = "1.0.0"
APP_TITLE = "素笺"
APP_EXE_NAME = "SuJian.exe"
DEFAULT_INSTALL_DIR = os.path.join(os.environ.get("LOCALAPPDATA", os.path.expanduser("~")),
                                   "SuJian")

# WPS 绿主题色
COLOR_PRIMARY = "#2e9e5b"
COLOR_PRIMARY_DARK = "#23804a"
COLOR_BG = "#f6f8f6"
COLOR_BG_HEADER = "#eef4ef"
COLOR_TEXT = "#333333"
COLOR_MUTED = "#777777"

STYLE_SHEET = f"""
* {{ font-family: "Microsoft YaHei", "微软雅黑"; font-size: 14px; color: {COLOR_TEXT}; }}
QDialog {{ background: {COLOR_BG}; }}
#header {{ background: {COLOR_BG_HEADER}; border-bottom: 1px solid #d9e5dc; }}
#headerTitle {{ font-size: 20px; font-weight: bold; color: {COLOR_PRIMARY_DARK}; }}
#headerSub {{ font-size: 12px; color: {COLOR_MUTED}; }}
#bigQuestion {{ font-size: 17px; font-weight: bold; color: {COLOR_PRIMARY_DARK}; }}
#hint {{ font-size: 12px; color: {COLOR_MUTED}; }}
QPushButton {{
    background: #ffffff; border: 1px solid #c8d6cd; border-radius: 4px;
    padding: 7px 22px; color: {COLOR_TEXT};
}}
QPushButton:hover {{ background: #f0f5f1; border-color: {COLOR_PRIMARY}; }}
QPushButton:pressed {{ background: #e4ede6; }}
QPushButton:disabled {{ background: #f0f0f0; color: #aaaaaa; border-color: #dddddd; }}
QPushButton#primary {{ background: {COLOR_PRIMARY}; border: none; color: white; font-weight: bold; }}
QPushButton#primary:hover {{ background: {COLOR_PRIMARY_DARK}; }}
QPushButton#primary:pressed {{ background: {COLOR_PRIMARY_DARK}; }}
QPushButton#danger {{ background: #ffffff; border: 1px solid #d0a0a0; color: #b03030; }}
QPushButton#danger:hover {{ background: #fdf0f0; border-color: #c07070; }}
QLineEdit {{
    background: white; border: 1px solid #c8d6cd; border-radius: 4px;
    padding: 6px 10px; selection-background-color: {COLOR_PRIMARY};
}}
QLineEdit:focus {{ border-color: {COLOR_PRIMARY}; }}
QProgressBar {{
    border: 1px solid #c8d6cd; border-radius: 6px; background: white;
    text-align: center; font-size: 12px; height: 24px;
}}
QProgressBar::chunk {{ background: {COLOR_PRIMARY}; border-radius: 5px; }}
QCheckBox::indicator {{ width: 18px; height: 18px; }}
QCheckBox::indicator:unchecked {{ border: 1px solid #aab8af; background: white; border-radius: 3px; }}
QCheckBox::indicator:checked {{ background: {COLOR_PRIMARY}; border: 1px solid {COLOR_PRIMARY}; border-radius: 3px; }}
"""


def bundled_exe_path():
    """定位安装程序内嵌的 SuJian.exe。

    打包后：位于 PyInstaller 解压目录 sys._MEIPASS\\dist\\SuJian.exe；
    开发模式：回退到项目根目录 dist\\SuJian.exe。
    """
    base = getattr(sys, "_MEIPASS", None)
    if base:
        for cand in (os.path.join(base, "dist", APP_EXE_NAME),
                     os.path.join(base, APP_EXE_NAME)):
            if os.path.isfile(cand):
                return cand
    project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    fallback = os.path.join(project_root, "dist", APP_EXE_NAME)
    if not os.path.isfile(fallback):
        raise FileNotFoundError("未找到内嵌主程序 SuJian.exe，请先执行 pyinstaller SuJian.spec")
    return fallback


def bundled_icon_path():
    """定位安装向导窗口图标（打包后与开发模式通用）。"""
    base = getattr(sys, "_MEIPASS", None)
    candidates = []
    if base:
        candidates += [os.path.join(base, "build_assets", "icon.ico"),
                       os.path.join(base, "build_assets", "icon.png")]
    candidates += [
        os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                     "build_assets", "icon.ico"),
        os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                     "build_assets", "icon.png"),
    ]
    for c in candidates:
        if os.path.isfile(c):
            return c
    return None


class InstallWorker(QThread):
    """后台安装线程：把安装步骤放到子线程，避免阻塞 GUI。"""
    step = Signal(int, str)   # 进度百分比, 状态文案
    error = Signal(str)
    done = Signal()

    def __init__(self, source_exe, install_dir):
        super().__init__()
        self.source_exe = source_exe
        self.install_dir = install_dir
        self.uninstaller_src = self._locate_uninstaller()
        self.target_exe = os.path.join(install_dir, APP_EXE_NAME)

    def _locate_uninstaller(self):
        """定位卸载组件：优先自包含 uninstaller.exe（不依赖用户机器 Python），
        否则回退 scripts/uninstaller.py 脚本。"""
        base = getattr(sys, "_MEIPASS", None)
        candidates = []
        if base:
            candidates += [os.path.join(base, "dist", "uninstaller.exe")]
        project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        candidates += [os.path.join(project_root, "dist", "uninstaller.exe"),
                       os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                    "uninstaller.py")]
        for c in candidates:
            if os.path.isfile(c):
                return c
        return candidates[-1]

    def run(self):
        try:
            # 1) 复制主程序
            self.step.emit(5, "准备安装目录...")
            os.makedirs(self.install_dir, exist_ok=True)
            self.step.emit(15, "正在复制主程序...")
            shutil.copy2(self.source_exe, self.target_exe)

            # 2) 复制卸载组件（供“应用和功能”后台调用）
            self.step.emit(40, "正在复制卸载组件...")
            if os.path.isfile(self.uninstaller_src):
                uninstaller_name = (os.path.basename(self.uninstaller_src)
                                    if self.uninstaller_src.lower().endswith(".exe")
                                    else "uninstaller.py")
                shutil.copy2(self.uninstaller_src,
                             os.path.join(self.install_dir, uninstaller_name))

            # 3) 创建桌面与开始菜单快捷方式
            self.step.emit(55, "正在创建桌面快捷方式...")
            desktop = os.path.join(os.environ.get("USERPROFILE", ""), "Desktop")
            if os.path.isdir(desktop):
                installer.make_shortcut(os.path.join(desktop, "SuJian.lnk"),
                                        self.target_exe)
            start_menu = os.path.join(os.environ.get("APPDATA", ""),
                                      "Microsoft", "Windows", "Start Menu", "Programs")
            if os.path.isdir(start_menu):
                installer.make_shortcut(os.path.join(start_menu, "SuJian.lnk"),
                                        self.target_exe)

            # 4) 注册文件关联
            self.step.emit(70, "正在注册文件关联（.xlsx/.xls/.docx/.txt/.pdf）...")
            installer.register_file_assoc(self.target_exe)

            # 5) 写入系统卸载入口
            self.step.emit(88, "正在写入系统卸载入口...")
            installer.write_uninstall_entry(self.target_exe, self.install_dir)

            self.step.emit(100, "安装完成")
            self.done.emit()
        except Exception as exc:  # noqa: BLE001 - 向 GUI 展示全部错误
            self.error.emit(str(exc))


class SetupWizard(QDialog):
    """五步安装向导主窗口。"""

    def __init__(self, source_exe):
        super().__init__()
        self.source_exe = source_exe
        self.install_dir = DEFAULT_INSTALL_DIR
        self.worker = None
        self.installed_exe = None
        self._setup_ui()
        self._apply_style()
        self._go(0)

    # ---------- UI 构建 ----------
    def _setup_ui(self):
        self.setWindowTitle(f"安装 - {APP_TITLE}")
        self.setWindowIcon(QIcon(bundled_icon_path() or ""))
        self.resize(600, 440)
        self.setMinimumSize(560, 420)

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        # 顶部标题区
        header = QWidget(self)
        header.setObjectName("header")
        header.setFixedHeight(76)
        h_lay = QVBoxLayout(header)
        h_lay.setContentsMargins(24, 12, 24, 12)
        self.lbl_title = QLabel(f"安装 - {APP_TITLE}", header)
        self.lbl_title.setObjectName("headerTitle")
        self.lbl_sub = QLabel(f"版本 {APP_VERSION}", header)
        self.lbl_sub.setObjectName("headerSub")
        h_lay.addWidget(self.lbl_title)
        h_lay.addWidget(self.lbl_sub)
        root.addWidget(header)

        # 页面区
        self.stack = QStackedWidget(self)
        self.pages = [
            self._build_welcome_page(),
            self._build_ask_page(),
            self._build_path_page(),
            self._build_progress_page(),
            self._build_finish_page(),
        ]
        for p in self.pages:
            self.stack.addWidget(p)
        root.addWidget(self.stack, 1)

        # 底部按钮栏
        btns = QHBoxLayout()
        btns.setContentsMargins(16, 10, 16, 14)
        btns.addStretch(1)
        self.btn_back = QPushButton("上一步", self)
        self.btn_next = QPushButton("下一步", self)
        self.btn_next.setObjectName("primary")
        self.btn_finish = QPushButton("完成", self)
        self.btn_finish.setObjectName("primary")
        self.btn_cancel = QPushButton("取消", self)
        self.btn_back.clicked.connect(lambda: self._go(self.stack.currentIndex() - 1))
        self.btn_next.clicked.connect(self._on_next_clicked)
        self.btn_finish.clicked.connect(self.accept)
        self.btn_cancel.clicked.connect(self.reject)
        for b in (self.btn_back, self.btn_next, self.btn_finish, self.btn_cancel):
            btns.addWidget(b)
        root.addLayout(btns)

    def _page_container(self):
        w = QWidget(self)
        lay = QVBoxLayout(w)
        lay.setContentsMargins(36, 26, 36, 20)
        lay.setSpacing(14)
        return w, lay

    def _build_welcome_page(self):
        w, lay = self._page_container()
        lay.addStretch(1)
        title = QLabel("欢迎使用素笺", w)
        title.setObjectName("bigQuestion")
        title.setAlignment(Qt.AlignCenter)
        lay.addWidget(title)
        desc = QLabel(
            "本安装向导将帮助您把素笺安装到电脑上。\n"
            "软件支持打开与编辑 Excel / Word / TXT / PDF 文档，\n"
            "安装后双击文档即可使用本软件打开。", w)
        desc.setObjectName("hint")
        desc.setAlignment(Qt.AlignCenter)
        lay.addWidget(desc)
        lay.addStretch(2)
        return w

    def _build_ask_page(self):
        w, lay = self._page_container()
        lay.addStretch(1)
        q = QLabel("是否安装素笺？", w)
        q.setObjectName("bigQuestion")
        q.setAlignment(Qt.AlignCenter)
        lay.addWidget(q)
        tip = QLabel("点击“安装”开始安装，点击“取消”退出安装向导。", w)
        tip.setObjectName("hint")
        tip.setAlignment(Qt.AlignCenter)
        lay.addWidget(tip)
        lay.addStretch(1)
        row = QHBoxLayout()
        row.addStretch(1)
        btn_install = QPushButton("安装", w)
        btn_install.setObjectName("primary")
        btn_install.setMinimumWidth(120)
        btn_install.clicked.connect(lambda: self._go(2))
        btn_cancel = QPushButton("取消", w)
        btn_cancel.setObjectName("danger")
        btn_cancel.setMinimumWidth(120)
        btn_cancel.clicked.connect(self.reject)
        row.addWidget(btn_install)
        row.addSpacing(16)
        row.addWidget(btn_cancel)
        row.addStretch(1)
        lay.addLayout(row)
        lay.addSpacing(10)
        return w

    def _build_path_page(self):
        w, lay = self._page_container()
        lay.addStretch(1)
        q = QLabel("选择安装位置", w)
        q.setObjectName("bigQuestion")
        lay.addWidget(q)
        tip = QLabel("选择要将素笺安装到的文件夹：", w)
        tip.setObjectName("hint")
        lay.addWidget(tip)

        row = QHBoxLayout()
        self.edit_path = QLineEdit(self.install_dir, w)
        btn_browse = QPushButton("浏览...", w)
        btn_browse.clicked.connect(self._on_browse)
        row.addWidget(self.edit_path, 1)
        row.addWidget(btn_browse)
        lay.addLayout(row)

        self.lbl_path_check = QLabel("", w)
        self.lbl_path_check.setObjectName("hint")
        lay.addWidget(self.lbl_path_check)
        lay.addStretch(2)
        return w

    def _build_progress_page(self):
        w, lay = self._page_container()
        lay.addStretch(1)
        q = QLabel("正在安装...", w)
        q.setObjectName("bigQuestion")
        q.setAlignment(Qt.AlignCenter)
        lay.addWidget(q)
        self.progress = QProgressBar(w)
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self.progress.setFormat("%p%")
        lay.addWidget(self.progress)
        self.lbl_progress = QLabel("准备开始安装...", w)
        self.lbl_progress.setAlignment(Qt.AlignCenter)
        self.lbl_progress.setObjectName("hint")
        lay.addWidget(self.lbl_progress)
        lay.addStretch(2)
        return w

    def _build_finish_page(self):
        w, lay = self._page_container()
        lay.addStretch(1)
        q = QLabel("安装完成", w)
        q.setObjectName("bigQuestion")
        q.setAlignment(Qt.AlignCenter)
        lay.addWidget(q)
        tip = QLabel("素笺已成功安装到您的电脑。", w)
        tip.setObjectName("hint")
        tip.setAlignment(Qt.AlignCenter)
        lay.addWidget(tip)
        self.chk_run = QCheckBox("立即运行素笺", w)
        self.chk_run.setChecked(True)
        self.chk_run.setStyleSheet("font-size: 14px;")
        lay.addWidget(self.chk_run, 0, Qt.AlignCenter)
        lay.addStretch(2)
        return w

    # ---------- 行为 ----------
    def _apply_style(self):
        self.setStyleSheet(STYLE_SHEET)

    def _go(self, index):
        """切换到指定页并刷新底部按钮状态。"""
        self.stack.setCurrentIndex(index)
        cur = self.stack.currentIndex()
        self.btn_back.setVisible(cur in (2,))
        self.btn_next.setVisible(cur in (0, 2))
        self.btn_finish.setVisible(cur == 4)
        self.btn_cancel.setVisible(cur != 1)
        self.btn_next.setText("下一步" if cur != 3 else "下一步")

        if cur == 2:
            self.edit_path.setText(self.install_dir)
            self._validate_path()
        if cur == 3:
            self._start_install()

    def _on_next_clicked(self):
        cur = self.stack.currentIndex()
        if cur == 0:
            self._go(1)
        elif cur == 2:
            if not self._validate_path():
                return
            self.install_dir = self.edit_path.text().strip() or DEFAULT_INSTALL_DIR
            self._go(3)

    def _validate_path(self):
        """校验安装路径是否可写；不存在则尝试创建。"""
        path = self.edit_path.text().strip() or DEFAULT_INSTALL_DIR
        ok = True
        msg = ""
        try:
            os.makedirs(path, exist_ok=True)
            probe = os.path.join(path, ".wps_setup_write_test")
            with open(probe, "w", encoding="utf-8") as f:
                f.write("ok")
            os.remove(probe)
            msg = f"将安装到：{path}"
        except Exception as exc:  # noqa: BLE001
            ok = False
            msg = f"该位置不可写：{exc}"
        self.lbl_path_check.setText(msg)
        self.btn_next.setEnabled(ok)
        return ok

    def _on_browse(self):
        chosen = QFileDialog.getExistingDirectory(
            self, "选择安装目录", self.edit_path.text().strip() or DEFAULT_INSTALL_DIR)
        if chosen:
            self.edit_path.setText(chosen)
            self._validate_path()

    def _start_install(self):
        self.progress.setValue(0)
        self.lbl_progress.setText("准备开始安装...")
        self.worker = InstallWorker(self.source_exe, self.install_dir)
        self.worker.step.connect(self._on_step)
        self.worker.error.connect(self._on_install_error)
        self.worker.done.connect(self._on_install_done)
        self.worker.start()

    def _on_step(self, percent, text):
        self.progress.setValue(percent)
        self.lbl_progress.setText(text)

    def _on_install_error(self, message):
        QMessageBox.critical(self, "安装失败", f"安装过程中出现错误：\n{message}")
        self._go(0)

    def _on_install_done(self):
        self.installed_exe = os.path.join(self.install_dir, APP_EXE_NAME)
        self.progress.setValue(100)
        self._go(4)

    def accept(self):
        # 完成页：若勾选则立即启动主程序
        if self.stack.currentIndex() == 4 and self.chk_run.isChecked() and self.installed_exe:
            try:
                subprocess.Popen([self.installed_exe], cwd=self.install_dir)
            except Exception:  # noqa: BLE001 - 启动失败不阻塞完成
                pass
        super().accept()


def main():
    app = QApplication(sys.argv)
    app.setApplicationName(APP_TITLE)
    source_exe = bundled_exe_path()
    wizard = SetupWizard(source_exe)
    wizard.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()

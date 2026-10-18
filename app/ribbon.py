# -*- coding: utf-8 -*-
"""仿 WPS 的精细 Ribbon 工具栏。

结构：顶部标签页（开始 / 插入 / 页面 / 审阅 / 视图）+ 下方分组面板。
- 每个分组带浅色圆角容器（分组边框），组内是「图标文字按钮」横排；
- 按钮悬停 / 按下 / 选中均有颜色过渡动画（QVariantAnimation，120ms OutCubic，复用对象不重复新建）；
- 标签页切换带淡入动画（QPropertyAnimation opacity，150ms OutCubic），底部绿色指示条。
内容区支持横向滚动，避免窗口过窄时按钮挤压。
"""
from PySide6.QtCore import Qt, QVariantAnimation, QPropertyAnimation, QEasingCurve
from PySide6.QtGui import QColor, QFontDatabase
from PySide6.QtWidgets import (
    QWidget, QTabBar, QStackedWidget, QVBoxLayout, QHBoxLayout,
    QToolButton, QLabel, QFrame, QScrollArea, QSizePolicy, QComboBox,
)
from app.animation import fade_in


class AnimatedToolButton(QToolButton):
    """带悬停 / 选中背景色过渡动画的工具按钮，模拟 WPS 细腻反馈。

    统一浅色 WPS 绿主题；文字完整显示（ToolButtonTextBesideIcon + 按文字自适应宽度），
    避免按钮文字被截断为省略号。
    """

    _instances = set()

    def __init__(self, text="", parent=None):
        super().__init__(parent)
        self.setText(text)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        # 图标 + 文字并排显示，文字不省略；宽度按文字自适应
        self.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextBesideIcon)
        self.setSizePolicy(QSizePolicy.Policy.Minimum, QSizePolicy.Policy.Fixed)
        self.setFixedHeight(36)
        text_width = self.fontMetrics().horizontalAdvance(text)
        self.setMinimumWidth(max(48, text_width + 20))
        self._base, self._hover, self._checked = (
            QColor("#fbfcfb"), QColor("#e6f4ec"), QColor("#cdeedd"))
        self._current = QColor(self._base)
        self._anim = None
        self._apply(self._base)
        AnimatedToolButton._instances.add(self)

    def destroy(self, destroyWindow=True, destroySubWindows=True):
        AnimatedToolButton._instances.discard(self)
        super().destroy(destroyWindow, destroySubWindows)

    # ---------- 样式 ----------
    def _sheet_for(self, color):
        return (
            "QToolButton {"
            "  background-color: %s;"
            "  border: 1px solid transparent;"
            "  border-radius: 4px;"
            "  padding: 3px 6px;"
            "  font-size: 12px;"
            "  color: #333333;"
            "}"
            "QToolButton:hover { border-color: #9ed3b2; }"
            "QToolButton:pressed { background-color: #d3ebdc; }"
            "QToolButton:checked {"
            "  background-color: #cdeedd;"
            "  border-color: #2e9e5b;"
            "  color: #1d7a44;"
            "  font-weight: bold;"
            "}"
            "QToolButton:disabled { color: #b8b8b8; background: transparent; }"
        ) % color.name()

    def _apply(self, color):
        self._current = color
        self.setStyleSheet(self._sheet_for(color))

    def _forget_anim(self, anim):
        """动画结束后清除引用，避免悬垂引用（访问已删除的 C++ 对象会闪退）。"""
        if self._anim is anim:
            self._anim = None

    def _animate_to(self, target):
        """复用同一个 QVariantAnimation 对象，先停旧再启新，避免叠加闪烁。"""
        if self._anim is None:
            self._anim = QVariantAnimation(self)
            self._anim.valueChanged.connect(self._apply)
            # 动画结束主动解除引用，不依赖 DeleteWhenStopped 避免悬垂
            self._anim.finished.connect(
                lambda a=self._anim: self._forget_anim(a))
        else:
            try:
                self._anim.stop()
            except RuntimeError:
                pass
        anim = self._anim
        start = QColor(self._current)
        anim.setDuration(120)
        anim.setStartValue(start)
        anim.setEndValue(target)
        anim.setEasingCurve(QEasingCurve.Type.OutCubic)
        anim.start()

    # ---------- 事件 ----------
    def enterEvent(self, event):
        self._animate_to(self._checked if self.isChecked() else self._hover)
        super().enterEvent(event)

    def leaveEvent(self, event):
        self._animate_to(self._checked if self.isChecked() else self._base)
        super().leaveEvent(event)

    def nextCheckState(self):
        super().nextCheckState()
        # 选中状态切换后立即刷新底色
        self._animate_to(self._checked if self.isChecked() else self._base)


class _ActionGroup(QFrame):
    """Ribbon 中的一个功能分组：小标题 + 一行工具按钮。"""

    def __init__(self, title, parent=None):
        super().__init__(parent)
        self.setObjectName("wpsGroupBox")
        outer = QVBoxLayout(self)
        outer.setContentsMargins(10, 4, 10, 4)
        outer.setSpacing(2)

        self.title_label = QLabel(title)
        self.title_label.setObjectName("groupTitle")
        outer.addWidget(self.title_label)

        self.buttons_layout = QHBoxLayout()
        self.buttons_layout.setContentsMargins(0, 0, 0, 0)
        self.buttons_layout.setSpacing(2)
        outer.addLayout(self.buttons_layout)

    def add_button(self, text, checkable=False, tooltip=""):
        """新增一个工具按钮并返回。"""
        btn = AnimatedToolButton(text, self)
        btn.setCheckable(checkable)
        if tooltip:
            btn.setToolTip(tooltip)
        self.buttons_layout.addWidget(btn)
        return btn

    def add_combo(self, items=None, editable=False, tooltip="", min_width=100):
        """新增一个下拉框（字体 / 字号 / 数字格式等）并返回。"""
        combo = QComboBox(self)
        combo.setEditable(editable)
        combo.setMinimumWidth(min_width)
        combo.setFixedHeight(30)
        combo.setToolTip(tooltip)
        if items:
            combo.addItems(items)
        self.buttons_layout.addWidget(combo)
        return combo


class WpsRibbon(QWidget):
    """Ribbon 主组件：标签栏 + 滚动分组面板。"""

    #: 标签页顺序（WPS 风格）
    PAGE_NAMES = ["开始", "插入", "页面", "引用", "审阅", "视图"]

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("wpsRibbon")
        self.buttons = {}          # 名称 -> QToolButton
        self._page_groups = {}     # 页面名 -> [group, ...]

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        # 顶部标签栏（底部绿色指示条由 QSS 绘制）
        self.tab_bar = QTabBar()
        self.tab_bar.setExpanding(True)
        self.tab_bar.setDrawBase(False)
        outer.addWidget(self.tab_bar)

        # 页面堆栈（索引与标签一一对应）
        self.stack = QStackedWidget()
        outer.addWidget(self.stack, 1)

        # 当前正在构建的页面名（_add_page 设置，_add_group 据此登记归属）
        self._current_page_name = None

        # 构建各标签页内容
        self._build_start()
        self._build_insert()
        self._build_page()
        self._build_reference()
        self._build_review()
        self._build_view()
        self.tab_bar.currentChanged.connect(self._on_page_changed)

    # ---------- 页面构建 ----------
    def _add_page(self, name):
        # 滚动区（内容过宽时可横向滚动）
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)

        page = QWidget()
        lay = QHBoxLayout(page)
        lay.setContentsMargins(8, 4, 8, 4)
        lay.setSpacing(0)
        lay.addStretch(1)
        scroll.setWidget(page)

        self._page_groups[name] = []
        self._current_page_name = name
        self.tab_bar.addTab(name)
        self.stack.addWidget(scroll)
        return lay

    def _add_group(self, layout, title):
        group = _ActionGroup(title)
        layout.insertWidget(layout.count() - 1, group)  # 保持尾部 stretch
        self._page_groups[self._current_page_name].append(group)
        return group

    def _reg_button(self, group, name, *args, **kwargs):
        """在分组中添加按钮并注册到全局按钮表。"""
        btn = group.add_button(*args, **kwargs)
        self.buttons[name] = btn
        return btn

    # ---------- 构建全部页面 ----------
    def _build_start(self):
        lay = self._add_page("开始")

        # 文件组
        g = self._add_group(lay, "文件")
        self._reg_button(g, "new", "新建", tooltip="新建空白文档")
        self._reg_button(g, "open", "打开", tooltip="打开文件 (Ctrl+O)")
        self._reg_button(g, "save", "保存", tooltip="保存 (Ctrl+S)")
        self._reg_button(g, "save_as", "另存为", tooltip="另存为 (Ctrl+Shift+S)")
        self._reg_button(g, "print_preview", "打印", tooltip="打印 / 打印预览")

        # 剪贴板组
        g = self._add_group(lay, "剪贴板")
        self._reg_button(g, "cut", "剪切", tooltip="剪切 (Ctrl+X)")
        self._reg_button(g, "copy", "复制", tooltip="复制 (Ctrl+C)")
        self._reg_button(g, "paste", "粘贴", tooltip="粘贴 (Ctrl+V)")
        self._reg_button(g, "format_brush", "格式刷", checkable=True,
                         tooltip="格式刷：先点击记录格式，再点击应用到目标文字")

        # 字体组
        g = self._add_group(lay, "字体")
        fam = QFontDatabase.families() or ["微软雅黑", "宋体", "黑体", "Arial"]
        combo = g.add_combo(items=fam, editable=False, tooltip="字体（系统全部字体）",
                            min_width=140)
        combo.setCurrentText("微软雅黑" if "微软雅黑" in fam else fam[0])
        self.buttons["font_family"] = combo
        size_items = [str(i) for i in range(8, 73)] + ["80", "96", "120", "144", "200", "300", "400"]
        combo = g.add_combo(items=size_items, editable=True, tooltip="字号（可输入自定义值）",
                            min_width=64)
        combo.setCurrentText("11")
        self.buttons["font_size"] = combo
        self._reg_button(g, "font_up", "增大", tooltip="增大字号")
        self._reg_button(g, "font_down", "减小", tooltip="减小字号")
        self._reg_button(g, "bold", "加粗", checkable=True, tooltip="加粗 (Ctrl+B)")
        self._reg_button(g, "italic", "斜体", checkable=True, tooltip="斜体 (Ctrl+I)")
        self._reg_button(g, "underline", "下划线", checkable=True, tooltip="下划线 (Ctrl+U)")
        self._reg_button(g, "strikethrough", "删除线", checkable=True, tooltip="删除线")
        self._reg_button(g, "font_color", "字体色", tooltip="字体颜色")
        self._reg_button(g, "fill_color", "填充色", tooltip="单元格 / 文字填充颜色")
        self._reg_button(g, "border", "边框", tooltip="给选中单元格设置边框")

        # 对齐组
        g = self._add_group(lay, "对齐")
        self._reg_button(g, "align_left", "左对齐", tooltip="左对齐")
        self._reg_button(g, "align_center", "居中", tooltip="水平居中")
        self._reg_button(g, "align_right", "右对齐", tooltip="右对齐")
        self._reg_button(g, "align_justify", "两端", tooltip="两端对齐")
        self._reg_button(g, "wrap_text", "自动换行", checkable=True,
                         tooltip="选中单元格 / 段落自动换行")
        self._reg_button(g, "merge", "合并", tooltip="合并选中单元格")
        self._reg_button(g, "unmerge", "取消合并", tooltip="取消合并")
        self._reg_button(g, "rotation", "旋转", tooltip="旋转文本方向")
        self._reg_button(g, "indent_up", "缩进+", tooltip="增加缩进")
        self._reg_button(g, "indent_down", "缩进-", tooltip="减少缩进")

        # 数字组
        g = self._add_group(lay, "数字")
        num_items = ["常规", "数值", "货币", "会计专用", "日期", "百分比",
                     "分数", "科学计数", "文本"]
        combo = g.add_combo(items=num_items, editable=False, tooltip="数字格式",
                            min_width=90)
        combo.setCurrentText("常规")
        self.buttons["num_format"] = combo
        self._reg_button(g, "decimal_down", "减少小数", tooltip="减少小数位")
        self._reg_button(g, "decimal_up", "增加小数", tooltip="增加小数位")

        # 单元格组（Excel）
        g = self._add_group(lay, "单元格")
        self._reg_button(g, "insert_row", "插入行", tooltip="插入行")
        self._reg_button(g, "delete_row", "删除行", tooltip="删除行")
        self._reg_button(g, "insert_col", "插入列", tooltip="插入列")
        self._reg_button(g, "delete_col", "删除列", tooltip="删除列")
        self._reg_button(g, "bg_color", "背景色", tooltip="单元格背景色")
        self._reg_button(g, "hide_row", "隐藏行", tooltip="隐藏选中行")
        self._reg_button(g, "hide_col", "隐藏列", tooltip="隐藏选中列")
        self._reg_button(g, "unhide", "取消隐藏", tooltip="显示所有隐藏行 / 列")

        # 编辑组
        g = self._add_group(lay, "编辑")
        self._reg_button(g, "auto_sum", "自动求和", tooltip="插入 SUM 公式")
        self._reg_button(g, "fill", "填充", tooltip="向下 / 向右填充选中区域")
        self._reg_button(g, "clear_content", "清除内容", tooltip="清空选中单元格内容")
        self._reg_button(g, "sort_asc", "升序", tooltip="按当前列升序排序")
        self._reg_button(g, "sort_desc", "降序", tooltip="按当前列降序排序")
        self._reg_button(g, "filter", "筛选", tooltip="按当前列关键词筛选")
        self._reg_button(g, "clear_filter", "取消筛选", tooltip="显示全部行")
        self._reg_button(g, "find", "查找选择", tooltip="查找与替换 (Ctrl+F)")
        self._reg_button(g, "select_all", "全选", tooltip="全选 (Ctrl+A)")
        self._reg_button(g, "undo", "撤销", tooltip="撤销 (Ctrl+Z)")
        self._reg_button(g, "redo", "重做", tooltip="重做 (Ctrl+Y)")

    def _build_insert(self):
        lay = self._add_page("插入")

        # 插图组
        g = self._add_group(lay, "插图 (Word)")
        self._reg_button(g, "insert_image", "插入图片", tooltip="在文档中插入图片")
        self._reg_button(g, "insert_table", "插入表格", tooltip="插入 3x3 表格")

        # 工作表组
        g = self._add_group(lay, "工作表 (Excel)")
        self._reg_button(g, "new_sheet", "新建表", tooltip="新建工作表")
        self._reg_button(g, "del_sheet", "删除表", tooltip="删除当前工作表")
        self._reg_button(g, "rename_sheet", "重命名", tooltip="重命名当前工作表")

    def _build_page(self):
        lay = self._add_page("页面")

        # PDF 浏览组（只读查看：翻页 / 缩放 / 旋转 / 适合宽度）
        g = self._add_group(lay, "PDF 浏览")
        self._reg_button(g, "prev_page", "上一页", tooltip="上一页")
        self._reg_button(g, "next_page", "下一页", tooltip="下一页")
        self._reg_button(g, "zoom_in", "放大", tooltip="放大")
        self._reg_button(g, "zoom_out", "缩小", tooltip="缩小")
        self._reg_button(g, "rotate", "旋转", tooltip="顺时针旋转 90 度")
        self._reg_button(g, "fit_width", "适合宽度", tooltip="按视口宽度缩放页面")

    def _build_reference(self):
        lay = self._add_page("引用")

        # 页眉页脚组（Word）
        g = self._add_group(lay, "页眉页脚")
        self._reg_button(g, "header", "页眉", tooltip="设置页眉文字")
        self._reg_button(g, "footer", "页脚", tooltip="设置页脚文字")
        self._reg_button(g, "page_number", "页码", tooltip="页脚插入自动页码")

        # 工具组
        g = self._add_group(lay, "工具")
        self._reg_button(g, "word_count", "字数统计", tooltip="统计字数与段落数")

    def _build_review(self):
        lay = self._add_page("审阅")

        # 校对组
        g = self._add_group(lay, "校对")
        self._reg_button(g, "find_review", "查找替换", tooltip="查找与替换 (Ctrl+F)")

    def _build_view(self):
        lay = self._add_page("视图")

        # 文本组
        g = self._add_group(lay, "文本")
        self._reg_button(g, "encoding", "编码", tooltip="选择文本编码方式")

        # 显示组
        g = self._add_group(lay, "显示")
        self._reg_button(g, "zoom_in_view", "放大", tooltip="放大")
        self._reg_button(g, "zoom_out_view", "缩小", tooltip="缩小")

    # ---------- 页面切换动画 ----------
    def _on_page_changed(self, index):
        # 旧内容完全退出后再显示新内容：先终止并卸载所有非目标页面的
        # 过渡特效/动画（清掉可能残留的半透明叠影），再切换并淡入新页。
        # 避免功能区文字（字体/字号/加粗/斜体）在新旧页面同帧渲染时重叠。
        for i in range(self.stack.count()):
            w = self.stack.widget(i)
            if w is not None and i != index:
                clear_transition(w)
        # 核心：先把堆栈切到目标页，再做淡入动画，否则点击标签不会切换内容
        self.stack.setCurrentIndex(index)
        w = self.stack.widget(index)
        if w is not None:
            fade_in(w, duration=150, start=0.4, end=1.0)

    # ---------- 模式切换 ----------
    def set_mode(self, doc_type):
        """按当前文档类型启用 / 禁用各功能按钮。"""
        if doc_type is None:
            self._enable_all(False)
            return

        is_excel = doc_type == "excel"
        is_word = doc_type == "word"
        is_text = doc_type == "text"
        is_pdf = doc_type == "pdf"

        # 默认全部禁用，再按类型打开对应能力
        self._enable_all(False)

        # 通用：文件组 + 剪贴板（PDF 不支持剪切/粘贴）
        self.buttons["new"].setEnabled(True)
        self.buttons["open"].setEnabled(True)
        self.buttons["save"].setEnabled(True)
        self.buttons["save_as"].setEnabled(True)
        self.buttons["print_preview"].setEnabled(not is_pdf)
        for name in ("cut", "copy", "paste"):
            self.buttons[name].setEnabled(not is_pdf)

        # 富文本格式（Word / Excel）
        for name in ("bold", "italic", "underline", "strikethrough",
                     "font_color", "fill_color", "font_up", "font_down",
                     "align_left", "align_center", "align_right", "align_justify"):
            self.buttons[name].setEnabled(is_word or is_excel)
        # 边框：仅 Excel 单元格语义
        self.buttons["border"].setEnabled(is_excel)
        # 字体 / 字号下拉框（Word / Excel）
        for name in ("font_family", "font_size"):
            self.buttons[name].setEnabled(is_word or is_excel)

        # 对齐增强 / 合并 / 旋转 / 缩进：仅 Excel（Word 无对应操作，避免按钮无反应）
        for name in ("wrap_text", "merge", "unmerge", "rotation",
                     "indent_up", "indent_down"):
            self.buttons[name].setEnabled(is_excel)
        for name in ("num_format", "decimal_up", "decimal_down"):
            self.buttons[name].setEnabled(is_excel)

        # 编辑类
        for name in ("undo", "redo", "select_all"):
            self.buttons[name].setEnabled(is_word or is_excel or is_text)
        self.buttons["find"].setEnabled(is_word or is_excel or is_text)
        self.buttons["find_review"].setEnabled(is_word or is_excel or is_text)
        # 格式刷（仅 Word / Excel 富文本）
        self.buttons["format_brush"].setEnabled(is_word or is_excel)

        # 单元格类（仅 Excel）
        for name in ("bg_color", "insert_row", "delete_row", "insert_col",
                     "delete_col", "auto_sum", "fill", "clear_content",
                     "sort_asc", "sort_desc", "filter", "clear_filter",
                     "hide_row", "hide_col", "unhide",
                     "new_sheet", "del_sheet", "rename_sheet"):
            self.buttons[name].setEnabled(is_excel)

        # 插图（仅 Word）
        for name in ("insert_image", "insert_table"):
            self.buttons[name].setEnabled(is_word)

        # 引用页（仅 Word）
        for name in ("header", "footer", "page_number", "word_count"):
            self.buttons[name].setEnabled(is_word)

        # PDF 浏览（只读查看：翻页 / 缩放 / 旋转 / 适合宽度）
        for name in ("prev_page", "next_page", "zoom_in", "zoom_out",
                     "rotate", "fit_width"):
            self.buttons[name].setEnabled(is_pdf)

        # 视图页：编码（仅文本），放大缩小（Word / Excel 字号增减）
        self.buttons["encoding"].setEnabled(is_text)
        self.buttons["zoom_in_view"].setEnabled(is_word or is_excel)
        self.buttons["zoom_out_view"].setEnabled(is_word or is_excel)

    def _enable_all(self, enabled):
        for btn in self.buttons.values():
            btn.setEnabled(enabled)

    def button(self, name):
        return self.buttons.get(name)

    def set_checked(self, name, checked):
        btn = self.buttons.get(name)
        if btn is not None:
            btn.setChecked(checked)

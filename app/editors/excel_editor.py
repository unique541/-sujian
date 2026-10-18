# -*- coding: utf-8 -*-
"""Excel 表格编辑器：基于 openpyxl + QTableWidget。

功能：单元格编辑、行列插入删除、单元格合并/取消合并、字体/背景色/对齐格式、
公式编辑（输入 = 开头即为公式）、多工作表切换与增删改名、查找替换、保存。
支持 .xlsx / .xlsm；.xls 通过 xlrd 读取后以 .xlsx 形式保存。
"""
import os

from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QColor, QFont, QBrush, QFontDatabase
from PySide6.QtWidgets import (
    QVBoxLayout, QTabWidget, QTableWidget, QTableWidgetItem,
    QColorDialog, QInputDialog, QMessageBox, QMenu, QStyledItemDelegate,
    QStyle, QStyleOptionViewItem,
)

from .base_editor import BaseEditor

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font as XlFont, PatternFill, Alignment as XlAlignment
    from openpyxl.styles import Border as XlBorder, Side as XlSide
    XL_OK = True
except ImportError:
    XL_OK = False

try:
    import xlrd
    HAS_XLRD = True
except ImportError:
    HAS_XLRD = False


def _qcolor_from_openpyxl(rgb):
    """openpyxl 颜色字符串 -> QColor；失败返回 None。"""
    if not rgb:
        return None
    s = str(rgb)
    # 格式可能为 FFRRGGBB / RRGGBB / 空
    if len(s) >= 6:
        hex_part = s[-6:]
        try:
            return QColor("#" + hex_part)
        except Exception:
            return None
    return None


def _openpyxl_color(qcolor):
    """QColor -> openpyxl 需要的 FFRRGGBB 字符串。"""
    if qcolor is None:
        return None
    r, g, b = qcolor.red(), qcolor.green(), qcolor.blue()
    return f"{r:02X}{g:02X}{b:02X}"


def _to_number(text):
    """把单元格文本尽量转换为 int / float，失败则保留原文。"""
    if text is None:
        return None
    s = text.strip()
    if s == "":
        return None
    try:
        if s.lstrip("-").isdigit():
            return int(s)
        f = float(s)
        if f.is_integer() and "." not in s:
            return int(f)
        return f
    except ValueError:
        return text


def _col_name(index):
    """列号（1-based）-> Excel 列字母，如 1->A, 27->AA。"""
    letters = ""
    n = index
    while n > 0:
        n, rem = divmod(n - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


#: 新建 / 加载时显示的默认网格行列数（仿 Excel/WPS 完整网格外观）
#: 列头显示 A~Z、行头显示 1~30；保存时只写有内容的单元格，不影响文件
DEFAULT_GRID_ROWS = 30
DEFAULT_GRID_COLS = 26

#: 滚动到底时每次追加的行 / 列数（保证网格持续可滚动）
APPEND_ROWS_ON_SCROLL = 50
APPEND_COLS_ON_SCROLL = 10


class _ExcelItemDelegate(QStyledItemDelegate):
    """支持旋转文本与缩进的单元格绘制委托。

    读取 editor._styles[sheet][(row, col)] 中的 rotation / indent 字段：
    - rotation: 角度（90 / -90 / 45），旋转绘制文本；
    - indent: 缩进级别（int），文本左侧偏移。
    """

    def __init__(self, view, editor, parent=None):
        super().__init__(parent)
        self._view = view
        self._editor = editor

    def _style_for(self, index):
        editor = self._editor
        try:
            sheet = editor.sheet_tabs.tabText(editor.sheet_tabs.currentIndex())
        except Exception:
            return {}
        return editor._styles.get(sheet, {}).get(
            (index.row() + 1, index.column() + 1), {})

    def paint(self, painter, option, index):
        st = self._style_for(index)
        rotation = st.get("rotation") or 0
        indent = int(st.get("indent") or 0)
        border = st.get("border")
        if not rotation and indent <= 0 and not border:
            super().paint(painter, option, index)
            return

        # 先画默认背景 / 选中高亮
        opt = QStyleOptionViewItem(option)
        self.initStyleOption(opt, index)
        style = opt.widget.style() if opt.widget else option.widget.style()
        style.drawControl(QStyle.ControlElement.CE_ItemViewItem, opt, painter, opt.widget)

        rect = opt.rect
        if border:
            painter.save()
            pen = painter.pen()
            pen.setColor(QColor(border))
            pen.setWidth(1)
            painter.setPen(pen)
            painter.drawRect(rect.adjusted(0, 0, -1, -1))
            painter.restore()

        text = opt.text
        if not text:
            return
        if opt.state & QStyle.StateFlag.State_Selected:
            painter.setPen(opt.palette.highlightedText().color())
        else:
            painter.setPen(opt.palette.text().color())
        painter.setFont(opt.font)

        if rotation:
            painter.save()
            painter.translate(rect.center())
            painter.rotate(rotation)
            draw_rect = QRect(-rect.width() // 2, -rect.height() // 2,
                              rect.width(), rect.height())
            painter.drawText(draw_rect,
                             int(opt.displayAlignment | Qt.AlignmentFlag.AlignCenter),
                             text)
            painter.restore()
        else:
            pad = indent * 8
            draw_rect = QRect(rect.x() + pad, rect.y(),
                              max(1, rect.width() - pad), rect.height())
            painter.drawText(draw_rect, int(opt.displayAlignment), text)


class ExcelEditor(BaseEditor):
    """Excel 编辑器：内存中维护 openpyxl 工作簿，界面用 QTableWidget 编辑。"""

    #: 当前单元格位置变化信号（用于状态栏）
    cursor_changed = Signal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        self.sheet_tabs = QTabWidget()
        self.sheet_tabs.setTabsClosable(False)
        layout.addWidget(self.sheet_tabs)

        self.workbook = None
        self._sheets = {}     # 工作表名 -> QTableWidget
        self._styles = {}     # 工作表名 -> {(row,col): dict}
        self._formulas = {}   # 工作表名 -> {(row,col): 公式字符串}
        self._merged = {}     # 工作表名 -> [(r1,c1,r2,c2)]
        self._num_formats = {}  # 工作表名 -> {(row,col): number_format}
        self._format_brush_src = None  # 格式刷模板：None 表示未进入格式刷模式

        self.sheet_tabs.currentChanged.connect(self._on_sheet_changed)

    # ---------- 接口实现 ----------
    def doc_type(self):
        return "excel"

    def new_document(self):
        if not XL_OK:
            raise RuntimeError("未检测到 openpyxl 库，请先运行：pip install openpyxl，然后重启软件")
        self._clear_all()
        wb = Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        self.workbook = wb
        self._add_sheet_view(ws)
        self.file_path = None
        self.set_modified(False)

    def load(self, path):
        if not XL_OK:
            raise RuntimeError("未检测到 openpyxl 库，请先运行：pip install openpyxl，然后重启软件")
        self._clear_all()
        lower = path.lower()
        if lower.endswith(".xls"):
            if not HAS_XLRD:
                raise RuntimeError("读取 .xls 需要 xlrd 库，请执行 pip install xlrd")
            wb = self._load_xls_to_workbook(path)
        else:
            wb = load_workbook(path, data_only=False)
        self.workbook = wb
        for ws in wb.worksheets:
            self._add_sheet_view(ws)
        self.file_path = path
        self.set_modified(False)

    def _load_xls_to_workbook(self, path):
        book = xlrd.open_workbook(path)
        wb = Workbook()
        wb.remove(wb.active)
        for sh in book.sheets():
            ws = wb.create_sheet(title=sh.name[:31] or "Sheet")
            for r in range(sh.nrows):
                for c in range(sh.ncols):
                    val = sh.cell_value(r, c)
                    if val is not None and val != "":
                        ws.cell(row=r + 1, column=c + 1, value=val)
        return wb

    def save(self):
        if not self.file_path:
            return False
        target = self.file_path
        # .xls 统一另存为 .xlsx
        if target.lower().endswith(".xls"):
            target = os.path.splitext(target)[0] + ".xlsx"
        self._dump_workbook(target)
        self.file_path = target
        self.set_modified(False)
        return True

    def save_as(self, path):
        if not path.lower().endswith((".xlsx", ".xlsm", ".xls")):
            path += ".xlsx"
        self._dump_workbook(path)
        self.file_path = path
        self.set_modified(False)
        return True

    # ---------- 内部数据结构 ----------
    def _clear_all(self):
        self.sheet_tabs.blockSignals(True)
        self.sheet_tabs.clear()
        self.sheet_tabs.blockSignals(False)
        self._sheets.clear()
        self._styles.clear()
        self._formulas.clear()
        self._merged.clear()
        self._num_formats.clear()

    def _add_sheet_view(self, ws):
        """根据 openpyxl worksheet 创建一个 QTableWidget 展示页。

        新建空表（1×1）时也统一显示 DEFAULT_GRID_ROWS × DEFAULT_GRID_COLS
        的完整网格（列头 A~Z、行头 1~30），外观与 Excel/WPS 一致。
        """
        rows = max(ws.max_row, DEFAULT_GRID_ROWS)
        cols = max(ws.max_column, DEFAULT_GRID_COLS)
        view = QTableWidget(rows, cols)
        view.setSelectionMode(QTableWidget.SelectionMode.ExtendedSelection)
        view.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectItems)
        view.horizontalHeader().setDefaultSectionSize(90)
        view.verticalHeader().setDefaultSectionSize(24)
        # 列头显示 A~Z 字母（超 26 列自动 AA…）；行头保持默认数字 1~30
        self._refresh_header_labels(view)

        # 自定义绘制委托：支持旋转文本与缩进
        view.setItemDelegate(_ExcelItemDelegate(view, self))
        # 自定义中文右键菜单（覆盖 QTableWidget 默认英文菜单）
        view.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        view.customContextMenuRequested.connect(
            lambda pos, v=view: self._show_cell_menu(v, pos))
        # 滚轮滚动到底自动追加行列
        view.verticalScrollBar().valueChanged.connect(self._on_vscroll_changed)
        view.horizontalScrollBar().valueChanged.connect(self._on_hscroll_changed)

        self._sheets[ws.title] = view
        styles = {}
        formulas = {}
        merged = []

        # 合并区域
        for rng in ws.merged_cells.ranges:
            r1, c1, r2, c2 = rng.min_row, rng.min_col, rng.max_row, rng.max_col
            merged.append((r1, c1, r2, c2))
            view.setSpan(r1 - 1, c1 - 1, r2 - r1 + 1, c2 - c1 + 1)

        # 单元格内容与样式
        for row in ws.iter_rows():
            for cell in row:
                r = cell.row
                c = cell.column
                val = cell.value
                text = "" if val is None else str(val)
                item = QTableWidgetItem(text)
                item.setData(Qt.ItemDataRole.UserRole, (r, c))
                view.setItem(r - 1, c - 1, item)

                if isinstance(val, str) and val.startswith("="):
                    formulas[(r, c)] = val

                st = {}
                f = cell.font
                if f is not None:
                    st["font_name"] = f.name or ""
                    st["font_size"] = f.size or 11
                    st["bold"] = bool(f.bold)
                    st["italic"] = bool(f.italic)
                    st["underline"] = bool(f.underline)
                    if f.color and f.color.rgb:
                        qc = _qcolor_from_openpyxl(f.color.rgb)
                        if qc:
                            st["color"] = qc.name()
                fill = cell.fill
                if fill is not None and fill.fgColor and fill.patternType:
                    qc = _qcolor_from_openpyxl(fill.fgColor.rgb)
                    if qc and qc.name() != "#000000":
                        st["fill"] = qc.name()
                al = cell.alignment
                if al is not None:
                    st["halign"] = al.horizontal or ""
                    st["valign"] = al.vertical or ""
                    st["wrap"] = bool(al.wrap_text)
                    if getattr(al, "text_rotation", 0):
                        st["rotation"] = int(al.text_rotation)
                    if getattr(al, "indent", 0):
                        st["indent"] = int(al.indent)
                bd = cell.border
                if bd is not None and bd.left is not None and bd.left.style:
                    bqc = _qcolor_from_openpyxl(bd.left.color.rgb if bd.left.color else None)
                    if bqc:
                        st["border"] = bqc.name()
                if st:
                    styles[(r, c)] = st
                    self._apply_style(item, st)

        self._styles[ws.title] = styles
        self._formulas[ws.title] = formulas
        self._merged[ws.title] = merged

        view.itemChanged.connect(self._on_item_changed)
        view.currentCellChanged.connect(self._on_current_cell_changed)
        self.sheet_tabs.addTab(view, ws.title)

    def _refresh_header_labels(self, view):
        """刷新列头字母标签：插入 / 删除列后保持 A~Z（超 26 列自动 AA…）。"""
        view.setHorizontalHeaderLabels(
            [_col_name(c) for c in range(1, view.columnCount() + 1)])

    def _apply_style(self, item, st):
        font = QFont(st.get("font_name", "微软雅黑") or "微软雅黑",
                     int(st.get("font_size", 11) or 11))
        font.setBold(bool(st.get("bold")))
        font.setItalic(bool(st.get("italic")))
        font.setUnderline(bool(st.get("underline")))
        font.setStrikeOut(bool(st.get("strike")))
        item.setFont(font)
        color = st.get("color")
        if color:
            item.setForeground(QBrush(QColor(color)))
        fill = st.get("fill")
        if fill:
            item.setBackground(QBrush(QColor(fill)))
        halign = st.get("halign", "")
        valign = st.get("valign", "")
        flag = Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter
        if halign == "center":
            flag = Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter
        elif halign == "right":
            flag = Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
        elif halign == "justify":
            flag = Qt.AlignmentFlag.AlignJustify | Qt.AlignmentFlag.AlignVCenter
        if valign == "top":
            flag = (flag & ~Qt.AlignmentFlag.AlignVCenter) | Qt.AlignmentFlag.AlignTop
        elif valign == "bottom":
            flag = (flag & ~Qt.AlignmentFlag.AlignVCenter) | Qt.AlignmentFlag.AlignBottom
        item.setTextAlignment(flag)

    # ---------- 保存 ----------
    def _dump_workbook(self, path):
        if not XL_OK:
            raise RuntimeError("未检测到 openpyxl 库，请先运行：pip install openpyxl，然后重启软件")
        wb = Workbook()
        wb.remove(wb.active)
        for idx in range(self.sheet_tabs.count()):
            name = self.sheet_tabs.tabText(idx)
            view = self._sheets.get(name)
            if view is None:
                continue
            ws = wb.create_sheet(title=name[:31] or "Sheet")
            styles = self._styles.get(name, {})
            formulas = self._formulas.get(name, {})
            for r in range(view.rowCount()):
                for c in range(view.columnCount()):
                    item = view.item(r, c)
                    if item is None or not item.text():
                        continue
                    key = (r + 1, c + 1)
                    if key in formulas:
                        value = formulas[key]
                    else:
                        value = _to_number(item.text())
                    cell = ws.cell(row=r + 1, column=c + 1, value=value)
                    st = styles.get(key)
                    if st:
                        xf = XlFont(
                            name=st.get("font_name") or None,
                            size=st.get("font_size"),
                            bold=bool(st.get("bold")),
                            italic=bool(st.get("italic")),
                            underline="single" if st.get("underline") else None,
                            strike=bool(st.get("strike")),
                        )
                        if st.get("color"):
                            try:
                                xf.color = _openpyxl_color(QColor(st["color"]))
                            except Exception:
                                pass
                        cell.font = xf
                        if st.get("fill"):
                            try:
                                cell.fill = PatternFill(
                                    fill_type="solid",
                                    fgColor=_openpyxl_color(QColor(st["fill"])),
                                )
                            except Exception:
                                pass
                        cell.alignment = XlAlignment(
                            horizontal=st.get("halign") or None,
                            vertical=st.get("valign") or None,
                            wrap_text=bool(st.get("wrap")),
                            text_rotation=int(st.get("rotation") or 0),
                            indent=int(st.get("indent") or 0),
                        )
                        if st.get("border"):
                            try:
                                side = XlSide(style="thin",
                                              color=_openpyxl_color(QColor(st["border"])))
                                cell.border = XlBorder(left=side, right=side,
                                                       top=side, bottom=side)
                            except Exception:
                                pass
                    fmt = self._num_formats.get(name, {}).get(key)
                    if fmt:
                        cell.number_format = fmt
            # 写入隐藏行 / 隐藏列
            for r in range(view.rowCount()):
                if view.isRowHidden(r):
                    ws.row_dimensions[r + 1].hidden = True
            for c in range(view.columnCount()):
                if view.isColumnHidden(c):
                    ws.column_dimensions[_col_name(c + 1)].hidden = True
            for (r1, c1, r2, c2) in self._merged.get(name, []):
                ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
        wb.save(path)

    # ---------- 单元格操作 ----------
    def _current_view(self):
        return self.sheet_tabs.currentWidget()

    def _current_sheet_name(self):
        return self.sheet_tabs.tabText(self.sheet_tabs.currentIndex())

    def _selected_keys(self, view=None):
        view = view or self._current_view()
        if view is None:
            return []
        keys = []
        for item in view.selectedItems():
            keys.append((item.row() + 1, item.column() + 1))
        return keys

    def _record_style(self, sheet, key, updater):
        styles = self._styles.setdefault(sheet, {})
        st = dict(styles.get(key, {}))
        updater(st)
        styles[key] = st
        return st

    def _on_item_changed(self, item):
        # 用户手动编辑内容时，公式记录按新值更新
        sheet = self._current_sheet_name()
        if not sheet:
            return
        key = (item.row() + 1, item.column() + 1)
        text = item.text()
        if text.startswith("="):
            self._formulas.setdefault(sheet, {})[key] = text
        else:
            self._formulas.get(sheet, {}).pop(key, None)
        self.set_modified(True)

    def _on_current_cell_changed(self, row, col, prev_r, prev_c):
        self.cursor_changed.emit(f"行：{row + 1}  列：{col + 1}")

    def _on_sheet_changed(self, index):
        if index >= 0:
            self.cursor_changed.emit(f"工作表：{self.sheet_tabs.tabText(index)}")

    # ---------- 滚动到底自动追加行列 ----------
    def _on_vscroll_changed(self, value):
        """垂直滚动条滚到底时追加 50 行，保证网格持续可滚动。

        追加行前关闭更新（setUpdatesEnabled False），统一完成后再恢复，
        避免每新增一行触发一次重绘造成滚动卡顿。
        """
        view = self._current_view()
        if view is None:
            return
        sb = view.verticalScrollBar()
        if sb.maximum() > 0 and value >= sb.maximum():
            view.setUpdatesEnabled(False)
            try:
                view.setRowCount(view.rowCount() + APPEND_ROWS_ON_SCROLL)
            finally:
                view.setUpdatesEnabled(True)
            self.set_modified(True)

    def _on_hscroll_changed(self, value):
        """水平滚动条滚到底时追加 10 列，并刷新列头字母。

        追加列前关闭更新，统一完成后恢复并刷新列头，减少逐列重绘。
        """
        view = self._current_view()
        if view is None:
            return
        sb = view.horizontalScrollBar()
        if sb.maximum() > 0 and value >= sb.maximum():
            view.setUpdatesEnabled(False)
            try:
                view.setColumnCount(view.columnCount() + APPEND_COLS_ON_SCROLL)
            finally:
                view.setUpdatesEnabled(True)
            self._refresh_header_labels(view)
            self.set_modified(True)

    # ---------- 单元格右键菜单（全中文） ----------
    def _show_cell_menu(self, view, pos):
        """单元格右键菜单：编辑 + 格式 + 行列操作，全部中文。"""
        menu = QMenu(view)
        has_sel = len(view.selectedItems()) > 0

        act_cut = menu.addAction("剪切")
        act_cut.setEnabled(has_sel)
        act_cut.triggered.connect(self.cut)
        act_copy = menu.addAction("复制")
        act_copy.setEnabled(has_sel)
        act_copy.triggered.connect(self.copy)
        act_paste = menu.addAction("粘贴")
        act_paste.setEnabled(True)
        act_paste.triggered.connect(self.paste)
        act_del = menu.addAction("删除")
        act_del.setEnabled(has_sel)
        act_del.triggered.connect(self.clear_content)
        act_all = menu.addAction("全选")
        act_all.triggered.connect(lambda: view.selectAll())
        menu.addSeparator()

        act_font = menu.addAction("设置字体…")
        act_font.triggered.connect(self._menu_choose_font)
        act_size = menu.addAction("设置字号…")
        act_size.triggered.connect(self._menu_choose_size)
        act_color = menu.addAction("字体颜色…")
        act_color.triggered.connect(self.set_font_color)
        act_fill = menu.addAction("填充颜色…")
        act_fill.triggered.connect(self.set_fill_color)
        menu.addSeparator()

        act_insert_row = menu.addAction("插入行")
        act_insert_row.triggered.connect(self.insert_row)
        act_delete_row = menu.addAction("删除行")
        act_delete_row.triggered.connect(self.delete_row)
        act_insert_col = menu.addAction("插入列")
        act_insert_col.triggered.connect(self.insert_column)
        act_delete_col = menu.addAction("删除列")
        act_delete_col.triggered.connect(self.delete_column)
        menu.addSeparator()

        act_merge = menu.addAction("合并单元格")
        act_merge.triggered.connect(self.merge_selection)
        act_unmerge = menu.addAction("取消合并")
        act_unmerge.triggered.connect(self.unmerge_selection)

        menu.exec(view.mapToGlobal(pos))

    def _menu_choose_font(self):
        """右键菜单选择字体（枚举系统全部字体）。"""
        families = QFontDatabase.families()
        view = self._current_view()
        cur_name = "微软雅黑"
        if view is not None and view.currentItem() is not None:
            key = (view.currentItem().row() + 1, view.currentItem().column() + 1)
            st = self._styles.get(self._current_sheet_name(), {}).get(key, {})
            cur_name = st.get("font_name") or cur_name
        idx = families.index(cur_name) if cur_name in families else 0
        family, ok = QInputDialog.getItem(self, "选择字体", "字体：", families,
                                          idx, False)
        if ok and family:
            self.set_font_family(family)

    def _menu_choose_size(self):
        """右键菜单选择字号（8~72 常用列表，可输入自定义值）。"""
        sizes = [str(i) for i in range(8, 73, 2)] + [str(i) for i in range(6, 8)]
        cur = 11
        view = self._current_view()
        if view is not None and view.currentItem() is not None:
            key = (view.currentItem().row() + 1, view.currentItem().column() + 1)
            st = self._styles.get(self._current_sheet_name(), {}).get(key, {})
            cur = int(st.get("font_size") or 11)
        if str(cur) not in sizes:
            sizes.insert(0, str(cur))
        text, ok = QInputDialog.getItem(self, "选择字号", "字号：", sizes,
                                        sizes.index(str(cur)), True)
        if ok:
            try:
                size = float(text.strip())
            except ValueError:
                QMessageBox.warning(self, "提示", "请输入有效的字号数值。")
                return
            self.set_font_size_value(max(1, min(400, size)))

    # ---------- 行列操作 ----------
    def insert_row(self):
        view = self._current_view()
        if view is None:
            return
        rows = {item.row() for item in view.selectedItems()}
        pos = min(rows) if rows else view.currentRow()
        if pos < 0:
            pos = 0
        view.insertRow(pos)
        # 合并信息整体下移
        self._shift_merged(self._current_sheet_name(), row_at=pos, shift=1)
        self.set_modified(True)

    def delete_row(self):
        view = self._current_view()
        if view is None:
            return
        rows = sorted({item.row() for item in view.selectedItems()}, reverse=True)
        if not rows:
            pos = view.currentRow()
            rows = [pos] if pos >= 0 else []
        for r in rows:
            view.removeRow(r)
        self._drop_merged_after_delete(self._current_sheet_name(), deleted_rows=rows)
        self._shift_styles_after_delete(self._current_sheet_name(), deleted_rows=rows)
        self.set_modified(True)

    def insert_column(self):
        view = self._current_view()
        if view is None:
            return
        cols = {item.column() for item in view.selectedItems()}
        pos = min(cols) if cols else view.currentColumn()
        if pos < 0:
            pos = 0
        view.insertColumn(pos)
        self._refresh_header_labels(view)
        self._shift_merged(self._current_sheet_name(), col_at=pos, shift=1)
        self.set_modified(True)

    def delete_column(self):
        view = self._current_view()
        if view is None:
            return
        cols = sorted({item.column() for item in view.selectedItems()}, reverse=True)
        if not cols:
            pos = view.currentColumn()
            cols = [pos] if pos >= 0 else []
        for c in cols:
            view.removeColumn(c)
        self._refresh_header_labels(view)
        self._drop_merged_after_delete(self._current_sheet_name(), deleted_cols=cols)
        self._shift_styles_after_delete(self._current_sheet_name(), deleted_cols=cols)
        self.set_modified(True)

    def _shift_merged(self, sheet, row_at=None, col_at=None, shift=1):
        """插入行/列后平移合并区域。"""
        merged = self._merged.get(sheet, [])
        new_merged = []
        for (r1, c1, r2, c2) in merged:
            if row_at is not None and r1 >= row_at + 1:
                r1 += shift
                r2 += shift
            if col_at is not None and c1 >= col_at + 1:
                c1 += shift
                c2 += shift
            new_merged.append((r1, c1, r2, c2))
        self._merged[sheet] = new_merged

    def _drop_merged_after_delete(self, sheet, deleted_rows=None, deleted_cols=None):
        """删除行/列：受影响合并区域移除，其余按删除位置平移。"""
        merged = self._merged.get(sheet, [])
        new_merged = []
        for (r1, c1, r2, c2) in merged:
            drop = False
            if deleted_rows:
                for dr in deleted_rows:
                    dd = dr + 1  # 被删除行号（1-based）
                    if dd in range(r1, r2 + 1):
                        drop = True
                    if dd < r1:
                        r1 -= 1
                        r2 -= 1
            if deleted_cols:
                for dc in deleted_cols:
                    dd = dc + 1  # 被删除列号（1-based）
                    if dd in range(c1, c2 + 1):
                        drop = True
                    if dd < c1:
                        c1 -= 1
                        c2 -= 1
            if not drop:
                new_merged.append((r1, c1, r2, c2))
        self._merged[sheet] = new_merged

    def _shift_styles_after_delete(self, sheet, deleted_rows=None, deleted_cols=None):
        """删除行/列后重映射样式与公式记录。"""
        styles = self._styles.get(sheet, {})
        formulas = self._formulas.get(sheet, {})
        new_styles = {}
        new_formulas = {}
        for (r, c), st in styles.items():
            rr, cc = r, c
            if deleted_rows:
                if rr in [d + 1 for d in deleted_rows]:
                    continue
                rr -= sum(1 for d in deleted_rows if d + 1 < rr)
            if deleted_cols:
                if cc in [d + 1 for d in deleted_cols]:
                    continue
                cc -= sum(1 for d in deleted_cols if d + 1 < cc)
            new_styles[(rr, cc)] = st
        for (r, c), val in formulas.items():
            rr, cc = r, c
            if deleted_rows:
                if rr in [d + 1 for d in deleted_rows]:
                    continue
                rr -= sum(1 for d in deleted_rows if d + 1 < rr)
            if deleted_cols:
                if cc in [d + 1 for d in deleted_cols]:
                    continue
                cc -= sum(1 for d in deleted_cols if d + 1 < cc)
            new_formulas[(rr, cc)] = val
        self._styles[sheet] = new_styles
        self._formulas[sheet] = new_formulas

    # ---------- 合并 / 取消合并 ----------
    def merge_selection(self):
        view = self._current_view()
        if view is None or len(view.selectedRanges()) == 0:
            return
        rng = view.selectedRanges()[0]
        top, left = rng.topRow(), rng.leftColumn()
        height, width = rng.rowCount(), rng.columnCount()
        if height < 1 or width < 1:
            return
        view.setSpan(top, left, height, width)
        merged = self._merged.setdefault(self._current_sheet_name(), [])
        # 避免重复记录
        record = (top + 1, left + 1, top + height, left + width)
        if record not in merged:
            merged.append(record)
        self.set_modified(True)

    def unmerge_selection(self):
        view = self._current_view()
        if view is None or len(view.selectedRanges()) == 0:
            return
        rng = view.selectedRanges()[0]
        merged = self._merged.get(self._current_sheet_name(), [])
        kept = []
        for (r1, c1, r2, c2) in merged:
            # 与选中范围相交的区域取消合并
            intersect = not (
                r2 < rng.topRow() + 1 or r1 > rng.bottomRow() + 1 or
                c2 < rng.leftColumn() + 1 or c1 > rng.rightColumn() + 1
            )
            if intersect:
                for row in range(r1, r2 + 1):
                    for col in range(c1, c2 + 1):
                        view.setSpan(row - 1, col - 1, 1, 1)
            else:
                kept.append((r1, c1, r2, c2))
        self._merged[self._current_sheet_name()] = kept
        self.set_modified(True)

    # ---------- 样式操作 ----------
    def _apply_to_selected(self, updater):
        view = self._current_view()
        sheet = self._current_sheet_name()
        if view is None:
            return
        items = [it for it in (view.selectedItems() or [view.currentItem()])
                 if it is not None]
        if not items:
            return
        for item in items:
            key = (item.row() + 1, item.column() + 1)
            st = self._record_style(sheet, key, updater)
            self._apply_style(item, st)
        self.set_modified(True)

    def set_bold(self, on):
        self._apply_to_selected(lambda st: st.update(bold=on))

    def set_italic(self, on):
        self._apply_to_selected(lambda st: st.update(italic=on))

    def set_underline(self, on):
        self._apply_to_selected(lambda st: st.update(underline=on))

    def set_font_color(self):
        view = self._current_view()
        if view is None:
            return
        item = view.currentItem()
        cur = item.foreground().color() if item else QColor("#000000")
        color = QColorDialog.getColor(cur, self, "选择字体颜色")
        if color.isValid():
            self._apply_to_selected(lambda st: st.update(color=color.name()))

    def set_fill_color(self):
        view = self._current_view()
        if view is None:
            return
        item = view.currentItem()
        cur = item.background().color() if item else QColor("#FFFFFF")
        color = QColorDialog.getColor(cur, self, "选择单元格背景色")
        if color.isValid() and color.name() != "#000000":
            self._apply_to_selected(lambda st: st.update(fill=color.name()))

    def set_halign(self, align):
        self._apply_to_selected(lambda st: st.update(halign=align))

    def set_valign(self, align):
        self._apply_to_selected(lambda st: st.update(valign=align))

    def set_align(self, align):
        """主窗口对齐按钮的统一入口（Excel 只处理水平对齐）。"""
        self.set_halign(align)

    # ---------- Ribbon / 右键菜单共用格式方法 ----------
    def set_font_family(self, family):
        """把字体族应用到选中单元格。"""
        self._apply_to_selected(lambda st: st.update(font_name=family))

    def set_font_size_value(self, size):
        """把指定字号应用到选中单元格。"""
        self._apply_to_selected(lambda st: st.update(font_size=float(size)))

    def set_font_color_explicit(self, color):
        """按给定颜色直接设置选中单元格字体颜色（供色板 / 按钮调用）。"""
        if color is None:
            return
        name = color.name() if isinstance(color, QColor) else str(color)
        self._apply_to_selected(lambda st: st.update(color=name))

    def set_fill_color_explicit(self, color):
        """按给定颜色直接设置选中单元格填充颜色（供色板 / 按钮调用）。"""
        if color is None:
            return
        name = color.name() if isinstance(color, QColor) else str(color)
        if name == "#000000":
            return
        self._apply_to_selected(lambda st: st.update(fill=name))

    def set_wrap_text(self, on):
        """自动换行：作用于选中单元格并写入样式。"""
        self._apply_to_selected(lambda st: st.update(wrap=bool(on)))
        view = self._current_view()
        if view is not None:
            view.resizeRowsToContents()
            view.viewport().update()

    def set_rotation(self):
        """旋转方向：弹窗选择文本方向角度并应用到选中单元格。"""
        names = ["水平（0°）", "倾斜（45°）", "竖排（90°）", "逆时针（-90°）"]
        angles = [0, 45, 90, -90]
        name, ok = QInputDialog.getItem(self, "旋转方向", "文本方向：", names, 0, False)
        if ok:
            angle = angles[names.index(name)]
            self._apply_to_selected(lambda st, a=angle: st.update(rotation=a))
            view = self._current_view()
            if view is not None:
                view.viewport().update()

    def increase_indent(self):
        """增加缩进级别（作用于选中单元格）。"""
        self._adjust_indent(1)

    def decrease_indent(self):
        """减少缩进级别（作用于选中单元格）。"""
        self._adjust_indent(-1)

    def _adjust_indent(self, delta):
        def updater(st):
            st["indent"] = max(0, int(st.get("indent") or 0) + delta)
        self._apply_to_selected(updater)
        view = self._current_view()
        if view is not None:
            view.viewport().update()

    def set_border(self):
        """给选中单元格设置细边框（弹窗选择颜色）。"""
        view = self._current_view()
        if view is None:
            return
        color = QColorDialog.getColor(QColor("#000000"), self, "选择边框颜色")
        if color.isValid():
            self._apply_to_selected(lambda st: st.update(border=color.name()))
            view.viewport().update()

    # ---------- 格式刷（剪贴板组） ----------
    def capture_format_brush(self):
        """记录当前选中单元格的完整样式为格式刷模板；无单元格时返回 False。"""
        view = self._current_view()
        if view is None:
            return False
        item = view.currentItem() or (view.selectedItems() or [None])[0]
        if item is None:
            return False
        sheet = self._current_sheet_name()
        key = (item.row() + 1, item.column() + 1)
        self._format_brush_src = dict(self._styles.get(sheet, {}).get(key, {}))
        return True

    def apply_format_brush(self):
        """把记录的格式刷模板应用到当前选中单元格；无模板时返回 False。"""
        if self._format_brush_src is None:
            return False
        template = dict(self._format_brush_src)
        view = self._current_view()
        if view is None:
            return False
        items = [it for it in (view.selectedItems() or [view.currentItem()])
                 if it is not None]
        if not items:
            return False
        sheet = self._current_sheet_name()
        for item in items:
            key = (item.row() + 1, item.column() + 1)
            self._styles.setdefault(sheet, {})[key] = dict(template)
            self._apply_style(item, template)
        self.set_modified(True)
        return True

    def clear_format_brush(self):
        """清除格式刷模板（退出格式刷模式时调用）。"""
        self._format_brush_src = None

    def set_font_size(self, delta):
        view = self._current_view()
        if view is None:
            return
        for item in view.selectedItems() or [view.currentItem()]:
            if item is None:
                continue
            key = (item.row() + 1, item.column() + 1)
            sheet = self._current_sheet_name()
            st = self._record_style(sheet, key, lambda st: None)
            cur = st.get("font_size", 11) or 11
            new_size = max(6, min(72, cur + delta))
            self._record_style(sheet, key, lambda st: st.update(font_size=new_size))
            st2 = self._styles[sheet][key]
            self._apply_style(item, st2)
        self.set_modified(True)

    # ---------- 工作表操作 ----------
    def add_sheet(self):
        default = f"Sheet{self.sheet_tabs.count() + 1}"
        name, ok = QInputDialog.getText(self, "新建工作表", "工作表名称：", text=default)
        if not ok or not name.strip():
            return
        name = name.strip()[:31]
        if name in self._sheets:
            QMessageBox.warning(self, "提示", "已存在同名工作表")
            return
        if self.workbook is None:
            self.workbook = Workbook()
        ws = self.workbook.create_sheet(title=name)
        self._add_sheet_view(ws)
        self.sheet_tabs.setCurrentIndex(self.sheet_tabs.count() - 1)
        self.set_modified(True)

    def delete_sheet(self):
        if self.sheet_tabs.count() <= 1:
            QMessageBox.information(self, "提示", "至少保留一个工作表")
            return
        name = self._current_sheet_name()
        box = QMessageBox(self)
        box.setWindowTitle("删除工作表")
        box.setIcon(QMessageBox.Icon.Question)
        box.setText(f"确定删除工作表“{name}”吗？")
        btn_ok = box.addButton("确定", QMessageBox.ButtonRole.AcceptRole)
        btn_cancel = box.addButton("取消", QMessageBox.ButtonRole.RejectRole)
        box.setDefaultButton(btn_ok)
        box.exec()
        if box.clickedButton() is not btn_ok:
            return
        idx = self.sheet_tabs.currentIndex()
        self.sheet_tabs.removeTab(idx)
        self._sheets.pop(name, None)
        self._styles.pop(name, None)
        self._formulas.pop(name, None)
        self._merged.pop(name, None)
        self.set_modified(True)

    def rename_sheet(self):
        name = self._current_sheet_name()
        new_name, ok = QInputDialog.getText(self, "重命名工作表", "新名称：", text=name)
        if not ok or not new_name.strip():
            return
        new_name = new_name.strip()[:31]
        if new_name in self._sheets and new_name != name:
            QMessageBox.warning(self, "提示", "已存在同名工作表")
            return
        idx = self.sheet_tabs.currentIndex()
        self.sheet_tabs.setTabText(idx, new_name)
        view = self._sheets.pop(name)
        self._sheets[new_name] = view
        for d in (self._styles, self._formulas, self._merged):
            if name in d:
                d[new_name] = d.pop(name)
        self.set_modified(True)

    # ---------- 编辑操作 ----------
    def undo(self):
        view = self._current_view()
        if view is not None:
            view.undo()

    def redo(self):
        view = self._current_view()
        if view is not None:
            view.redo()

    def cut(self):
        view = self._current_view()
        if view is not None:
            view.cut()
            self.set_modified(True)

    def copy(self):
        view = self._current_view()
        if view is not None:
            view.copy()

    def paste(self):
        view = self._current_view()
        if view is not None:
            view.paste()
            self.set_modified(True)

    # ---------- 查找替换 ----------
    def find_next(self, text, case_sensitive=False):
        view = self._current_view()
        if view is None:
            return False
        if case_sensitive:
            text_compare = text
        else:
            text_compare = text.lower()
        start_row, start_col = view.currentRow(), view.currentColumn()
        rows, cols = view.rowCount(), view.columnCount()
        # 从当前单元格之后开始查找，支持循环
        for step in range(rows * cols):
            r = (start_row + (step + start_col) // cols) % rows
            c = (start_col + step) % cols
            if step and r == start_row and c == start_col:
                break
            item = view.item(r, c)
            if item is None:
                continue
            val = item.text() if not case_sensitive else item.text()
            if (val if case_sensitive else val.lower()) == text_compare or \
               text_compare in (val if case_sensitive else val.lower()):
                view.setCurrentCell(r, c)
                view.scrollToItem(item)
                self.cursor_changed.emit(f"行：{r + 1}  列：{c + 1}")
                return True
        return False

    def replace_all(self, find_text, repl_text, case_sensitive=False):
        view = self._current_view()
        if view is None:
            return 0
        count = 0
        for r in range(view.rowCount()):
            for c in range(view.columnCount()):
                item = view.item(r, c)
                if item is None:
                    continue
                val = item.text()
                comp = val if case_sensitive else val.lower()
                find_comp = find_text if case_sensitive else find_text.lower()
                if find_comp in comp:
                    item.setText(val.replace(find_text, repl_text) if case_sensitive
                                 else self._replace_ignore_case(val, find_text, repl_text))
                    count += 1
        if count:
            self.set_modified(True)
        return count

    @staticmethod
    def _replace_ignore_case(text, find_text, repl_text):
        out = []
        i = 0
        low = text.lower()
        f_low = find_text.lower()
        while i < len(text):
            if low.startswith(f_low, i):
                out.append(repl_text)
                i += len(find_text)
            else:
                out.append(text[i])
                i += 1
        return "".join(out)

    # ---------- 状态栏 ----------
    def status_text(self):
        view = self._current_view()
        if view is None:
            return ""
        return f"{self._current_sheet_name()}：{view.rowCount()} 行 × {view.columnCount()} 列"

    # ---------- 字体扩展 ----------
    def set_strikethrough(self, on):
        self._apply_to_selected(lambda st: st.update(strike=on))

    # ---------- 隐藏行 / 隐藏列 ----------
    def hide_row(self):
        view = self._current_view()
        if view is None:
            return
        rows = {item.row() for item in view.selectedItems()} or {view.currentRow()}
        for r in rows:
            if 0 <= r < view.rowCount():
                view.setRowHidden(r, True)
        self.set_modified(True)

    def unhide_row(self):
        view = self._current_view()
        if view is None:
            return
        for r in range(view.rowCount()):
            if view.isRowHidden(r):
                view.setRowHidden(r, False)
        self.set_modified(True)

    def hide_column(self):
        view = self._current_view()
        if view is None:
            return
        cols = {item.column() for item in view.selectedItems()} or {view.currentColumn()}
        for c in cols:
            if 0 <= c < view.columnCount():
                view.setColumnHidden(c, True)
        self.set_modified(True)

    def unhide_column(self):
        view = self._current_view()
        if view is None:
            return
        for c in range(view.columnCount()):
            if view.isColumnHidden(c):
                view.setColumnHidden(c, False)
        self.set_modified(True)

    # ---------- 数字格式 ----------
    #: 常用数字格式（名称 -> openpyxl number_format），对照标准 Excel「开始-数字」组
    NUM_FORMATS = {
        "常规": "General",
        "数值": "0.00",
        "货币": "¥#,##0.00",
        "会计专用": "_ ¥#,##0.00_ ;_ -¥#,##0.00_ ;_ ¥0.00_ ;_ @_",
        "日期": "yyyy/m/d",
        "百分比": "0.00%",
        "分数": "# ?/?",
        "科学计数": "0.00E+00",
        "文本": "@",
    }

    def set_num_format(self):
        view = self._current_view()
        if view is None:
            return
        items = [it for it in (view.selectedItems() or [view.currentItem()])
                 if it is not None]
        if not items:
            QMessageBox.information(self, "提示", "请先选中单元格，再设置数字格式。")
            return
        name, ok = QInputDialog.getItem(
            self, "设置数字格式", "选择格式：",
            list(self.NUM_FORMATS.keys()), 0, False)
        if not ok:
            return
        self.set_num_format_explicit(name)

    def set_num_format_explicit(self, name):
        """Ribbon 数字格式下拉直接按名称应用格式。"""
        view = self._current_view()
        if view is None:
            return
        fmt = self.NUM_FORMATS.get(name)
        if fmt is None:
            return
        items = [it for it in (view.selectedItems() or [view.currentItem()])
                 if it is not None]
        if not items:
            return
        sheet = self._current_sheet_name()
        fmts = self._num_formats.setdefault(sheet, {})
        for item in items:
            fmts[(item.row() + 1, item.column() + 1)] = fmt
        self.set_modified(True)

    def set_decimal_places(self, places):
        """按指定小数位数应用「数值」格式（0~8 位）。"""
        view = self._current_view()
        if view is None:
            return
        fmt = "0" + (("." + "0" * places) if places > 0 else "")
        items = [it for it in (view.selectedItems() or [view.currentItem()])
                 if it is not None]
        if not items:
            return
        sheet = self._current_sheet_name()
        fmts = self._num_formats.setdefault(sheet, {})
        for item in items:
            fmts[(item.row() + 1, item.column() + 1)] = fmt
        self.set_modified(True)

    def increase_decimal(self):
        """增加小数位：弹窗输入位数后应用。"""
        n, ok = QInputDialog.getInt(self, "增加小数位", "小数位数：", 2, 0, 8)
        if ok:
            self.set_decimal_places(n)

    def decrease_decimal(self):
        """减少小数位：弹窗输入位数后应用。"""
        n, ok = QInputDialog.getInt(self, "减少小数位", "小数位数：", 0, 0, 8)
        if ok:
            self.set_decimal_places(n)

    # ---------- 排序与筛选 ----------
    def sort_asc(self):
        view = self._current_view()
        if view is None:
            return
        col = view.currentColumn()
        if col < 0:
            return
        view.sortItems(col, Qt.SortOrder.AscendingOrder)

    def sort_desc(self):
        view = self._current_view()
        if view is None:
            return
        col = view.currentColumn()
        if col < 0:
            return
        view.sortItems(col, Qt.SortOrder.DescendingOrder)

    def filter_by_keyword(self):
        """按当前列关键词筛选：隐藏不匹配的行。"""
        view = self._current_view()
        if view is None:
            return
        col = view.currentColumn()
        if col < 0:
            return
        text, ok = QInputDialog.getText(self, "自动筛选", f"按第 {col + 1} 列筛选，输入关键词：")
        if not ok:
            return
        kw = text.strip()
        for r in range(view.rowCount()):
            item = view.item(r, col)
            val = item.text() if item else ""
            view.setRowHidden(r, kw != "" and kw not in val)
        self.set_modified(True)

    def clear_filter(self):
        view = self._current_view()
        if view is None:
            return
        for r in range(view.rowCount()):
            view.setRowHidden(r, False)
        self.set_modified(True)

    # ---------- 自动求和 / 清除内容 ----------
    def auto_sum(self):
        """对选中列连续数值区域求和，结果写入当前单元格（=SUM(...)）。"""
        view = self._current_view()
        if view is None:
            return
        sel = view.selectedItems()
        if not sel or view.currentItem() is None:
            QMessageBox.information(self, "自动求和", "请先选择数值区域，并把光标放到目标单元格")
            return
        sheet = self._current_sheet_name()
        col = view.currentColumn()
        rng = set()
        for item in sel:
            if item.column() == col:
                rng.add(item.row())
        total = 0.0
        ok_all = True
        for r in rng:
            item = view.item(r, col)
            if item is None:
                continue
            try:
                total += float(item.text())
            except (TypeError, ValueError):
                ok_all = False
        if not rng:
            QMessageBox.information(self, "自动求和", "请选择同一列中的多个数值单元格")
            return
        cur = view.currentItem()
        cur_row, cur_col = cur.row(), cur.column()
        key = (cur_row + 1, cur_col + 1)
        if ok_all:
            formula = f"=SUM({_col_name(col + 1)}{min(rng) + 1}:{_col_name(col + 1)}{max(rng) + 1})"
            cur.setText(formula)
            self._formulas.setdefault(sheet, {})[key] = formula
            self.cursor_changed.emit(f"求和结果：{total:.4f}")
        else:
            cur.setText(str(total))
        self.set_modified(True)

    def clear_content(self):
        """清除选中单元格内容与样式。"""
        view = self._current_view()
        if view is None:
            return
        items = [it for it in (view.selectedItems() or [view.currentItem()])
                 if it is not None]
        if not items:
            return
        sheet = self._current_sheet_name()
        for item in items:
            key = (item.row() + 1, item.column() + 1)
            self._styles.get(sheet, {}).pop(key, None)
            self._formulas.get(sheet, {}).pop(key, None)
            self._num_formats.get(sheet, {}).pop(key, None)
            item.setText("")
        self.set_modified(True)

    # ---------- 填充 ----------
    def fill_dialog(self):
        """Ribbon「填充」按钮：弹窗选择填充方向。"""
        names = ["向下填充", "向右填充"]
        name, ok = QInputDialog.getItem(self, "填充", "选择填充方向：", names, 0, False)
        if ok:
            if name == "向下填充":
                self.fill_down()
            else:
                self.fill_right()

    def _fill_source_items(self, same_axis):
        """获取填充源行/列：返回 (src_index, targets) 或 None。"""
        view = self._current_view()
        if view is None:
            return None
        sel = view.selectedItems()
        if not sel:
            return None
        if same_axis:
            # 按列填充：选中区域同一列，首行为源
            col = view.currentColumn()
            rows = sorted({it.row() for it in sel if it.column() == col})
            if len(rows) < 2:
                return None
            return (col, rows)
        # 按行填充：选中区域同一行，首列为源
        row = view.currentRow()
        cols = sorted({it.column() for it in sel if it.row() == row})
        if len(cols) < 2:
            return None
        return (row, cols)

    def fill_down(self):
        """向下填充：把选中区域首行内容（含样式）复制到下方选中行。"""
        view = self._current_view()
        if view is None:
            return
        src = self._fill_source_items(True)
        if src is None:
            QMessageBox.information(self, "填充", "请选中同一列的至少两个单元格（上方为源）")
            return
        col, rows = src
        src_row = rows[0]
        sheet = self._current_sheet_name()
        src_item = view.item(src_row, col)
        src_st = self._styles.get(sheet, {}).get((src_row + 1, col + 1))
        for r in rows[1:]:
            if view.item(r, col) is None:
                view.setItem(r, col, QTableWidgetItem(""))
            item = view.item(r, col)
            item.setText(src_item.text() if src_item else "")
            if src_st:
                self._styles[sheet][(r + 1, col + 1)] = dict(src_st)
                self._apply_style(item, src_st)
        self.set_modified(True)

    def fill_right(self):
        """向右填充：把选中区域首列内容（含样式）复制到右侧选中列。"""
        view = self._current_view()
        if view is None:
            return
        src = self._fill_source_items(False)
        if src is None:
            QMessageBox.information(self, "填充", "请选中同一行的至少两个单元格（左侧为源）")
            return
        row, cols = src
        src_col = cols[0]
        sheet = self._current_sheet_name()
        src_item = view.item(row, src_col)
        src_st = self._styles.get(sheet, {}).get((row + 1, src_col + 1))
        for c in cols[1:]:
            if view.item(row, c) is None:
                view.setItem(row, c, QTableWidgetItem(""))
            item = view.item(row, c)
            item.setText(src_item.text() if src_item else "")
            if src_st:
                self._styles[sheet][(row + 1, c + 1)] = dict(src_st)
                self._apply_style(item, src_st)
        self.set_modified(True)

    # ---------- 打印预览 ----------
    def print_preview(self):
        """把当前表格渲染到打印预览对话框。"""
        view = self._current_view()
        if view is None:
            return
        from PySide6.QtPrintSupport import QPrintPreviewDialog, QPrinter
        printer = QPrinter(QPrinter.PrinterMode.HighResolution)
        preview = QPrintPreviewDialog(printer, self)

        def paint(p):
            view.render(p)
        preview.paintRequested.connect(paint)
        preview.exec()

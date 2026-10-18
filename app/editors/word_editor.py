# -*- coding: utf-8 -*-
"""Word 文档编辑器：基于 QTextEdit 的富文本编辑，通过 python-docx 读写 .docx。

功能：字体（加粗/斜体/下划线/颜色/字号）、段落对齐、列表、插入图片、插入表格、
查找替换、保存/另存为。
"""
import os
from io import BytesIO

from PySide6.QtCore import Qt, QUrl, QBuffer, QIODevice, Signal
from PySide6.QtGui import (
    QTextCursor, QTextDocument, QTextImageFormat, QTextListFormat,
    QTextCharFormat, QColor, QFont, QFontDatabase, QImage, QTextFormat,
    QTextTable, QKeyEvent,
)
from PySide6.QtWidgets import (
    QVBoxLayout, QTextEdit, QColorDialog, QFileDialog, QInputDialog,
    QMessageBox, QMenu,
)

from .base_editor import BaseEditor

try:
    from docx import Document
    from docx.shared import RGBColor, Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.text.run import Run
    from docx.text.paragraph import Paragraph
    from docx.table import Table
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    DOCX_OK = True
except ImportError:
    DOCX_OK = False


def _align_html(alignment):
    """docx 段落对齐 -> HTML style 片段。"""
    try:
        from docx.enum.text import WD_ALIGN_PARAGRAPH
    except ImportError:
        return ""
    if alignment == WD_ALIGN_PARAGRAPH.CENTER:
        return ' style="text-align:center"'
    if alignment == WD_ALIGN_PARAGRAPH.RIGHT:
        return ' style="text-align:right"'
    if alignment == WD_ALIGN_PARAGRAPH.JUSTIFY:
        return ' style="text-align:justify"'
    return ""


def _run_to_html(run, doc, images, img_counter):
    """把一个 docx run 转为 HTML 片段；遇到图片时提取并注册到 QTextDocument 资源。"""
    if not DOCX_OK:
        return run.text
    parts = []
    # 该 run 内是否包含图片（a:blip）
    blips = run._r.findall(".//" + qn("a:blip"))
    text_segments = []  # 文字段（按 blip 位置切分过于复杂，统一放前面展示）
    # 若存在图片，把 run 文本与图片并列输出
    for blip in blips:
        embed = blip.get(qn("r:embed"))
        if not embed:
            continue
        try:
            part = doc.part.related_parts[embed]
        except KeyError:
            continue
        data = part.blob
        img = QImage.fromData(data)
        if img.isNull():
            continue
        name = f"wpsimg{img_counter[0]}.png"
        img_counter[0] += 1
        images[name] = img
        parts.append(f'<img src="{name}" width="{img.width() // 2}" height="{img.height() // 2}"/>')
    if run.text:
        text = run.text
        style = ""
        if run.bold:
            style += "font-weight:bold;"
        if run.italic:
            style += "font-style:italic;"
        if run.underline:
            style += "text-decoration:underline;"
        if run.font.size:
            style += f"font-size:{int(run.font.size.pt)}pt;"
        if run.font.name:
            style += f"font-family:'{run.font.name}';"
        if run.font.color and run.font.color.rgb is not None:
            rgb = str(run.font.color.rgb)
            if rgb and rgb != "000000":
                style += f"color:#{rgb};"
        if style:
            text = f'<span style="{style}">{text}</span>'
        parts.append(text)
    return "".join(parts)


def _paragraph_to_html(paragraph, doc, images, img_counter):
    """docx 段落 -> HTML <p> 片段。"""
    if not DOCX_OK:
        return "<p></p>"
    style = ""
    if paragraph.style and paragraph.style.name:
        sname = paragraph.style.name
        if sname.startswith("List Bullet"):
            style += "margin-left:24px;list-style-type:disc;"
        elif sname.startswith("List Number"):
            style += "margin-left:24px;list-style-type:decimal;"
    html = "".join(
        _run_to_html(r, doc, images, img_counter) for r in paragraph.runs
    )
    align = _align_html(paragraph.alignment)
    return f'<p{align} style="{style}">{html or "&nbsp;"}</p>'


def _table_to_html(table, doc, images, img_counter):
    """docx 表格 -> HTML <table> 片段。"""
    rows = []
    for row in table.rows:
        cells = []
        for cell in row.cells:
            cell_html = ""
            for p in cell.paragraphs:
                cell_html += _paragraph_to_html(p, doc, images, img_counter)
            cells.append(f"<td>{cell_html}</td>")
        rows.append("<tr>" + "".join(cells) + "</tr>")
    return '<table border="1" cellspacing="0" cellpadding="4">' + "".join(rows) + "</table>"


def _docx_to_html(path, images):
    """把 .docx 转成 HTML 字符串，并把图片存入 images（name -> QImage）。"""
    if not DOCX_OK:
        raise RuntimeError("未检测到 python-docx 库，请先运行：pip install python-docx，然后重启软件")
    doc = Document(path)
    img_counter = [0]
    body = doc.element.body
    parts = []
    for child in body.iterchildren():
        if child.tag == qn("w:p"):
            para = Paragraph(child, doc)
            parts.append(_paragraph_to_html(para, doc, images, img_counter))
        elif child.tag == qn("w:tbl"):
            table = Table(child, doc)
            parts.append(_table_to_html(table, doc, images, img_counter))
    return "".join(parts)


class WordEditor(BaseEditor):
    """基于 QTextEdit 的 Word 富文本编辑器。"""

    #: 光标位置变化信号（用于状态栏）
    cursor_changed = Signal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        self.edit = QTextEdit(self)
        self.edit.setAcceptRichText(True)
        self.edit.document().setDefaultFont(QFont("微软雅黑", 11))
        layout.addWidget(self.edit)

        # 自定义中文右键菜单（覆盖 QTextEdit 默认英文菜单）
        self.edit.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.edit.customContextMenuRequested.connect(self._show_context_menu)
        # 拦截 Delete / Backspace：光标位于图片上时直接删除图片
        self.edit.installEventFilter(self)

        # 页眉 / 页脚 / 页码（保存时写入 docx）
        self._header_text = ""
        self._footer_text = ""
        self._want_page_number = False
        # 格式刷模板（None 表示未进入格式刷模式）
        self._format_brush = None

        self.edit.textChanged.connect(lambda: self.set_modified(True))
        self.edit.cursorPositionChanged.connect(self._on_cursor_moved)

    def _on_cursor_moved(self):
        cursor = self.edit.textCursor()
        line = cursor.blockNumber() + 1
        col = cursor.positionInBlock() + 1
        self.cursor_changed.emit(f"行：{line}  列：{col}")

    # ---------- 接口实现 ----------
    def doc_type(self):
        return "word"

    def new_document(self):
        self.edit.clear()
        self.file_path = None
        self.set_modified(False)

    def load(self, path):
        if not DOCX_OK:
            raise RuntimeError("未检测到 python-docx 库，请先运行：pip install python-docx，然后重启软件")
        images = {}
        html = _docx_to_html(path, images)
        doc = self.edit.document()
        # 先把图片注册为文档资源，再设置 HTML
        for name, img in images.items():
            doc.addResource(QTextDocument.ResourceType.ImageResource, QUrl(name), img)
        doc.setHtml(html)
        self.file_path = path
        self.set_modified(False)

    def save(self):
        if not self.file_path:
            return False
        return self.save_as(self.file_path)

    def save_as(self, path):
        if not path.lower().endswith(".docx"):
            path += ".docx"
        if not DOCX_OK:
            raise RuntimeError("未检测到 python-docx 库，请先运行：pip install python-docx，然后重启软件")
        self._html_to_docx(path)
        self.file_path = path
        self.set_modified(False)
        return True

    # ---------- QTextDocument 结构 -> docx ----------
    def _html_to_docx(self, path):
        docx_doc = Document()
        doc = self.edit.document()
        self._walk_frames(doc.rootFrame(), docx_doc, doc)
        self._apply_header_footer(docx_doc)
        docx_doc.save(path)

    def _apply_header_footer(self, docx_doc):
        """把页眉 / 页脚 / 页码写入 docx（默认应用到第一个节）。"""
        if not (self._header_text or self._footer_text or self._want_page_number):
            return
        section = docx_doc.sections[0]

        if self._header_text:
            section.header.paragraphs[0].text = self._header_text
            try:
                from docx.enum.text import WD_ALIGN_PARAGRAPH
                section.header.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            except Exception:
                pass

        if self._footer_text or self._want_page_number:
            para = section.footer.paragraphs[0]
            if self._footer_text:
                para.text = self._footer_text
            if self._want_page_number:
                # 页脚追加 PAGE 域（自动页码）
                run = para.add_run("  第 ")
                fld = OxmlElement("w:fldSimple")
                fld.set(qn("w:instr"), " PAGE ")
                para._p.append(fld)
                run2 = para.add_run(" 页")
            try:
                from docx.enum.text import WD_ALIGN_PARAGRAPH
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            except Exception:
                pass

    def _walk_frames(self, frame, docx_doc, doc):
        """递归遍历 QTextFrame 树：表格与文本块分别处理。"""
        child = frame.firstChild()
        while child is not None:
            if isinstance(child, QTextTable):
                self._process_table(child, docx_doc, doc)
            else:
                self._process_text_frame(child, docx_doc, doc)
            child = child.nextSibling()

    def _process_text_frame(self, frame, docx_doc, doc):
        start = frame.firstPosition()
        end = frame.lastPosition()
        cursor = QTextCursor(doc)
        cursor.setPosition(start)
        block = cursor.block()
        while block.isValid() and block.position() <= end:
            self._process_block(block, docx_doc, doc)
            block = block.next()

    def _process_block(self, block, docx_doc, doc):
        para = docx_doc.add_paragraph()

        # 段落对齐
        fmt = block.blockFormat()
        alignment = fmt.alignment()
        if alignment & Qt.AlignmentFlag.AlignHCenter:
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif alignment & Qt.AlignmentFlag.AlignRight:
            para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        elif alignment & Qt.AlignmentFlag.AlignJustify:
            para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        else:
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT

        # 列表
        text_list = block.textList()
        if text_list is not None:
            lst_fmt = text_list.format()
            if lst_fmt.style() == QTextListFormat.Style.ListDecimal:
                para.style = "List Number"
            else:
                para.style = "List Bullet"

        # 文本片段与图片
        it = block.begin()
        while not it.atEnd():
            frag = it.fragment()
            if frag.isValid():
                cf = frag.charFormat()
                text = frag.text()
                if text:
                    run = para.add_run(text)
                    if cf.fontWeight() >= QFont.Weight.Bold:
                        run.bold = True
                    if cf.fontItalic():
                        run.italic = True
                    if cf.fontUnderline():
                        run.underline = True
                    if cf.fontPointSize() > 0:
                        run.font.size = Pt(cf.fontPointSize())
                    color = cf.foreground().color()
                    if color.isValid() and color.name() != "#000000":
                        try:
                            run.font.color.rgb = RGBColor(color.red(), color.green(), color.blue())
                        except ValueError:
                            pass
                    families = cf.fontFamilies()
                    if families:
                        run.font.name = families[0]
                # 图片片段
                if cf.isImageFormat():
                    img_fmt = cf.toImageFormat()
                    name = img_fmt.name()
                    image = doc.resource(QTextDocument.ResourceType.ImageResource, QUrl(name))
                    if isinstance(image, QImage) and not image.isNull():
                        buffer = QBuffer()
                        buffer.open(QIODevice.OpenModeFlag.WriteOnly)
                        image.save(buffer, "PNG")
                        run = para.add_run()
                        run.add_picture(BytesIO(bytes(buffer.data())))
            it += 1

    def _process_table(self, table, docx_doc, doc):
        rows = table.rows()
        cols = table.columns()
        tbl = docx_doc.add_table(rows=rows, cols=cols)
        try:
            tbl.style = "Table Grid"
        except KeyError:
            pass
        for r in range(rows):
            for c in range(cols):
                cell = table.cellAt(r, c)
                texts = []
                cur = cell.firstCursorPosition()
                end = cell.lastCursorPosition()
                block = cur.block()
                while block.isValid() and block.position() <= end.position():
                    if block.text().strip():
                        texts.append(block.text())
                    block = block.next()
                tbl.cell(r, c).text = "\n".join(texts)

    # ---------- 右键菜单（全中文）与图片删除 ----------
    def eventFilter(self, obj, event):
        """拦截 Delete/Backspace：光标位于插入图片上时直接删除该图片。"""
        if obj is self.edit and isinstance(event, QKeyEvent) and event.type() == QKeyEvent.Type.KeyPress:
            key = event.key()
            if key in (Qt.Key.Key_Delete, Qt.Key.Key_Backspace):
                if self._image_at_cursor():
                    if self._remove_image_at_cursor():
                        event.accept()
                        return True
        return super().eventFilter(obj, event)

    def _image_at_cursor(self):
        """判断当前光标位置或紧邻字符是否为插入图片。"""
        cursor = self.edit.textCursor()
        if cursor.hasSelection():
            return "\ufffc" in cursor.selectedText()
        if cursor.charFormat().isImageFormat():
            return True
        # 光标紧贴在图片前方时，检查其后一个字符
        doc = self.edit.document()
        if cursor.position() < doc.characterCount() - 1:
            c2 = QTextCursor(doc)
            c2.setPosition(cursor.position())
            c2.movePosition(QTextCursor.MoveOperation.Right,
                            QTextCursor.MoveMode.KeepAnchor)
            if c2.selectedText() == "\ufffc":
                return True
        # 光标紧贴在图片后方时，检查其前一个字符
        if cursor.position() > 0:
            c3 = QTextCursor(doc)
            c3.setPosition(cursor.position() - 1)
            c3.movePosition(QTextCursor.MoveOperation.Right,
                            QTextCursor.MoveMode.KeepAnchor)
            if c3.selectedText() == "\ufffc":
                return True
        return False

    def _remove_image_at_cursor(self):
        """选中并删除光标处的插入图片；成功返回 True。"""
        cursor = self.edit.textCursor()
        if cursor.hasSelection() and "\ufffc" in cursor.selectedText():
            cursor.removeSelectedText()
            self.set_modified(True)
            return True
        doc = self.edit.document()
        # 当前字符是图片
        if cursor.charFormat().isImageFormat():
            start = cursor.position()
            if start > 0:
                sel = QTextCursor(doc)
                sel.setPosition(start - 1)
                sel.movePosition(QTextCursor.MoveOperation.Right,
                                QTextCursor.MoveMode.KeepAnchor)
                if sel.selectedText() == "\ufffc":
                    sel.removeSelectedText()
                    self.set_modified(True)
                    return True
        # 光标后一个字符是图片
        if cursor.position() < doc.characterCount() - 1:
            sel2 = QTextCursor(doc)
            sel2.setPosition(cursor.position())
            sel2.movePosition(QTextCursor.MoveOperation.Right,
                              QTextCursor.MoveMode.KeepAnchor)
            if sel2.selectedText() == "\ufffc":
                sel2.removeSelectedText()
                self.set_modified(True)
                return True
        return False

    def _show_context_menu(self, pos):
        """弹出全中文右键菜单：编辑 + 图片删除 + 选区格式。"""
        menu = QMenu(self.edit)
        cursor = self.edit.textCursor()
        has_sel = cursor.hasSelection()
        on_image = self._image_at_cursor()

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
        act_copy.setEnabled(has_sel or on_image)
        act_copy.triggered.connect(self.copy)
        act_paste = menu.addAction("粘贴")
        act_paste.setEnabled(self.edit.canPaste())
        act_paste.triggered.connect(self.paste)
        act_del = menu.addAction("删除")
        act_del.setEnabled(has_sel)
        act_del.triggered.connect(lambda: self.edit.textCursor().removeSelectedText())
        act_sel_all = menu.addAction("全选")
        act_sel_all.triggered.connect(
            lambda: self.edit.selectAll())
        menu.addSeparator()

        if on_image:
            act_del_img = menu.addAction("删除图片")
            act_del_img.triggered.connect(self._remove_image_at_cursor)
            menu.addSeparator()

        # 选区格式：字体 / 字号 / 颜色
        act_font = menu.addAction("字体…")
        act_font.triggered.connect(self._menu_choose_font)
        act_size = menu.addAction("字号…")
        act_size.triggered.connect(self._menu_choose_size)
        act_color = menu.addAction("字体颜色…")
        act_color.triggered.connect(self.set_font_color)

        menu.exec(self.edit.mapToGlobal(pos))

    def _menu_choose_font(self):
        """右键菜单选择字体（枚举系统全部字体）。"""
        families = QFontDatabase.families()
        current = self.edit.textCursor().charFormat().fontFamilies()
        cur_name = current[0] if current else "微软雅黑"
        idx = families.index(cur_name) if cur_name in families else 0
        family, ok = QInputDialog.getItem(self, "选择字体", "字体：", families,
                                          idx, False)
        if ok and family:
            self.set_font_family(family)

    def _menu_choose_size(self):
        """右键菜单选择字号（8~72 常用列表，可输入自定义值）。"""
        cursor = self.edit.textCursor()
        cur = cursor.charFormat().fontPointSize() or 11
        sizes = [str(i) for i in range(8, 73, 2)] + [str(i) for i in range(6, 8)]
        if str(int(cur)) not in sizes:
            sizes.insert(0, str(int(cur)))
        text, ok = QInputDialog.getItem(self, "选择字号", "字号：", sizes,
                                        sizes.index(str(int(cur))), True)
        if ok:
            try:
                size = float(text.strip())
            except ValueError:
                QMessageBox.warning(self, "提示", "请输入有效的字号数值。")
                return
            self.set_font_size_value(max(1, min(400, size)))

    def set_font_family(self, family):
        """把字体族应用到当前选区（无选区则应用到光标处后续输入）。"""
        fmt = QTextCharFormat()
        fmt.setFontFamilies([family])
        self._merge_char_format(fmt)

    def set_font_size_value(self, size):
        """把指定字号应用到当前选区（无选区则应用到光标处后续输入）。"""
        fmt = QTextCharFormat()
        fmt.setFontPointSize(float(size))
        self._merge_char_format(fmt)

    # ---------- 格式操作 ----------
    def _merge_char_format(self, fmt):
        cursor = self.edit.textCursor()
        if not cursor.hasSelection():
            cursor.select(QTextCursor.SelectionType.WordUnderCursor)
        cursor.mergeCharFormat(fmt)
        self.edit.mergeCurrentCharFormat(fmt)
        self.set_modified(True)

    def set_bold(self, on):
        fmt = QTextCharFormat()
        fmt.setFontWeight(QFont.Weight.Bold if on else QFont.Weight.Normal)
        self._merge_char_format(fmt)

    def set_italic(self, on):
        fmt = QTextCharFormat()
        fmt.setFontItalic(on)
        self._merge_char_format(fmt)

    def set_underline(self, on):
        fmt = QTextCharFormat()
        fmt.setFontUnderline(on)
        self._merge_char_format(fmt)

    def set_strikethrough(self, on):
        """设置 / 取消删除线。"""
        fmt = QTextCharFormat()
        fmt.setFontStrikeOut(on)
        self._merge_char_format(fmt)

    def set_highlight(self):
        """给选中文字设置高亮背景色（默认浅黄）。"""
        cursor = self.edit.textCursor()
        color = QColorDialog.getColor(
            QColor("#ffff00"), self, "选择高亮颜色")
        if color.isValid():
            fmt = QTextCharFormat()
            fmt.setBackground(color)
            self._merge_char_format(fmt)

    def set_font_color(self):
        cursor = self.edit.textCursor()
        color = QColorDialog.getColor(
            cursor.charFormat().foreground().color(), self, "选择字体颜色")
        if color.isValid():
            fmt = QTextCharFormat()
            fmt.setForeground(color)
            self._merge_char_format(fmt)

    def set_fill_color(self):
        """给选中文字设置背景（填充）颜色，与 Ribbon「填充色」按钮联动。"""
        cursor = self.edit.textCursor()
        color = QColorDialog.getColor(
            cursor.charFormat().background().color(), self, "选择文字背景色")
        if color.isValid():
            fmt = QTextCharFormat()
            fmt.setBackground(color)
            self._merge_char_format(fmt)

    def set_font_size(self, delta):
        cursor = self.edit.textCursor()
        cur_size = cursor.charFormat().fontPointSize() or 11
        new_size = max(6, min(72, cur_size + delta))
        fmt = QTextCharFormat()
        fmt.setFontPointSize(new_size)
        self._merge_char_format(fmt)

    # ---------- 格式刷 ----------
    def capture_format_brush(self):
        """记录当前光标处字符格式为格式刷模板；无可用格式时返回 False。"""
        cursor = self.edit.textCursor()
        self._format_brush = QTextCharFormat(cursor.charFormat())
        return True

    def apply_format_brush(self):
        """把记录的格式刷模板应用到当前选区；无模板时返回 False。"""
        if self._format_brush is None:
            return False
        self._merge_char_format(QTextCharFormat(self._format_brush))
        return True

    def clear_format_brush(self):
        """清除格式刷模板（退出格式刷模式时调用）。"""
        self._format_brush = None

    def set_align(self, align):
        cursor = self.edit.textCursor()
        if align == "center":
            cursor.setBlockFormat(cursor.blockFormat())
            self.edit.setAlignment(Qt.AlignmentFlag.AlignHCenter)
        elif align == "right":
            self.edit.setAlignment(Qt.AlignmentFlag.AlignRight)
        elif align == "justify":
            self.edit.setAlignment(Qt.AlignmentFlag.AlignJustify)
        else:
            self.edit.setAlignment(Qt.AlignmentFlag.AlignLeft)
        self.set_modified(True)

    def toggle_list(self, kind="bullet"):
        cursor = self.edit.textCursor()
        lst = cursor.currentList()
        if lst is not None:
            cursor.currentList().remove(cursor.block())
            return
        fmt = QTextListFormat()
        if kind == "number":
            fmt.setStyle(QTextListFormat.Style.ListDecimal)
        else:
            fmt.setStyle(QTextListFormat.Style.ListDisc)
        cursor.createList(fmt)
        self.set_modified(True)

    def insert_image(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "选择图片", "", "图片文件 (*.png *.jpg *.jpeg *.bmp *.gif *.webp)",
            options=QFileDialog.Option.DontUseNativeDialog)
        if not path:
            return
        image = QImage(path)
        if image.isNull():
            return
        doc = self.edit.document()
        name = f"wpsimg_user_{os.path.basename(path)}"
        doc.addResource(QTextDocument.ResourceType.ImageResource, QUrl(name), image)
        cursor = self.edit.textCursor()
        img_fmt = QTextImageFormat()
        img_fmt.setName(name)
        img_fmt.setWidth(image.width() // 2)
        img_fmt.setHeight(image.height() // 2)
        cursor.insertImage(img_fmt)
        self.set_modified(True)

    def insert_table(self, rows=3, cols=3):
        cursor = self.edit.textCursor()
        table = cursor.insertTable(rows, cols)
        for r in range(rows):
            for c in range(cols):
                cell_cursor = table.cellAt(r, c).firstCursorPosition()
                cell_cursor.insertText(f"第{r + 1}行 第{c + 1}列")
        self.set_modified(True)

    # ---------- 页眉页脚 / 页码 / 字数 / 打印 ----------
    def insert_header(self):
        """设置文档页眉文字（保存为 docx 时写入页眉）。"""
        text, ok = QInputDialog.getText(self, "插入页眉", "页眉文字：",
                                        text=self._header_text)
        if ok:
            self._header_text = text.strip()
            self.set_modified(True)

    def insert_footer(self):
        """设置文档页脚文字（保存为 docx 时写入页脚）。"""
        text, ok = QInputDialog.getText(self, "插入页脚", "页脚文字：",
                                        text=self._footer_text)
        if ok:
            self._footer_text = text.strip()
            self.set_modified(True)

    def insert_page_number(self):
        """启用页脚页码（保存为 docx 时在页脚追加 PAGE 域）。"""
        self._want_page_number = True
        self.set_modified(True)

    def word_count(self):
        """弹窗显示字数统计。"""
        plain = self.edit.toPlainText()
        chars = len(plain)
        chars_no_space = len(plain.replace(" ", "").replace("\n", "").replace("\t", ""))
        words = len(plain.split())
        paras = len(plain.splitlines())
        QMessageBox.information(
            self, "字数统计",
            f"字符数（计空格）：{chars}\n"
            f"字符数（不计空格）：{chars_no_space}\n"
            f"字数（按空白切分）：{words}\n"
            f"段落数：{paras}")

    def print_preview(self):
        """打印预览（调用系统打印对话框）。"""
        from PySide6.QtPrintSupport import QPrintDialog, QPrinter
        printer = QPrinter(QPrinter.PrinterMode.HighResolution)
        dialog = QPrintDialog(printer, self)
        if dialog.exec() == QPrintDialog.DialogCode.Accepted:
            self.edit.print_(printer)

    # ---------- 编辑操作 ----------
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
        found = doc.find(text, cursor, self._find_flags(case_sensitive))
        if found.isNull():
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
    def status_text(self):
        chars = self.edit.document().characterCount()
        words = len(self.edit.toPlainText().split())
        return f"字数：{chars}  词数：{words}"

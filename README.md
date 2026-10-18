---
AIGC:
    Label: "1"
    ContentProducer: 001191440300708461136T1XGW3
    ProduceID: 3b8ddf290b7a1c33bf8fdf1e3c983246_74bc6eb0ca2b11f18b7f525400393706
    ReservedCode1: VP+FxYZKixHFR9jYqfXQj4VCFAxi6B8ltoBVrwmyNFKEX+uJNHLuggPBpJyblB1gQRJnzoAniCfyTtymA6okz9r32VtA7lGYFPYXpqiIQL9673BHF9gyYOwpUKk3gKy27LFzbUZ+DFaVluSesDpaycJyorFqnrTSme9VKsHelkkbFFxWHqY0hEttTZs=
    ContentPropagator: 001191440300708461136T1XGW3
    PropagateID: 3b8ddf290b7a1c33bf8fdf1e3c983246_74bc6eb0ca2b11f18b7f525400393706
    ReservedCode2: VP+FxYZKixHFR9jYqfXQj4VCFAxi6B8ltoBVrwmyNFKEX+uJNHLuggPBpJyblB1gQRJnzoAniCfyTtymA6okz9r32VtA7lGYFPYXpqiIQL9673BHF9gyYOwpUKk3gKy27LFzbUZ+DFaVluSesDpaycJyorFqnrTSme9VKsHelkkbFFxWHqY0hEttTZs=
---







# 素笺 (SuJian) — 仿 WPS 界面风格的本地办公软件（Python + PySide6）

一个本地运行的桌面办公软件，界面与功能尽可能贴近 WPS：绿色主色调、精细 Ribbon
工具栏、启动欢迎页、快速访问栏、多文档标签页、深浅色主题与流畅动画。支持打开与
编辑四种常见文档类型，无需联网、数据完全本地保存。

## 界面特色

- **启动欢迎页**：展示新建入口（表格/文档/文本/PDF）与最近打开文件，点击即创建或打开。
- **精细 Ribbon**：开始 / 插入 / 页面 / 引用 / 审阅 / 视图 六个标签页，按文档类型
  动态启用分组；按钮均带图标 + 文字，具备悬停 / 按下 / 禁用态，标签页有绿色指示条。
- **快速访问栏**：打开 / 保存 / 撤销 / 重做 / 深色主题 固定按钮，多文档标签页可
  点击 × 关闭、未保存自动提示。
- **主题与动画**：WPS 绿主题 QSS（菜单、按钮、对话框、状态栏、滚动条统一风格），
  支持深色 / 浅色一键切换；标签切换、Ribbon 页切换、按钮悬停、关闭淡出均使用
  150–250ms 缓动动画，不阻塞 UI。

## 功能特性

| 文档类型 | 打开 | 编辑能力 |
| --- | --- | --- |
| **Excel**（.xlsx / .xlsm / .xls） | ✅ | 单元格编辑、公式（`=` 开头，含自动求和）、行列插入/删除/隐藏/取消隐藏、合并/取消合并、字体/字号/加粗/斜体/下划线/字体颜色、单元格背景色、对齐、数字格式、多工作表增删改名、排序升/降序、关键词筛选/取消筛选、清除内容、查找替换、撤销重做、打印预览、保存/另存为 |
| **Word**（.docx） | ✅ | 富文本编辑（粗体/斜体/下划线/删除线/颜色/高亮/字号）、段落左/中/右/两端对齐、项目符号与编号列表、格式刷、插入图片、插入 3x3 表格、插入页眉/页脚/自动页码、字数统计、查找替换、撤销重做、打印预览、保存/另存为 |
| **TXT**（.txt / .md / .log / .py / .json / .csv 等纯文本） | ✅ | 多行文本编辑、查找替换、多编码探测与切换（UTF-8 / GBK / GB18030 / Big5 / UTF-16 / Latin-1）、状态栏行列号、撤销重做、保存/另存为 |
| **PDF**（.pdf） | ✅（查看 + 标注） | 上一页/下一页、放大/缩小、顺时针旋转、矩形高亮标注、双击添加文字批注、导出带标注的 PDF、另存为 |

通用能力：新建、打开（QFileDialog 多选，按扩展名自动路由到对应编辑器）、保存、
另存为、最近文件列表、多文档标签页（可关闭）、状态栏实时显示页数/字数/行列/编码
信息、未保存修改关闭前提示、查找替换（Ctrl+F）。

## 目录结构

```
SuJian/
├── main.py                     # 启动入口
├── requirements.txt            # 依赖清单
├── README.md                   # 本说明文件
├── SuJian.spec                 # PyInstaller 正式打包配置（单文件 exe）
├── build_assets/
│   └── icon.ico                # 绿色程序图标（多尺寸）
├── scripts/
│   ├── installer.py            # 正式安装器（快捷方式 / 卸载入口 / 文件关联）
│   └── uninstaller.py          # 卸载器（可由系统“应用和功能”触发）
└── app/
    ├── __init__.py
    ├── main_window.py          # 主窗口：快速访问栏 + Ribbon + 欢迎页 + 标签页 + 状态栏 + 深浅主题
    ├── ribbon.py               # 仿 WPS Ribbon 工具栏（开始/插入/页面/引用/审阅/视图，模式联动）
    ├── animation.py            # 统一动画模块（fade_in / fade_out_then / 背景色动画）
    ├── welcome_page.py         # 启动欢迎页（新建入口 + 最近文件）
    ├── editors/
    │   ├── __init__.py
    │   ├── base_editor.py      # 编辑器抽象基类（统一接口 + 修改状态信号）
    │   ├── text_editor.py      # TXT 纯文本编辑器（多编码）
    │   ├── word_editor.py      # Word 富文本编辑器（python-docx 读写）
    │   ├── excel_editor.py     # Excel 表格编辑器（openpyxl 读写）
    │   └── pdf_editor.py       # PDF 查看与标注编辑器（pypdfium2）
    └── dialogs/
        ├── __init__.py
        └── find_replace.py     # 通用查找/替换对话框
```

## 安装依赖

建议使用 Python 3.9+ 与虚拟环境：

```bash
cd SuJian
python -m venv .venv
# Windows:
.venv\Scripts\activate
# macOS / Linux:
source .venv/bin/activate

pip install -r requirements.txt
```

依赖清单：`PySide6`（GUI）、`openpyxl`（Excel 读写）、`python-docx`（Word 读写）、
`xlrd`（老版 .xls 读取）、`pypdfium2`（PDF 页面渲染）、`Pillow`（图像辅助）。

## 运行

```bash
python main.py
```

## 打包与安装

### 1. 打包为正式 EXE

项目根目录提供 `SuJian.spec`（PyInstaller 配置：单文件 exe、无控制台窗口、
内嵌绿色图标、隐藏导入 openpyxl / python-docx / xlrd / pypdfium2 / PIL）。

```bash
pip install pyinstaller
cd SuJian
pyinstaller SuJian.spec --clean --noconfirm
```

产物为 `dist\SuJian.exe`（单文件，可直接双击运行）。

### 1.5 生成图形化安装程序 setup.exe

项目提供 `scripts/setup_gui.py`（PySide6 五步安装向导，绿色主题、全中文）
与 `scripts/setup_gui.spec`（PyInstaller 配置：单文件、无控制台、内嵌
`dist/SuJian.exe` 与 `build_assets/icon.ico`）。生成命令：

```bash
# 前提：已按上一节打包出 dist\SuJian.exe
pyinstaller scripts/setup_gui.spec --clean --noconfirm
```

产物为 `dist\setup.exe`（单文件，含主程序约 62MB+）。分发这个 setup.exe 即可：

- 双击打开 → 欢迎页（下一步）；
- 询问“是否安装”（安装 / 取消，取消即退出）；
- 选择安装路径（默认 `%LOCALAPPDATA%\SuJian`，可浏览、自动校验可写）；
- 进度页自动完成：复制主程序、创建桌面/开始菜单快捷方式、注册
  `.xlsx/.xls/.docx/.txt/.pdf` 文件关联、写入系统卸载入口；
- 完成页默认勾选“立即运行”，点“完成”启动软件。

setup.exe 运行时会把内嵌的 `SuJian.exe` 释放到所选安装目录并执行注册，
不需要目标机器预装 Python。卸载方式同“正式安装”章节。

### 2. 正式安装

安装器 `scripts/installer.py` 负责把打包好的 exe 安装成“正式软件”：

```bash
# 默认安装到 %LOCALAPPDATA%\SuJian，自动完成以下全部事项
python scripts/installer.py

# 可选参数
python scripts/installer.py --source D:\path\SuJian.exe   # 指定 exe
python scripts/installer.py --target D:\Apps\SuJian       # 指定安装目录
```

安装器会自动执行：

- 复制 `SuJian.exe`（及卸载器脚本）到安装目录；
- 创建**桌面快捷方式**与**开始菜单快捷方式**「素笺」；
- 写入系统**卸载入口**：在“设置 → 应用 → 应用和功能”中显示
  “素笺”，可从这里直接卸载；
- 注册**文件关联**：双击 `.xlsx / .xls / .docx / .txt / .pdf` 文件，
  自动用本软件打开（基于 HKCU 当前用户注册，无需管理员权限）。

### 3. 卸载

- 方式一（推荐）：打开“设置 → 应用 → 应用和功能”，找到
  “素笺”点击卸载；
- 方式二（命令行）：

```bash
python scripts/uninstaller.py          # 交互式卸载
python scripts/uninstaller.py --silent # 静默卸载（系统卸载入口调用）
```

卸载器会删除：安装目录、桌面 / 开始菜单快捷方式、系统卸载入口注册表项、
本软件创建的文件关联（仅当扩展名仍指向本软件 ProgID 时才移除，不破坏
用户安装前已有的其他关联）。

### 4. 打包说明

- 单文件模式首次启动比目录模式略慢（需解压到临时目录），属正常现象；
- 若用户机器未安装 Python，请在开发机完成安装后分发 `dist\SuJian.exe`；
- 图标源文件为 `build_assets/icon.ico`（16~256 多尺寸），如需更换图标
  直接替换该文件后重新打包即可。

## 使用提示

1. **新建文档**：欢迎页点击类型卡片，或 Ribbon「开始 → 新建」。
2. **打开文件**：快速访问栏 / Ribbon「开始 → 打开」，支持多选，按扩展名自动路由
   （xlsx/xls/xlsm→Excel、docx→Word、txt 等→TXT、pdf→PDF）；菜单「文件 → 最近文件」
   可快速打开近期文档。
3. **Excel 公式**：单元格输入以 `=` 开头的表达式（如 `=A1+B1`、`=SUM(A1:A5)`），
   保存后写入 xlsx 公式；「自动求和」一键插入。
4. **Excel 合并单元格**：先选中区域，再点「开始 → 单元格 → 合并」。
5. **Word 插入图片/表格/页码**：在「插入」标签页操作；表格默认 3x3，可自行调整。
6. **PDF 标注**：点「页面 → 高亮」后拖拽鼠标框选区域完成高亮；双击页面任意位置可
   添加文字批注；「导出PDF」会把高亮与批注固化到新的 PDF 文件。
7. **TXT 编码**：打开时自动探测编码；如需切换保存编码，用「页面 → 文本 → 编码」。
8. **深色主题**：快速访问栏右侧「深色」开关，或「视图 → 深色主题」菜单。

## 已知限制

- Excel 的 .xls（97-2003）以只读方式加载，保存时统一另存为 .xlsx。
- PDF 为“查看 + 标注”模式：高亮/批注需要“导出PDF”固化，不支持直接改写原 PDF 文本。
- Word 的复杂排版（多级样式、域代码等）在编辑器内以近似方式呈现。
- 项目仅供学习交流，请勿用于商业用途。
*（内容由AI生成，仅供参考）*
*（内容由AI生成，仅供参考）*
*（内容由AI生成，仅供参考）*
*（内容由AI生成，仅供参考）*

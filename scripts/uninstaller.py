# -*- coding: utf-8 -*-
"""SuJian（素笺）卸载器。

用法（在项目根目录下执行）：
    python scripts/uninstaller.py              # 交互式卸载（默认稍作停留便于查看结果）
    python scripts/uninstaller.py --silent     # 静默卸载（供系统“应用和功能”后台调用）

功能：
    1. 删除安装目录（默认 %LOCALAPPDATA%\\SuJian，含 exe 与复制过去的卸载器）；
    2. 删除桌面快捷方式与开始菜单快捷方式（SuJian.lnk）；
    3. 删除系统卸载入口注册表项（HKCU ...\\Uninstall\\SuJian）；
    4. 清理文件关联：仅当扩展名默认值仍指向 SuJian.Document 时才移除，
       不会覆盖用户安装本软件之前已有的其他关联；同时删除本软件创建的 ProgID 键。

说明：卸载只清理本软件创建的内容，不触碰用户的文档与系统设置。
"""
import argparse
import os
import shutil
import sys
import winreg

APP_NAME = "SuJian"
APP_DISPLAY_NAME = "素笺"
PROG_ID = "SuJian.Document"
ASSOC_EXTS = [".xlsx", ".xls", ".docx", ".txt", ".pdf"]
UNINSTALL_KEY = r"Software\Microsoft\Windows\CurrentVersion\Uninstall\SuJian"


def detect_install_dir(target):
    """安装目录：优先 --target；其次若本脚本/程序位于安装目录内则取当前目录；
    否则使用默认 %LOCALAPPDATA%\\SuJian。

    兼容 PyInstaller 打包：onefile 运行时 __file__ 指向 _MEI 临时解压目录，
    必须改用 sys.executable 所在目录（即安装目录内复制出来的 uninstaller.exe）。
    """
    if target:
        return os.path.abspath(target)
    if getattr(sys, "frozen", False):
        here = os.path.dirname(os.path.abspath(sys.executable))
    else:
        here = os.path.dirname(os.path.abspath(__file__))
    default = os.path.join(os.environ.get("LOCALAPPDATA", ""), APP_NAME)
    if default and here.lower() == default.lower():
        return here
    return default


def delete_file_if_exists(path):
    """安全删除单个文件；不存在时忽略。"""
    if path and os.path.isfile(path):
        os.remove(path)
        print("  [已删除]", path)
    elif path and os.path.isdir(path):
        print("  [跳过] 不是文件:", path)


def remove_shortcuts():
    """删除桌面与开始菜单快捷方式。"""
    desktop = os.path.join(os.environ.get("USERPROFILE", ""), "Desktop")
    start_menu = os.path.join(os.environ.get("APPDATA", ""),
                              "Microsoft", "Windows", "Start Menu", "Programs")
    for folder in (desktop, start_menu):
        lnk = os.path.join(folder, APP_NAME + ".lnk")
        if os.path.exists(lnk):
            os.remove(lnk)
            print("  [快捷方式已删除]", lnk)
        else:
            print("  [无快捷方式]", lnk)


def remove_install_dir(install_dir):
    """递归删除安装目录（仅限确认过的目录，避免误删系统路径）。"""
    if not install_dir or not os.path.isdir(install_dir):
        print("  [跳过] 安装目录不存在:", install_dir)
        return
    # 防御性校验：只允许删除明确属于本应用的目录（名称含 SuJian）
    if APP_NAME.lower() not in os.path.basename(install_dir).lower():
        print("[错误] 目标目录不是 SuJian 安装目录，拒绝删除:", install_dir)
        sys.exit(1)
    shutil.rmtree(install_dir, ignore_errors=True)
    if os.path.isdir(install_dir):
        print("  [注意] 安装目录中正在运行的卸载器文件暂时无法删除，"
              "其余内容已清理；可在重启后手动删除:", install_dir)
    print("  [安装目录已删除]", install_dir)


def remove_registry():
    """清理卸载入口、扩展名默认值与本软件 ProgID。"""
    # 1) 删除卸载入口键
    try:
        winreg.DeleteKey(winreg.HKEY_CURRENT_USER, UNINSTALL_KEY)
        print("  [卸载入口已删除] HKCU\\" + UNINSTALL_KEY)
    except FileNotFoundError:
        print("  [无卸载入口] HKCU\\" + UNINSTALL_KEY)

    # 2) 扩展名默认值仍指向本软件 ProgID 时，移除该默认值（不删除键本身，避免
    #    破坏用户在安装前已有的其他关联记录）；同时清理 OpenWithProgids 中
    #    本软件登记的值，让“打开方式”列表不再出现本软件
    for ext in ASSOC_EXTS:
        sub = r"Software\Classes" + ext
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, sub, 0,
                                winreg.KEY_READ | winreg.KEY_WRITE) as key:
                try:
                    val, _ = winreg.QueryValueEx(key, "")
                    if val == PROG_ID:
                        winreg.DeleteValue(key, "")
                        print(f"  [关联已移除] {ext}")
                    else:
                        print(f"  [保留] {ext} 当前关联为: {val}（非本软件，不动）")
                except FileNotFoundError:
                    print(f"  [无默认值] {ext}，跳过")
        except FileNotFoundError:
            print(f"  [无注册键] {ext}，跳过")

        # 清理 OpenWithProgids 中本软件的登记值
        owp = r"Software\Classes" + ext + r"\OpenWithProgids"
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, owp, 0,
                                winreg.KEY_READ | winreg.KEY_WRITE) as key:
                try:
                    winreg.DeleteValue(key, PROG_ID)
                    print(f"  [打开方式已清理] {ext} OpenWithProgids 移除 {PROG_ID}")
                except FileNotFoundError:
                    pass
        except FileNotFoundError:
            pass

    # 3) 删除本软件创建的 ProgID 键
    try:
        winreg.DeleteKey(winreg.HKEY_CURRENT_USER,
                         r"Software\Classes" + PROG_ID + r"\shell\open\command")
        winreg.DeleteKey(winreg.HKEY_CURRENT_USER,
                         r"Software\Classes" + PROG_ID + r"\shell\open")
        winreg.DeleteKey(winreg.HKEY_CURRENT_USER,
                         r"Software\Classes" + PROG_ID + r"\shell")
        winreg.DeleteKey(winreg.HKEY_CURRENT_USER,
                         r"Software\Classes" + PROG_ID + r"\DefaultIcon")
        winreg.DeleteKey(winreg.HKEY_CURRENT_USER,
                         r"Software\Classes" + PROG_ID)
        print("  [ProgID 已删除]", PROG_ID)
    except FileNotFoundError:
        print("  [ProgID 不存在]", PROG_ID)


def main():
    parser = argparse.ArgumentParser(description="SuJian（素笺）卸载器")
    parser.add_argument("--target", help="安装目录（默认 %%LOCALAPPDATA%%/SuJian）")
    parser.add_argument("--silent", action="store_true",
                        help="静默卸载，结束后不暂停（供系统“应用和功能”调用）")
    args = parser.parse_args()

    install_dir = detect_install_dir(args.target)
    print(f"开始卸载 SuJian（素笺）（安装目录: {install_dir}）")

    remove_shortcuts()
    remove_registry()
    remove_install_dir(install_dir)

    print("卸载完成，已移除快捷方式 / 注册表卸载项 / 文件关联 / 安装目录。")
    if not args.silent:
        try:
            input("按回车键关闭窗口...")
        except EOFError:
            pass


if __name__ == "__main__":
    main()

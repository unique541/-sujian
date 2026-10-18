# -*- coding: utf-8 -*-
"""SuJian（素笺）正式安装器。

用法（在项目根目录下执行）：
    python scripts/installer.py                          # 使用 dist/SuJian.exe，安装到 %LOCALAPPDATA%/SuJian
    python scripts/installer.py --source path/to/SuJian.exe
    python scripts/installer.py --target D:/Apps/SuJian

功能：
    1. 复制 exe 到安装目录（默认 %LOCALAPPDATA%/SuJian）；
    2. 复制卸载器 uninstaller.py 到安装目录，供系统“应用和功能”调用；
    3. 创建“桌面快捷方式”与“开始菜单快捷方式”（SuJian）；
    4. 写入系统卸载入口注册表项（DisplayName / UninstallString / DisplayIcon / InstallLocation / Publisher）；
    5. 注册文件关联：双击 .xlsx / .xls / .docx / .txt / .pdf 时用本软件打开
       （HKEY_CURRENT_USER\\Software\\Classes，仅当前用户、无需管理员权限）。

说明：本脚本写入的位置均为“当前用户”级（HKCU），不会影响系统其他用户，
      卸载时也只会清理本软件自己创建的项，不触碰用户既有设置。
"""
import argparse
import os
import shutil
import subprocess
import sys
import winreg

APP_NAME = "SuJian"
APP_DISPLAY_NAME = "素笺"
APP_VERSION = "1.0.0"
APP_PUBLISHER = "SuJian"
PROG_ID = "SuJian.Document"          # 自定义 ProgID，多个扩展名共用
ASSOC_EXTS = [".xlsx", ".xls", ".docx", ".txt", ".pdf"]   # 要关联的扩展名
UNINSTALL_KEY = r"Software\Microsoft\Windows\CurrentVersion\Uninstall\SuJian"


def _pythonw():
    """返回可用于后台运行卸载脚本的 Python 解释器路径（优先 pythonw.exe）。

    注意：此函数只在“以 Python 脚本方式”运行安装器时可靠。
    当安装器被打包为 setup.exe（PyInstaller onefile）时 sys.executable
    指向 setup.exe 本身，其目录下不存在 pythonw.exe——因此卸载入口
    优先使用自包含的 uninstaller.exe（见 write_uninstall_entry）。
    """
    base = os.path.dirname(sys.executable)
    pyw = os.path.join(base, "pythonw.exe")
    return pyw if os.path.exists(pyw) else sys.executable


def find_source_exe(project_root):
    """默认取打包产物 dist/SuJian.exe；不存在时提示先打包。"""
    exe = os.path.join(project_root, "dist", "SuJian.exe")
    if not os.path.exists(exe):
        print("[错误] 未找到打包产物:", exe)
        print("       请先在项目根目录执行： pyinstaller SuJian.spec --clean --noconfirm")
        sys.exit(1)
    return exe


def make_shortcut(lnk_path, target_exe):
    """通过 PowerShell WScript.Shell 创建 .lnk 快捷方式。"""
    ps = (
        "$ws = New-Object -ComObject WScript.Shell;"
        f"$sc = $ws.CreateShortcut('{lnk_path}');"
        f"$sc.TargetPath = '{target_exe}';"
        f"$sc.IconLocation = '{target_exe},0';"
        "$sc.Description = '素笺';"
        "$sc.Save()"
    )
    subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps],
                   check=True)


def reg_set_value(root, subkey, value_name, value):
    """写入注册表字符串值，自动创建中间键。"""
    with winreg.CreateKey(root, subkey) as key:
        winreg.SetValueEx(key, value_name, 0, winreg.REG_SZ, value)


def reg_set_default(root, subkey, value):
    """写入注册表默认值（空名称）。"""
    with winreg.CreateKey(root, subkey) as key:
        winreg.SetValueEx(key, "", 0, winreg.REG_SZ, value)


def register_file_assoc(exe_path):
    """注册扩展名关联到 SuJian.Document。"""
    for ext in ASSOC_EXTS:
        # 1) 扩展名 -> ProgID（默认值），双击即走本 ProgID 的打开命令
        reg_set_default(winreg.HKEY_CURRENT_USER,
                        r"Software\Classes" + ext, PROG_ID)
        # 2) OpenWithProgids：让本软件出现在“打开方式”候选列表，
        #    并在“设置-默认应用”中可被切换为该扩展名的默认程序
        reg_set_value(winreg.HKEY_CURRENT_USER,
                      r"Software\Classes" + ext + r"\OpenWithProgids",
                      PROG_ID, "")
        # 3) ProgID 打开命令：exe 与参数均用引号包裹，防止路径含空格导致双击失败
        reg_set_default(winreg.HKEY_CURRENT_USER,
                        r"Software\Classes" + PROG_ID + r"\shell\open\command",
                        f'"{exe_path}" "%1"')
        # 4) ProgID 显示名称（“打开方式”列表与右键菜单中的程序名）
        reg_set_value(winreg.HKEY_CURRENT_USER,
                      r"Software\Classes" + PROG_ID,
                      "FriendlyAppName", APP_DISPLAY_NAME)
        # 5) 图标：使用 exe 内嵌图标（索引 0）
        reg_set_default(winreg.HKEY_CURRENT_USER,
                        r"Software\Classes" + PROG_ID + r"\DefaultIcon",
                        f'"{exe_path}",0')
        print(f"  [关联] {ext} -> {PROG_ID}")


def _uninstaller_cmd(install_dir):
    """构造系统卸载入口命令。

    优先使用自包含的 uninstaller.exe（不依赖用户机器是否安装 Python，
    避免打包后的 setup.exe 把解释器路径写成自身目录导致“卸载打开的是
    安装程序 / 找不到 pythonw.exe”）；回退到 uninstaller.py 脚本方式。
    """
    uninstaller_exe = os.path.join(install_dir, "uninstaller.exe")
    if os.path.isfile(uninstaller_exe):
        return f'"{uninstaller_exe}"'
    uninstaller_py = os.path.join(install_dir, "uninstaller.py")
    return f'"{_pythonw()}" "{uninstaller_py}"'


def write_uninstall_entry(exe_path, install_dir):
    """写入系统卸载入口（应用和功能可见，可点击卸载）。"""
    uninstaller_cmd = _uninstaller_cmd(install_dir)
    reg_set_value(winreg.HKEY_CURRENT_USER, UNINSTALL_KEY, "DisplayName", APP_DISPLAY_NAME)
    reg_set_value(winreg.HKEY_CURRENT_USER, UNINSTALL_KEY, "DisplayVersion", APP_VERSION)
    reg_set_value(winreg.HKEY_CURRENT_USER, UNINSTALL_KEY, "Publisher", APP_PUBLISHER)
    reg_set_value(winreg.HKEY_CURRENT_USER, UNINSTALL_KEY, "DisplayIcon", f'"{exe_path}",0')
    reg_set_value(winreg.HKEY_CURRENT_USER, UNINSTALL_KEY, "InstallLocation", install_dir)
    reg_set_value(winreg.HKEY_CURRENT_USER, UNINSTALL_KEY,
                  "UninstallString", f"{uninstaller_cmd}")
    reg_set_value(winreg.HKEY_CURRENT_USER, UNINSTALL_KEY,
                  "QuietUninstallString", f"{uninstaller_cmd} --silent")
    print("  [卸载入口] 已写入：设置 → 应用 → 应用和功能 → 素笺")
    print("            命令:", uninstaller_cmd)


def install(args):
    project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    source_exe = args.source or find_source_exe(project_root)
    install_dir = args.target or os.path.join(os.environ.get("LOCALAPPDATA", os.path.expanduser("~")),
                                              APP_NAME)

    print(f"[1/5] 复制程序到安装目录：{install_dir}")
    os.makedirs(install_dir, exist_ok=True)
    target_exe = os.path.join(install_dir, APP_NAME + ".exe")
    shutil.copy2(source_exe, target_exe)
    # 复制卸载组件到安装目录，供“应用和功能”调用：
    # 优先自包含 uninstaller.exe（不依赖用户机器 Python），否则回退 uninstaller.py
    uninstaller_exe_src = os.path.join(project_root, "dist", "uninstaller.exe")
    if os.path.isfile(uninstaller_exe_src):
        shutil.copy2(uninstaller_exe_src,
                     os.path.join(install_dir, "uninstaller.exe"))
        print("      已复制卸载器:", os.path.join(install_dir, "uninstaller.exe"))
    else:
        shutil.copy2(os.path.join(project_root, "scripts", "uninstaller.py"),
                     os.path.join(install_dir, "uninstaller.py"))
        print("      已复制卸载器:",
              os.path.join(install_dir, "uninstaller.py"),
              "（提示：建议先打包 uninstaller.exe 以彻底摆脱对 Python 的依赖）")
    print("      已复制:", target_exe)

    print("[2/5] 创建快捷方式")
    desktop = os.path.join(os.environ["USERPROFILE"], "Desktop")
    start_menu = os.path.join(os.environ["APPDATA"],
                              "Microsoft", "Windows", "Start Menu", "Programs")
    desktop_lnk = os.path.join(desktop, "SuJian.lnk")
    start_lnk = os.path.join(start_menu, "SuJian.lnk")
    if os.path.isdir(desktop):
        make_shortcut(desktop_lnk, target_exe)
        print("  [桌面快捷方式]", desktop_lnk)
    else:
        print("  [跳过] 未找到桌面目录:", desktop)
    if os.path.isdir(start_menu):
        make_shortcut(start_lnk, target_exe)
        print("  [开始菜单快捷方式]", start_lnk)
    else:
        print("  [跳过] 未找到开始菜单目录:", start_menu)

    print("[3/5] 注册文件关联")
    register_file_assoc(target_exe)

    print("[4/5] 写入系统卸载入口")
    write_uninstall_entry(target_exe, install_dir)

    print("[5/5] 安装完成")
    print("  - 双击 .xlsx / .xls / .docx / .txt / .pdf 将使用本软件打开；")
    print("  - 卸载：设置 → 应用 → 应用和功能 → 找到“素笺”点卸载，")
    print("          或运行：python scripts/uninstaller.py")


def main():
    parser = argparse.ArgumentParser(description="SuJian（素笺）安装器")
    parser.add_argument("--source", help="要安装的 exe 路径（默认 dist/SuJian.exe）")
    parser.add_argument("--target", help="安装目录（默认 %%LOCALAPPDATA%%/SuJian）")
    install(parser.parse_args())


if __name__ == "__main__":
    main()

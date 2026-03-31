import os
import sys
import shutil

# 项目使用的所有字体清单：(文件名, 注册表显示名)
FONT_LIST = [
    ("仿宋_GB2312.ttf", "仿宋_GB2312 (TrueType)"),
    ("simsun.ttc",       "宋体 & 新宋体 (TrueType)"),
    ("simhei.ttf",       "黑体 (TrueType)"),
]


def check_and_install_font():
    """检查系统是否安装了项目所需的全部字体，若缺失则自动安装至当前用户 Fonts 目录"""
    if sys.platform not in ("win32", "darwin"):
        return

    if sys.platform == "win32":
        # Windows: 全局与当前用户的 Fonts 目录
        system_font_dir = os.path.join(os.environ.get('WINDIR', 'C:\\Windows'), 'Fonts')
        user_font_dir = os.path.join(os.environ.get('LOCALAPPDATA', ''), 'Microsoft', 'Windows', 'Fonts')
    else:
        # macOS: 系统 Fonts 与用户 Fonts 目录
        system_font_dir = '/Library/Fonts'
        user_font_dir = os.path.expanduser('~/Library/Fonts')

    # 定位资源包内的字体目录
    if getattr(sys, 'frozen', False):
        # 打包环境 (PyInstaller)
        base_path = sys._MEIPASS
    else:
        # 开发环境
        base_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

    font_source_dir = os.path.join(base_path, 'font')

    for font_filename, font_reg_name in FONT_LIST:
        _install_single_font(
            font_filename,
            font_reg_name,
            font_source_dir,
            system_font_dir,
            user_font_dir,
        )


def _install_single_font(font_filename, font_reg_name, font_source_dir, system_font_dir, user_font_dir):
    """安装单个字体文件（如果尚未安装）"""
    sys_font_path = os.path.join(system_font_dir, font_filename)
    usr_font_path = os.path.join(user_font_dir, font_filename)

    # 1. 检查是否已存在
    if os.path.exists(sys_font_path) or os.path.exists(usr_font_path):
        return  # 已安装

    # 2. 找到源文件
    source_font_path = os.path.join(font_source_dir, font_filename)
    if not os.path.exists(source_font_path):
        print(f"警告: 找不到字体源文件 {source_font_path}")
        return

    try:
        # 3. 复制字体到用户的字体目录
        if not os.path.exists(user_font_dir):
            os.makedirs(user_font_dir, exist_ok=True)

        shutil.copy2(source_font_path, usr_font_path)

        # 4. Windows 需要修改注册表，Mac 拷贝过去即可自动生效
        if sys.platform == "win32":
            import winreg
            key_path = r"Software\Microsoft\Windows NT\CurrentVersion\Fonts"
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_SET_VALUE) as key:
                winreg.SetValueEx(key, font_reg_name, 0, winreg.REG_SZ, usr_font_path)

        print(f"成功自动安装缺失字体: {font_reg_name}")
    except Exception as e:
        print(f"自动安装字体失败 ({font_filename}): {e}")

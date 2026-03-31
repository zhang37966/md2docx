import os
import sys
import shutil
import winreg

FONT_FILENAME = "仿宋_GB2312.ttf"
FONT_REG_ZNAME = "仿宋_GB2312 (TrueType)"

def check_and_install_font():
    """检查系统是否安装了仿宋_GB2312，若无则自动安装至当前用户 Fonts 目录"""
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

    sys_font_path = os.path.join(system_font_dir, FONT_FILENAME)
    usr_font_path = os.path.join(user_font_dir, FONT_FILENAME)

    # 1. 检查文件是否存在
    if os.path.exists(sys_font_path) or os.path.exists(usr_font_path):
        return  # 字体已存在

    # 2. 定位资源包内的字体位置
    if getattr(sys, 'frozen', False):
        # 打包环境 (PyInstaller)
        base_path = sys._MEIPASS
    else:
        # 开发环境
        base_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

    source_font_path = os.path.join(base_path, 'font', FONT_FILENAME)

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
                winreg.SetValueEx(key, FONT_REG_ZNAME, 0, winreg.REG_SZ, usr_font_path)
            
        print(f"成功自动安装缺失字体: {FONT_REG_ZNAME}")
    except Exception as e:
        print(f"自动安装字体失败: {e}")

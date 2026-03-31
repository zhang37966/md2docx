import os
import sys
import subprocess
import shutil

def clean_build_dirs():
    """清理旧的打包临时文件和输出文件夹"""
    print("清理旧的打包临时文件...")
    for dir_name in ['build', 'dist']:
        if os.path.exists(dir_name):
            try:
                shutil.rmtree(dir_name)
                print(f"✅ 已删除 {dir_name}/")
            except Exception as e:
                print(f"⚠️ 无法删除 {dir_name}/: {e}")
    
    spec_file = "MD2DOCX.spec"
    if os.path.exists(spec_file):
        os.remove(spec_file)
        print(f"✅ 已删除 {spec_file}")

def build_exe():
    """执行 PyInstaller 打包"""
    print("\n🚀 开始打包 MD转Word工具.exe ...")
    
    # 确保依赖已安装
    try:
        import PyInstaller
    except ImportError:
        print("未检测到 pyinstaller，正在安装...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
        print("✅ pyinstaller 安装完成")

    # 配置打包参数
    # -F/--onefile : 打包为单文件
    # -w/--windowed: 隐藏控制台黑窗口 (GUI 程序必备)
    # -n/--name    : 生成的 exe 名字
    # --clean      : 每次打包前清理
    # 根据操作系统动态选择分隔符 (Windows 是 ; Mac/Linux 是 :)
    separator = ";" if sys.platform == "win32" else ":"

    pyinstaller_cmd = [
        sys.executable, "-m", "PyInstaller",
        "--clean",
        "-F",
        "-w",
        "-n", "MD2DOCX",
        "--add-data", f"font{separator}font",  # 包含字体目录
        "main.py"
    ]

    print("\n运行打包命令：", " ".join(pyinstaller_cmd))
    
    # 运行打包
    process = subprocess.run(pyinstaller_cmd)
    
    if process.returncode == 0:
        if sys.platform == "win32":
            out_path = os.path.abspath(os.path.join("dist", "MD2DOCX.exe"))
        elif sys.platform == "darwin":
            out_path = os.path.abspath(os.path.join("dist", "MD2DOCX.app"))
        else:
            out_path = os.path.abspath(os.path.join("dist", "MD2DOCX"))
            
        print("\n" + "="*50)
        print("🎉 打包成功！")
        print(f"📁 你的独立程序位于： {out_path}")
        print("="*50)
    else:
        print("\n❌ 打包失败，请检查上面的错误输出。")

if __name__ == "__main__":
    # 切换到脚本所在目录
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    
    clean_build_dirs()
    build_exe()

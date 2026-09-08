"""程序入口文件，只负责启动 GUI 主界面。"""

from gui import run_app


if __name__ == "__main__":
    # 直接运行 main.py 时启动应用；被其他模块导入时不自动弹窗。
    run_app()
                                        
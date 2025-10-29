#WinDrawer.py
import json
import os
import sys
from pathlib import Path

import win32api
import win32event
import winerror
from PyQt6.QtWidgets import QApplication
from visual import (CR_Mwin, Anim_AppearMwin, BD_kSC, add_func_menu_button,
                    Strict_Spec_CfgFile, )

def create_resources():#创建必要的资源和配置文件，符合规范
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(__file__)
    data_dir = os.path.join(base_path, "data")
    shortcut_dir = os.path.join(data_dir, "ExeLink")
    if not os.path.exists(shortcut_dir):
        os.makedirs(shortcut_dir, exist_ok=True)
    if getattr(sys, 'frozen', False):
        config_file = Path("config.json")
    else:
        config_file = Path("config.json")
    if not config_file.exists():
        default_config = {
            "#": [
                "LstScanPth=Last Scan Path最后扫描路径",
                "Win=Windows窗口列表",
                "N=Name名称",
                "Geo=Geometry几何属性",
                "Btn=Buttons按钮列表",
                "Pth=Path路径"
            ],
            "LstScanPth": "data/ExeLink",
            "win_order": ["Win_Win0"],
            "win_data": {
                "Win_Win0": {
                    "win_btn_order": [],
                    "win_btn_data": {},
                    "Win_Win0_N": "未分类",
                    "Win_Win0_Geo": [10, 10, 364, 364]
                }
            }
        }
        with open(config_file, 'w', encoding='utf-8') as f:
            json.dump(default_config, f, ensure_ascii=False, indent=4)

def init():
    app = QApplication(sys.argv)
    app.setApplicationName("WinDrawer")
    return app

_app_mutex = None
def check_single_instance(): # 互斥锁
    global _app_mutex
    mutex_name = "Global\\WinDrawer_SingleInstance_Mutex"
    try:
        mutex = win32event.OpenMutex(win32event.SYNCHRONIZE, False, mutex_name)
        if mutex:
            win32api.CloseHandle(mutex)
            return False
    except Exception as e:
        pass
    # 创建互斥锁并保存为全局变量
    try:
        _app_mutex = win32event.CreateMutex(None, True, mutex_name)
        last_error = win32api.GetLastError()
        if last_error == winerror.ERROR_ALREADY_EXISTS:
            return False
        else:
            return True
    except Exception as e:
        print(f"[DEBUG] 创建互斥锁异常: {e}")
        return False

def Mfun():
    if not check_single_instance():
        sys.exit(1)
    if not Strict_Spec_CfgFile():
        create_resources()
    app = init()
    Mwin = CR_Mwin()
    BD_kSC(Mwin)
    Anim_AppearMwin(Mwin)
    add_func_menu_button(Mwin)
    sys.exit(app.exec())

if __name__ == '__main__':
    Mfun()


'''更新计划：
.bat快捷方式处理
与win键同时相应
网格对齐
'''
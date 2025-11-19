#WinDrawer.py
import json
import os
import sys
from pathlib import Path

import win32api
from PyQt6.QtWidgets import QApplication
from visual import (CR_Mwin, BD_kSC, add_tool_window_button,
                    Strict_Spec_CfgFile, check_single_instance, Mwin_ToggleState)
from PyQt6.QtCore import Qt

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
                "State=是否折叠状态"
            ],
            "LstScanPth": "data/ExeLink",
            "win_order": ["Win_Win0"],
            "win_data": {
                "Win_Win0": {
                    "win_btn_order": [],
                    "win_btn_data": {},
                    "Win_Win0_N": "未分类",
                    "Win_Win0_State": True,
                    "Win_Win0_Geo": [10, 10, 364, 364]
                }
            }
        }
        with open(config_file, 'w', encoding='utf-8') as f:
            json.dump(default_config, f, ensure_ascii=False, indent=4)

os.environ["QT_ENABLE_HIGHDPI_SCALING"] = "0"
os.environ["QT_SCALE_FACTOR"] = "1"
def init():
    QApplication.setHighDpiScaleFactorRoundingPolicy(Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    app = QApplication(sys.argv)
    app.setApplicationName("WinDrawer")
    return app

def Mfun():
    if not check_single_instance():
        sys.exit(1)
    if not Strict_Spec_CfgFile():
        create_resources()
    app = init()
    Mwin = CR_Mwin()
    BD_kSC(Mwin)
    Mwin_ToggleState(Mwin)
    sys.exit(app.exec())

if __name__ == '__main__':
    Mfun()
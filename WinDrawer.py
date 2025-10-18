#WinDrawer.py
import sys
from PyQt6.QtWidgets import QApplication
from visual import CR_Mwin, Anim_AppearMwin, BD_kSC, add_func_menu_button, Strict_Spec_CfgFile
from pathlib import Path
import json

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

def Mfun():
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

#visual.py
from PyQt6.QtWidgets import (QMainWindow, QApplication, QSystemTrayIcon, QMenu,
                            QMdiArea, QMdiSubWindow, QWidget, QVBoxLayout, QFrame,
                            QLabel, QSizePolicy, QHBoxLayout, QPushButton, QGridLayout,
                            QScrollArea, QDialog, QLineEdit, QMenu, QMessageBox, QFileIconProvider,
                            QFileDialog)
from PyQt6.QtCore import (QPropertyAnimation, QEasingCurve, QRect, Qt, QEvent,
                          QTimer, QObject, pyqtSignal, QFileInfo, QSize)
from PyQt6.QtGui import QIcon, QCursor
import os, json, shutil, sys, ctypes
from pathlib import Path
from math import ceil
import win32api, win32con, pythoncom, win32com.client, win32timezone
from pynput.keyboard import HotKey, Listener


def Strict_Spec_CfgFile():#返回布尔值，True表示符合规范，False表示不符合
    config_file = Path("config.json")
    if not config_file.exists():
        return False
    try:
        with open(config_file, 'r', encoding='utf-8') as f:
            config = json.load(f)
    except:
        return False
    if "win_order" not in config or "win_data" not in config:
        return False
    # 检查win_order和win_data的一致性
    win_order = config.get("win_order", [])
    win_data = config.get("win_data", {})
    for win_key in win_order:
        if win_key not in win_data:
            return False
    # 检查每个窗口内部的按钮顺序和按钮数据
    for win_key, win_config in win_data.items():
        if "win_btn_order" not in win_config or "win_btn_data" not in win_config:
            return False
        btn_order = win_config.get("win_btn_order", [])
        btn_data = win_config.get("win_btn_data", {})
        for btn_key in btn_order:
            if btn_key not in btn_data:
                return False
    return True  # 所有检查通过返回True

def reorganize_config_numbers(config):#重整配置文件中窗口和按钮的序号，使其按顺序排列
    try:
        print("[INFO] 开始重整配置文件序号...")
        # === 窗口部分的重整 ===
        win_order = config.get("win_order", [])
        win_data = config.get("win_data", {})
        # 1. 记录窗口列表长度
        win_count = len(win_order)
        print(f"[DEBUG] 找到 {win_count} 个窗口需要重整")
        # 2-3. 生成新的窗口顺序列表
        new_win_order = [f"Win_Win{i}" for i in range(win_count)]
        # 4. 创建新的窗口数据字典
        new_win_data = {}
        for i, old_win_key in enumerate(win_order):
            if old_win_key not in win_data:
                continue
            new_win_key = new_win_order[i]
            win_config = win_data[old_win_key].copy()
            # 更新窗口名称键
            old_name_key = f"{old_win_key}_N"
            new_name_key = f"{new_win_key}_N"
            if old_name_key in win_config:
                win_config[new_name_key] = win_config.pop(old_name_key)
            # 更新窗口几何键
            old_geo_key = f"{old_win_key}_Geo"
            new_geo_key = f"{new_win_key}_Geo"
            if old_geo_key in win_config:
                win_config[new_geo_key] = win_config.pop(old_geo_key)
            # === 按钮部分的重整 ===
            btn_order = win_config.get("win_btn_order", [])
            btn_data = win_config.get("win_btn_data", {})
            # 1. 记录按钮列表长度
            btn_count = len(btn_order)
            print(f"[DEBUG] 窗口 {new_win_key} 有 {btn_count} 个按钮需要重整")
            # 2-3. 生成新的按钮顺序列表
            new_btn_order = [f"{new_win_key}_Btn{i}" for i in range(btn_count)]
            # 4. 创建新的按钮数据字典
            new_btn_data = {}
            for j, old_btn_key in enumerate(btn_order):
                if old_btn_key not in btn_data:
                    continue
                new_btn_key = new_btn_order[j]
                btn_config = btn_data[old_btn_key].copy()
                # 更新按钮名称键
                old_btn_name_key = f"{old_btn_key}_N"
                new_btn_name_key = f"{new_btn_key}_N"
                if old_btn_name_key in btn_config:
                    btn_config[new_btn_name_key] = btn_config.pop(old_btn_name_key)
                # 更新按钮路径键
                old_btn_path_key = f"{old_btn_key}_Pth"
                new_btn_path_key = f"{new_btn_key}_Pth"
                if old_btn_path_key in btn_config:
                    btn_config[new_btn_path_key] = btn_config.pop(old_btn_path_key)
                new_btn_data[new_btn_key] = btn_config
            # 更新窗口配置中的按钮数据
            win_config["win_btn_order"] = new_btn_order
            win_config["win_btn_data"] = new_btn_data
            new_win_data[new_win_key] = win_config
        # 更新配置
        config["win_order"] = new_win_order
        config["win_data"] = new_win_data
        print("[INFO] 配置文件序号重整完成")
        return True
    except Exception as e:
        print(f"[ERROR] 重整配置文件序号失败: {e}")
        return False

# 全局变量定义
SCR_WIDTH = None  # 屏幕可用宽度（像素），用于计算窗口居中位置
SCR_HEIGHT = None  # 屏幕可用高度（像素），用于确定窗口尺寸
MWIN_WIDTH = None  # 主窗口宽度（像素），设置为屏幕宽度的约80%
MWIN_HEIGHT = None  # 主窗口高度（像素），设置为屏幕高度的约80%
START_X = None  # 窗口初始X坐标（像素），水平居中
START_Y = None  # 窗口初始Y坐标（像素）
APPEAR_END_Y = 0   # 弹出动画结束Y坐标（像素），窗口完全显示在屏幕顶部
HIDE_END_Y = None  # 隐藏动画结束Y坐标（像素），窗口回到初始部分显示状态
SideLen = 64  # 按钮的边长（像素）
sidelen = 16  # 间隔的距离（像素）
MIN_SIZE  = sidelen*2 + SideLen  # 最小尺寸限制
XYCtrlCTN = (364, 364)  # 窗口默认大小
MAX_TEXT_LENGTH = 8  # 每行最大字符数
WRAP_SYMBOLS = ['：', ':', '-', ' ']  # 触发换行的符号

AUTO_HIDE = True  # 自动收起功能开关
AUTO_CAPS = True  # 自动关闭大写锁定开关

class CtrlCTN(QMdiSubWindow):  #子窗口类
    config_updated_signal = pyqtSignal()
    def __init__(self,name='未命名',parent=None,btn_names=None,btn_paths=None,
                 win_key='Win_Win0',win_geo=None):
        super().__init__(parent)
        if win_geo and len(win_geo) == 4:
            x, y, width, height = win_geo
            self.setGeometry(x, y, width, height)
        else:
            self.setGeometry(10, 10, XYCtrlCTN[0], XYCtrlCTN[1])
        self.setWindowTitle(name)
        self.setWindowFlags(Qt.WindowType.CustomizeWindowHint |
                            Qt.WindowType.SubWindow |
                            Qt.WindowType.WindowTitleHint )
        self.setStyleSheet("""QMdiSubWindow {border: 2px solid gray !important;}""")
        self.config_updated_signal.connect(self.ref_btns)
        self.now_Geoattr = None
        self.default_pos = (9,9)
        self.border_margin = 8
        self._is_resizing = False
        self.btn_names = btn_names or []
        self.btn_paths = btn_paths or []
        self.win_key = win_key
        self.win_geo = win_geo
        self.config_path = "config.json"
        self.parent_mwin = parent
        #创建 QScrollArea 作为内容容器，支持滚动条
        self.scroll_area = QScrollArea(self)
        self.scroll_area.setWidgetResizable(True)  # 允许内容 widget 随滚动区域调整
        self.scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded) #垂直滚动条按需显示
        self.scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)#禁用横向滚动条
        self.content_widget = QWidget()
        self.scroll_area.setWidget(self.content_widget)
        #创建子窗口布局管理器
        self.layout = QGridLayout(self.content_widget)
        self.setWidget(self.scroll_area)
        self.layout.setHorizontalSpacing(sidelen)
        self.layout.setVerticalSpacing(sidelen)
        self.layout.setContentsMargins(sidelen, sidelen, sidelen, sidelen)
        self.setMinimumSize(MIN_SIZE+sidelen, MIN_SIZE+sidelen+12)
        # 创建菜单按钮
        self.menu_btn = QPushButton("≡", self)
        self.menu_btn.setFixedSize(16, 16)
        self.menu_btn.setObjectName("windowMenuBtn")
        self.menu_btn.setStyleSheet("""
            QPushButton#windowMenuBtn {
                background-color: transparent;
                border: none;
                color: #000000;           
                font-weight: bold;        
                font-size: 14px;         
            }
            QPushButton#windowMenuBtn:hover {
                background-color: #e0e0e0;
                border-radius: 2px;     
            }
            QPushButton#windowMenuBtn:pressed {
                background-color: #c0c0c0;  
            }
        """)
        # 创建功能菜单
        self.window_menu = QMenu(self)
        rename_action = self.window_menu.addAction("重命名窗口")  # 增
        rename_action.triggered.connect(self._rename_window)
        delete_action = self.window_menu.addAction("删除窗口")
        delete_action.triggered.connect(self._delete_window)
        self.menu_btn.clicked.connect(self._show_window_menu)
        # 更新几何属性时调整按钮位置
        QTimer.singleShot(100, self._adjust_menu_button_position)
        self.buttons = self._CR_btns(name,self.layout)
        self.respos = None
        self.installEventFilter(self)
        QTimer.singleShot(100, self._upd_Gemotry)
        self.config_save_timer = QTimer()
        self.config_save_timer.setSingleShot(True)
        self.config_save_timer.timeout.connect(self._delayed_save_config)

    def _emit_refresh_signal(self):  # 发送刷新信号给所有子窗口
        try:
            # 先重新加载配置文件确保数据同步
            if os.path.exists(self.config_path):
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
            if self.parent_mwin and hasattr(self.parent_mwin, 'centralWidget'):
                mdi_area = self.parent_mwin.centralWidget()
                if isinstance(mdi_area, QMdiArea):
                    sub_windows = mdi_area.subWindowList()
                    for i, sub_window in enumerate(sub_windows):
                        if hasattr(sub_window, 'config_updated_signal'):
                            print(f"[DEBUG] 向第 {i} 个子窗口发送刷新信号")
                            sub_window.config_updated_signal.emit()
        except Exception as e:
            print(f"[ERROR] 发送刷新信号失败: {e}")

    def _adjust_menu_button_position(self):  # 调整菜单按钮位置到右上角
        try:
            border_width = 2
            btn_x = self.width() - self.menu_btn.width() - border_width - 4  # 额外4px内边距
            btn_y = border_width + 4
            self.menu_btn.move(btn_x, btn_y)
        except Exception as e:
            print(f"调整菜单按钮位置时出错: {e}")

    def _show_window_menu(self):  # 显示窗口菜单
        btn_global_pos = self.menu_btn.mapToGlobal(self.menu_btn.rect().bottomRight())
        self.window_menu.exec(btn_global_pos)

    def _delete_window(self):  # 删除当前窗口，并将按钮配置移动到默认窗口中
        try:
            if self.win_key == "Win_Win0":
                dialog = MessageDialog(
                    parent=self,editable=False,default_text="",modal=True,
                    title="无法删除",message="默认窗口不能删除",
                    width=300,height=150)
                dialog.exec()
                return
            if os.path.exists(self.config_path):
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
            else:
                config = {"LstScanPth": "data/ExeLink", "win_order": [], "win_data": {}}
            win_data = config.get("win_data", {})
            if self.win_key not in win_data:
                return
            current_win_config = win_data[self.win_key]
            current_btns = current_win_config.get("win_btn_data", {})
            if "Win_Win0" in win_data and current_btns:
                win0_config = win_data["Win_Win0"]
                win0_btns = win0_config.get("win_btn_data", {})
                win0_btn_order = win0_config.get("win_btn_order", [])
                # 检查最大按钮索引
                max_index = -1
                for btn_key in win0_btn_order:
                    if btn_key.startswith("Win_Win0_Btn"):
                        try:
                            index = int(btn_key.replace("Win_Win0_Btn", ""))
                            if index > max_index:
                                max_index = index
                        except ValueError:
                            pass
                for btn_key, btn_config in current_btns.items():
                    max_index += 1
                    new_btn_key = f"Win_Win0_Btn{max_index}"
                    new_btn_config = {}
                    for old_key, value in btn_config.items():
                        # 提取原按钮名称和路径的键名部分
                        if old_key.endswith("_N"):
                            new_btn_config[f"{new_btn_key}_N"] = value
                        elif old_key.endswith("_Pth"):
                            new_btn_config[f"{new_btn_key}_Pth"] = value
                        else:
                            new_btn_config[old_key] = value  # 保留其他配置
                    # 检查是否已存在相同配置的按钮
                    btn_exists = False
                    for existing_btn_key, existing_config in win0_btns.items():
                        if (existing_config.get(f"{existing_btn_key}_Pth") ==
                                new_btn_config.get(f"{new_btn_key}_Pth")):
                            btn_exists = True
                            break
                    if not btn_exists:
                        win0_btns[new_btn_key] = new_btn_config
                        win0_btn_order.append(new_btn_key)
                win0_config["win_btn_data"] = win0_btns
                win0_config["win_btn_order"] = win0_btn_order
                win_data["Win_Win0"] = win0_config
            if self.win_key in config.get("win_order", []):
                config["win_order"].remove(self.win_key)
            if self.win_key in win_data:
                del win_data[self.win_key]
            config["win_data"] = win_data
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=4)
            self._emit_refresh_signal()
            self.close()
        except Exception as e:
            print(f"[ERROR] 删除窗口失败: {e}")

    def _delete_button(self, button_index):  # 删除按钮：移除快捷方式、配置文件和界面按钮
        try:
            # 先读取配置文件获取按钮名称
            with open(self.config_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
            # 获取当前按钮配置
            win_data = config.get("win_data", {})
            current_win_config = win_data[self.win_key]
            current_btns = current_win_config.get("win_btn_data", {})
            btn_order = current_win_config.get("win_btn_order", [])
            if button_index >= len(btn_order):
                return

            # 获取按钮名称用于显示
            btn_key = btn_order[button_index]
            btn_config = current_btns[btn_key]
            btn_name = btn_config.get(f"{btn_key}_N", f"按钮{button_index}")

            # 定义删除操作函数
            def perform_deletion():
                try:
                    # 重新读取配置确保数据最新
                    with open(self.config_path, 'r', encoding='utf-8') as f:
                        config = json.load(f)
                    win_data = config.get("win_data", {})
                    current_win_config = win_data[self.win_key]
                    current_btns = current_win_config.get("win_btn_data", {})
                    btn_order = current_win_config.get("win_btn_order", [])

                    # 获取按钮路径并删除对应的快捷方式文件
                    btn_path = btn_config.get(f"{btn_key}_Pth", "")
                    if btn_path:
                        try:
                            path_obj = Path(btn_path)
                            if path_obj.exists():
                                path_obj.unlink()  # 删除文件
                                print(f"[INFO] 已删除快捷方式文件: {btn_path}")
                            else:
                                print(f"[WARNING] 快捷方式文件不存在: {btn_path}")
                        except Exception as file_e:
                            print(f"[ERROR] 删除快捷方式文件失败 {btn_path}: {file_e}")

                    # 从配置中删除按钮
                    if btn_key in current_btns:
                        del current_btns[btn_key]
                    if btn_key in btn_order:
                        btn_order.remove(btn_key)

                    # 更新配置文件
                    current_win_config["win_btn_data"] = current_btns
                    current_win_config["win_btn_order"] = btn_order
                    win_data[self.win_key] = current_win_config
                    config["win_data"] = win_data
                    with open(self.config_path, 'w', encoding='utf-8') as f:
                        json.dump(config, f, ensure_ascii=False, indent=4)
                    self._emit_refresh_signal()
                    print(f"[INFO] 按钮 {btn_name} 已删除")
                except Exception as e:
                    print(f"[ERROR] 执行删除操作失败: {e}")

            # 创建非模态确认对话框
            dialog = MessageDialog(
                parent=self, editable=False, default_text=btn_name, modal=False,
                title="确认删除", message="确定要删除这个按钮吗？",
                width=300, height=150, auto_close=3000
            )

            # 连接对话框关闭信号
            def on_dialog_finished(result_code):
                # 只有明确点击取消按钮才不执行删除
                if result_code == QDialog.DialogCode.Rejected:
                    print("[INFO] 用户取消删除操作")
                    return
                # 其他情况（确定、自动关闭、窗口关闭）都执行删除
                perform_deletion()

            dialog.finished.connect(on_dialog_finished)
            dialog.show()

        except Exception as e:
            print(f"[ERROR] 删除按钮失败: {e}")

    def _upd_Gemotry(self):#初始化更新几何属性
        self._upd_now_Geoattr()
        if self._check_Gemotry() in {'N','F'}:
            self._move_DefPos()
        else:
            self.respos = self.now_Geoattr['rect']

    def _upd_now_Geoattr(self):#更新当前几何属性
        try:
            geo = self.geometry()
            parent = self.parent()
            if parent:#计算各边距离
                parent_geo = parent.geometry()
                left=geo.x();   top=geo.y()
                right=parent_geo.width() - (geo.x() + geo.width())
                bottom=parent_geo.height() - (geo.y() + geo.height())
                self.now_Geoattr = {'rect':geo,'dists':(left,top,right,bottom),}
            else:
                self.now_Geoattr = {'rect':geo,'dists':None}
        except Exception as e:
            print(f"ERROR in _upd_now_Geoattr: {e}")

    def _move_DefPos(self):#移动到默认位置
        try:
            rect = QRect(self.default_pos[0], self.default_pos[1], self.width(), self.height())
            self.setGeometry(QRect(rect))
            self._upd_now_Geoattr()
        except Exception as e:
            print(f"ERROR in _move_DefPos: {e}")

    def _check_Gemotry(self):#检查当前几何属性是否符合要求，不能超出边界并与边界保持一定距离
        try:
            if self.now_Geoattr is None:
                return 'N'
            if self.now_Geoattr['dists'] is None:
                return 'N'  #None
            left, top, right, bottom = self.now_Geoattr['dists']
            if (left >= self.border_margin and top >= self.border_margin and
                right >= self.border_margin and bottom >= self.border_margin):
                return "T"  # True
            else:
                return "F"  # False
        except Exception as e:
            return 'N'

    def _apply_Gemotry(self):#依检查结果决定是否回退窗口属性
        try:
            check_result = self._check_Gemotry()
            if check_result == 'T':
                self._upd_now_Geoattr()
                self.respos = self.now_Geoattr['rect']
                return
            if check_result == 'F':
                if self.now_Geoattr and self.now_Geoattr['dists'] != 'N':
                    self.setGeometry(self.respos)
                    self._upd_now_Geoattr()
                    print(f'{self.respos}')
                else:
                    self._move_DefPos()
            elif check_result == 'N':
                self._move_DefPos()
        except Exception as e:
            print(f"ERROR in _apply_Gemotry: {e}")

    def _CR_btns(self,name,layout):#创建按钮
        buttons = []
        btns_text=self.btn_names
        maxcols = (XYCtrlCTN[0] - 2*sidelen) // (sidelen + SideLen)
        for i,buttonTEXT in enumerate(btns_text):
            button = QPushButton(wrap_button_text(buttonTEXT),self.content_widget)
            button.setFixedSize(64,64)
            if i < len(self.btn_paths):
                path = self.btn_paths[i]
                # 获取文件图标
                icon = get_file_icon(path)
                if not icon.isNull():
                    button.setIcon(icon)
                    button.setIconSize(QSize(48, 48))
                    button.setText("")
                else:
                    button.setText(wrap_button_text(buttonTEXT))
                button.clicked.connect(lambda checked, idx=i: self._open_button_file(idx))
                button.setStyleSheet("""
                    QPushButton {
                        background-color: transparent;
                        border: none;
                        border-radius: 4px;
                        padding: 2px;
                        text-align: center;
                        font-size: 9px;
                    }
                    QPushButton:hover {
                        background-color: #404040;
                        border: 1px solid #606060;
                    }
                    QPushButton:pressed {
                        background-color: #606060;
                    }
                """)
            self._setup_button_context_menu(button, i)
            row = i//maxcols
            col = i % maxcols
            layout.addWidget(button,row,col)
            buttons.append(button)
        return buttons

    def _open_button_file(self, button_index):  # 从配置文件读取文件路径并打开
        try:
            with open(self.config_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
            # 获取当前按钮配置
            win_data = config.get("win_data", {})
            current_win_config = win_data[self.win_key]
            current_btns = current_win_config.get("win_btn_data", {})
            btn_order = current_win_config.get("win_btn_order", [])
            # 获取按钮键名和配置
            btn_key = btn_order[button_index]
            btn_config = current_btns[btn_key]
            btn_path = btn_config.get(f"{btn_key}_Pth", "")
            if not btn_path:
                print(f"[WARNING] 按钮 {btn_key} 没有有效的路径")
                return
            try:
                result = win32api.ShellExecute(0, "open", btn_path, None, None, win32con.SW_SHOWNORMAL)
                if result <= 32:
                    print(f"[WARNING] ShellExecute返回错误代码 {result} for {btn_path}")
                    os.startfile(btn_path)
            except Exception as e:
                print(f"[ERROR] 打开文件失败 {btn_path}: {e}")
            if AUTO_HIDE and self.parent_mwin:
                Anim_HideMwin(self.parent_mwin)
                # 更新全局快捷键监听器的状态
                if hasattr(self.parent_mwin, 'hotkey_listener'):
                    self.parent_mwin.hotkey_listener.TGL = False
        except Exception as e:
            print(f"[ERROR] 打开按钮文件失败: {e}")

    def _setup_button_context_menu(self, button, index):#为按钮设置右键菜单
        button.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        button.customContextMenuRequested.connect(
            lambda pos, btn=button, idx=index: self._show_button_context_menu(btn, idx, pos)
        )
        button.installEventFilter(self)

    def _show_button_context_menu(self, button, index, position):#显示按钮的右键菜单
        # 创建菜单容器
        menu_container = QMenu(self)
        menu_container.setWindowFlags(
            Qt.WindowType.FramelessWindowHint |
            Qt.WindowType.Popup |
            Qt.WindowType.WindowStaysOnTopHint
        )
        menu_container.setWindowOpacity(0.8)
        menu_container.setStyleSheet("""
            QMenu {
                background-color: #000000;
                border: 1px solid #cccccc;
            }
            QMenu::item {
                padding: 4px 8px;
                background-color: transparent;
            }
            QMenu::item:selected {
                background-color: #404040;
            }
        """)
        # 创建菜单按钮
        menu_items = [
            ("打开", lambda: self._open_button_file(index)),
            ("重命名", lambda: self._rename_button(index)),
            ("移动", lambda: self._move_button_to_window(index)),
            ("删除", lambda: self._delete_button(index))
        ]
        for text, func in menu_items:
            action = menu_container.addAction(text)
            action.triggered.connect(func)
        # 显示菜单在按钮右侧
        btn_global_pos = button.mapToGlobal(position)
        menu_container.exec(btn_global_pos)

    def _rename_button(self, button_index):  # 增
        try:
            with open(self.config_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
            win_data = config.get("win_data", {})
            current_win_config = win_data[self.win_key]
            current_btns = current_win_config.get("win_btn_data", {})
            btn_order = current_win_config.get("win_btn_order", [])
            if button_index >= len(btn_order):
                return
            btn_key = btn_order[button_index]
            btn_config = current_btns[btn_key]
            old_name = btn_config.get(f"{btn_key}_N", f"按钮{button_index}")
            dialog = MessageDialog(
                parent=self, editable=True, default_text=old_name, modal=True,
                title="修改按钮名称", message="输入新的按钮名称",
                width=350, height=150, auto_close=0)
            result = dialog.exec()
            # 只有确定按钮被点击时才保存
            if result == QDialog.DialogCode.Accepted:
                new_name = dialog.user_input.strip()
                if new_name and new_name != old_name:
                    btn_config[f"{btn_key}_N"] = new_name
                    current_btns[btn_key] = btn_config
                    current_win_config["win_btn_data"] = current_btns
                    win_data[self.win_key] = current_win_config
                    config["win_data"] = win_data
                    with open(self.config_path, 'w', encoding='utf-8') as f:
                        json.dump(config, f, ensure_ascii=False, indent=4)
                    self._emit_refresh_signal()
                    print(f"[INFO] 按钮 {button_index} 已重命名为: {new_name}")
        except Exception as e:
            print(f"[ERROR] 重命名按钮失败: {e}")

    def _rename_window(self):  # 增
        try:
            old_name = self.windowTitle()
            dialog = MessageDialog(
                parent=self, editable=True, default_text=old_name, modal=True,
                title="修改窗口名称", message="输入新的窗口名称",
                width=350, height=150, auto_close=0)
            result = dialog.exec()
            # 只有确定按钮被点击时才保存
            if result == QDialog.DialogCode.Accepted:
                new_name = dialog.user_input.strip()
                if new_name and new_name != old_name:
                    self.setWindowTitle(new_name)
                    if os.path.exists(self.config_path):
                        with open(self.config_path, 'r', encoding='utf-8') as f:
                            config = json.load(f)
                    else:
                        config = {"LstScanPth": "data/ExeLink", "win_order": [], "win_data": {}}
                    win_data = config.get("win_data", {})
                    if self.win_key in win_data:
                        win_data[self.win_key][f"{self.win_key}_N"] = new_name
                        config["win_data"] = win_data
                        with open(self.config_path, 'w', encoding='utf-8') as f:
                            json.dump(config, f, ensure_ascii=False, indent=4)
        except Exception as e:
            print(f"[ERROR] 重命名窗口失败: {e}")

    def _move_button_to_window(self, button_index):  # 移动按钮到其他窗口
        print(f"移动按钮 {button_index} 到其他窗口")
        try:
            # 读取配置文件
            if not os.path.exists(self.config_path):
                return
            with open(self.config_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
            # 获取当前按钮配置
            win_data = config.get("win_data", {})
            if self.win_key not in win_data:
                return
            current_win_config = win_data[self.win_key]
            current_btns = current_win_config.get("win_btn_data", {})
            btn_order = current_win_config.get("win_btn_order", [])
            if not current_btns or button_index >= len(btn_order):
                return
            # 按顺序获取按钮键
            btn_key = btn_order[button_index]
            if btn_key not in current_btns:
                return
            btn_config = current_btns[btn_key]
            # 获取其他窗口列表（排除当前窗口）
            other_windows = {}
            for win_key, win_config in win_data.items():
                if win_key != self.win_key:
                    win_name = win_config.get(f"{win_key}_N", "未命名")
                    other_windows[win_key] = win_name
            if not other_windows:
                print("没有其他窗口可移动")
                return
            # 创建窗口选择菜单
            menu_container = QWidget(self, Qt.WindowType.Popup)
            menu_container.setWindowFlags(
                Qt.WindowType.FramelessWindowHint |
                Qt.WindowType.Popup |
                Qt.WindowType.WindowStaysOnTopHint
            )
            menu_container.setWindowOpacity(0.9)
            menu_container.setStyleSheet("""
                    QWidget {
                        background-color: #000000;
                        border: 1px solid #cccccc;
                    }
                """)
            layout = QVBoxLayout(menu_container)
            layout.setSpacing(0)
            layout.setContentsMargins(0, 0, 0, 0)
            # 添加窗口选项
            for win_key, win_name in other_windows.items():
                win_btn = QPushButton(win_name, menu_container)
                win_btn.setFixedSize(120, 24)
                win_btn.setStyleSheet("""
                        QPushButton {
                            border: none;
                            background-color: #000000;
                            font-size: 12px;
                            padding: 2px;
                            text-align: left;
                        }
                        QPushButton:hover {
                            background-color: #404040;
                        }
                    """)
                win_btn.clicked.connect(
                    lambda checked, target_win=win_key: self._perform_button_move(
                        config, btn_key, btn_config, target_win, menu_container))
                layout.addWidget(win_btn)
            # 设置菜单大小
            menu_height = len(other_windows) * 24
            menu_container.setFixedSize(120, menu_height)
            # 显示菜单在鼠标位置
            cursor_pos = QCursor.pos()
            menu_container.move(cursor_pos.x(), cursor_pos.y())
            menu_container.show()
            menu_container.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose)
        except Exception as e:
            print(f"[ERROR] 移动按钮失败: {e}")

    def _perform_button_move(self, config, btn_key, btn_config, target_win_key, menu_container):
        # 执行按钮移动操作
        try:
            if menu_container:
                menu_container.close()
                menu_container.deleteLater()
            # 从原窗口删除按钮
            win_data = config.get("win_data", {})
            if self.win_key not in win_data:
                return
            current_win_config = win_data[self.win_key]
            current_btns = current_win_config.get("win_btn_data", {})
            current_btn_order = current_win_config.get("win_btn_order", [])
            if btn_key in current_btns:
                del current_btns[btn_key]
            if btn_key in current_btn_order:
                current_btn_order.remove(btn_key)
            # 添加到目标窗口
            if target_win_key not in win_data:
                return
            target_win_config = win_data[target_win_key]
            target_btns = target_win_config.get("win_btn_data", {})
            target_btn_order = target_win_config.get("win_btn_order", [])
            # 寻找可用按钮键名
            btn_index = 0
            while f"{target_win_key}_Btn{btn_index}" in target_btns:
                btn_index += 1
            new_btn_key = f"{target_win_key}_Btn{btn_index}"
            # 创建新的按钮配置
            btn_name = btn_config.get(f"{btn_key}_N", "")
            if not btn_name:
                for key, value in btn_config.items():
                    if key.endswith("_N") and value:
                        btn_name = value
                        break
            btn_path = btn_config.get(f"{btn_key}_Pth", "")
            if not btn_path:
                for key, value in btn_config.items():
                    if key.endswith("_Pth") and value:
                        btn_path = value
                        break
            new_btn_config = {
                f"{new_btn_key}_N": btn_name,
                f"{new_btn_key}_Pth": btn_path
            }
            # 添加到目标窗口
            target_btns[new_btn_key] = new_btn_config
            target_btn_order.append(new_btn_key)
            # 修改此段：更新配置数据结构
            target_win_config["win_btn_data"] = target_btns
            target_win_config["win_btn_order"] = target_btn_order
            win_data[target_win_key] = target_win_config
            current_win_config["win_btn_data"] = current_btns
            current_win_config["win_btn_order"] = current_btn_order
            win_data[self.win_key] = current_win_config
            config["win_data"] = win_data
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=4)
            print(f"[INFO] 按钮已移动到窗口 {target_win_key}")
            self._emit_refresh_signal()
        except Exception as e:
            print(f"[ERROR] 执行按钮移动失败: {e}")

    def _rearrange_Btns(self):#按钮重排
        scrollbar_width = self.scroll_area.verticalScrollBar().width() \
            if self.scroll_area.verticalScrollBar().isVisible() else 0
        available_width = max(MIN_SIZE, self.width() - 2 * sidelen - scrollbar_width - 4)  # -4为边框宽度
        maxcols = max(1, available_width // (sidelen + SideLen ))
        for i in reversed(range(self.layout.count())):
            widget = self.layout.itemAt(i).widget()
            if widget:
                self.layout.removeWidget(widget)
        for i, button in enumerate(self.buttons):
            row = i//maxcols
            col = i % maxcols
            self.layout.addWidget(button,row,col)

    def ref_btns(self):  #删除旧按钮并重新创建
        try:
            if os.path.exists(self.config_path):
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
            else:
                config = {"LstScanPth": "data/ExeLink", "win_order": [], "win_data": {}}
            for button in self.buttons:
                self.layout.removeWidget(button)
                button.deleteLater()
            self.buttons.clear()
            win_data = config.get("win_data", {})
            win_config = win_data.get(self.win_key, {})
            btns_config = win_config.get("win_btn_data", {})
            btn_order = win_config.get("win_btn_order", [])
            # 更新按钮名称和路径
            self.btn_names = []
            self.btn_paths = []
            for btn_key in btn_order:
                btn_config = btns_config[btn_key]
                btn_name = btn_config.get(f"{btn_key}_N", "")
                btn_path = btn_config.get(f"{btn_key}_Pth", "") or btn_config.get(f"{btn_key}_Url", "")
                if btn_name and btn_path:
                    self.btn_names.append(btn_name)
                    self.btn_paths.append(btn_path)
            self.buttons = self._CR_btns(self.windowTitle(), self.layout)
            self._rearrange_Btns()
            print(f"[INFO] 窗口 {self.win_key} 按钮已刷新")
        except Exception as e:
            print(f"[ERROR] 刷新按钮失败: {e}")

    def _delayed_save_config(self):#按照预定结构保存当前配置
        try:
            if os.path.exists(self.config_path):
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
            else:
                config = {
                    "LstScanPth": "data/ExeLink",
                    "win_order": [],
                    "win_data": {}
                }
            current_geo = self.geometry()
            geo_list = [current_geo.x(), current_geo.y(),
                       current_geo.width(), current_geo.height()]
            win_data = config.get("win_data", {})
            if self.win_key not in win_data:
                win_data[self.win_key] = {
                    "win_btn_order": [],
                    "win_btn_data": {},
                    f"{self.win_key}_N": self.windowTitle(),
                    f"{self.win_key}_Geo": geo_list}
            else:
                win_data[self.win_key][f"{self.win_key}_Geo"] = geo_list
            if self.win_key not in config.get("win_order", []):
                if "win_order" not in config:
                    config["win_order"] = []
                config["win_order"].append(self.win_key)
            config["win_data"] = win_data
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=4)
            print(f"[INFO] 窗口 {self.win_key} 的几何配置已保存")
        except Exception as e:
            print(f"[ERROR] 保存配置失败: {e}")

    def save_config(self):
        self.config_save_timer.stop()
        self.config_save_timer.start(500)

    def resizeEvent(self, event):#窗口大小事件
        if self._is_resizing:
            return super().resizeEvent(event)
        self._is_resizing = True
        try:
            super().resizeEvent(event)
            self._rearrange_Btns()
            self._upd_now_Geoattr()
            self.respos = self.now_Geoattr['rect']
            self.save_config()
            self._adjust_menu_button_position()
            event.accept()
        except Exception as e:
            print(f"ERROR in resizeEvent: {e}")
        finally:
            self._is_resizing = False

    def mouseReleaseEvent(self, event):#鼠标释放事件
        try:
            super().mouseReleaseEvent(event)
            self._upd_now_Geoattr()
            self._apply_Gemotry()
            self.save_config()
        except Exception as e:
            print(f"ERROR in mouseReleaseEvent: {e}")

class GlobalHotkeyListener(QObject):
    Evt = pyqtSignal()
    def __init__(self, hotkey: str = '<alt>+<caps_lock>'):
        super().__init__()
        self.hotkey = hotkey
        self.hotkey_listener = None
        self.listener_thread = None
        self.TGL = True

    def start_listening(self):
        try:
            def on_activate():
                self.Evt.emit()
            self.hotkey_listener = HotKey(
                HotKey.parse(self.hotkey),
                on_activate)
            def for_canonical(f):
                def wrapper(key):
                    canonical_key = self.listener_thread.canonical(key)
                    return f(canonical_key)
                return wrapper
            self.listener_thread = Listener(
                on_press=for_canonical(self.hotkey_listener.press),
                on_release=for_canonical(self.hotkey_listener.release)
            )
            self.listener_thread.daemon = True
            self.listener_thread.start()
            print(f"[INFO] 全局快捷键监听已启动: {self.hotkey}")
        except Exception as e:
            print(f"[ERROR] 启动快捷键监听失败: {e}")

    def stop_listening(self):
        if self.listener_thread:
            self.listener_thread.stop()

class MenuBtn(QPushButton):#菜单按钮类
    def __init__(self, func_name, func, parent=None):  # 初始化菜单按钮
        super().__init__(func_name, parent)
        self.func = func
        self.setFixedSize(64, 64)
        self.clicked.connect(self._on_click)
        self.setStyleSheet("""
                    QPushButton {
                        background-color: transparent;
                        border: 1px solid transparent;
                        border-radius: 4px;
                    }
                    QPushButton:hover {
                        background-color: #404040;
                    }
                    QPushButton:pressed {
                        background-color: #606060;
                    }
                """)

    def _on_click(self):  # 按钮点击事件处理
        try:
            self.func()
        except Exception as e:
            print(f"[ERROR] 执行功能时出错: {e}")

class MessageDialog(QDialog): #用于统一管理项目中的各种通知和输入弹窗,通过控制输入框的可编辑性和默认内容实现不同功能
    # parent:父窗口引用|editable:输入框是否可编辑|default_text:输入框默认显示文本
    # title:窗口标题文本|message:输入框提示信息|placeholder:输入框占位符提示文本
    # style_sheet:样式表|width:对话框宽度|height:对话框高度|auto_close:自动关闭时间
    def __init__(self, parent, editable, default_text, modal,
                 title="提示", message="", placeholder="",
                 style_sheet="", width=300, height=200, auto_close=0):
        super().__init__(parent)
        self.user_input = default_text  # 用户输入内容
        self.result_type = "cancel"  # 用户操作结果类型
        self.dialog = self  # 对话框实例引用
        self.setModal(modal)
        self._setup_ui(title, message, default_text, editable,
                       placeholder, style_sheet, width, height)
        if auto_close > 0:
            QTimer.singleShot(auto_close, self._auto_close_dialog)

    def _setup_ui(self, title, message, default_text, editable,
                  placeholder, style_sheet, width, height):
        self.setWindowTitle(title)
        self.setFixedSize(width, height)
        self.setWindowFlags(Qt.WindowType.Dialog | Qt.WindowType.WindowTitleHint |
                            Qt.WindowType.CustomizeWindowHint)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)
        # 添加提示信息标签（如果有）
        if message:
            message_label = QLabel(message)
            layout.addWidget(message_label)
        # 创建输入框
        self.input_field = QLineEdit()
        self.input_field.setText(default_text)
        self.input_field.setReadOnly(not editable)
        if placeholder:
            self.input_field.setPlaceholderText(placeholder)
        self.input_field.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        layout.addWidget(self.input_field)
        # 创建按钮布局
        button_layout = QHBoxLayout()
        button_layout.setAlignment(Qt.AlignmentFlag.AlignRight)
        self.ok_btn = QPushButton("确定")
        self.cancel_btn = QPushButton("取消")
        button_layout.addWidget(self.ok_btn)
        button_layout.addWidget(self.cancel_btn)
        layout.addLayout(button_layout)
        layout.addStretch(1)
        # 连接信号
        self.ok_btn.clicked.connect(self._on_ok)
        self.cancel_btn.clicked.connect(self._on_cancel)
        self.input_field.returnPressed.connect(self._on_ok)
        # 应用样式表
        if style_sheet:
            self.setStyleSheet(style_sheet)
        else:
            self._apply_default_style()

    def _apply_default_style(self): #应用默认样式表
        default_style = """
        MessageDialog {
            background-color: #2b2b2b;
            color: #ffffff;
            font-family: "Microsoft YaHei";
            font-size: 14px;
            border: 1px solid #555555;
            border-radius: 4px;
        }
         MessageDialog QWidget#qt_calendar_navigationbar {
            background-color: #2b2b2b;
        }
        MessageDialog::title {
            background-color: #2b2b2b;
            color: #ffffff;
            font-size: 14px;
            font-weight: bold;
            padding: 4px 8px;
            height: 20px;
        }
        QLabel {
            color: #cccccc;
            padding: 4px 4px;
            font-size: 13px;
            background-color: transparent;
        }
        QLineEdit {
            background-color: #2e2e2e;
            color: #ffffff;
            border: none;
            border-radius: 2px;
            padding: 6px 8px;
            margin: 4px 8px;
            font-size: 13px;
            selection-background-color: #404040;
        }
        QLineEdit:focus {
            background-color: #323232;
        }
        QLineEdit:read-only {
            background-color: #2e2e2e;
            color: #888888;
        }
        QPushButton {
            background-color: #3a3a3a;
            color: #ffffff;
            border: 1px solid #4a4a4a;
            border-radius: 2px;
            padding: 6px 12px;
            margin: 4px 2px;
            font-size: 13px;
            min-width: 60px;
        }
        QPushButton:hover {
            background-color: #454545;
            border: 1px solid #555555;
        }
        QPushButton:pressed {
            background-color: #505050;
        }
        QPushButton:focus {
            outline: none;
            border: 1px solid #606060;
        }
        """
        self.setStyleSheet(default_style)

    def _on_ok(self): #确定按钮点击事件
        self.user_input = self.input_field.text()
        self.result_type = "ok"
        self.accept()

    def _on_cancel(self): #取消按钮点击事件
        self.result_type = "cancel"
        self.reject()

    def get_result(self): #获取对话框结果
        return {
            "result": self.result_type,
            "input": self.user_input,
            "dialog": self.dialog
        }

    def _auto_close_dialog(self): #自动关闭对话框
        if self.isVisible():
            self.result_type = "auto_close"
            self.accept()

def get_file_icon(file_path):  # 使用获取文件图标
    try:
        print(f"[DEBUG] 开始获取图标，文件路径: {file_path}")

        if file_path.lower().endswith('.lnk'):
            print("[DEBUG] 处理.lnk文件")
            import pythoncom
            from win32com.shell import shell, shellcon
            shortcut = pythoncom.CoCreateInstance(
                shell.CLSID_ShellLink,
                None,
                pythoncom.CLSCTX_INPROC_SERVER,
                shell.IID_IShellLink
            )
            print("[DEBUG] ShellLink实例创建成功")
            shortcut.QueryInterface(pythoncom.IID_IPersistFile).Load(file_path)
            # 获取目标路径
            target_path = shortcut.GetPath(0)[0]
            print(f"[DEBUG] 解析到目标路径: {target_path}")
            file_info = QFileInfo(target_path)
        elif file_path.lower().endswith('.url'):
            print("[DEBUG] 处理.url文件")
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                # 查找 IconFile 属性
                for line in content.split('\n'):
                    if line.lower().startswith('iconfile='):
                        icon_path = line.split('=', 1)[1].strip()
                        print(f"[DEBUG] 找到IconFile路径: {icon_path}")
                        if icon_path and Path(icon_path).exists():
                            icon = QIcon(icon_path)
                            if not icon.isNull():
                                print("[DEBUG] 使用IconFile图标成功")
                                return icon
                        break
            except Exception as e:
                print(f"[ERROR] 获取URL图标失败 {file_path}: {e}")
        else:
            print("[DEBUG] 处理普通文件")
            file_info = QFileInfo(file_path)

        print(f"[DEBUG] 使用QFileIconProvider获取图标")
        icon_provider = QFileIconProvider()
        icon = icon_provider.icon(file_info)
        if not icon.isNull():
            print("[DEBUG] 图标获取成功")
        else:
            print("[DEBUG] 图标获取失败，返回空图标")
        return icon
    except Exception as e:
        print(f"[ERROR] 获取文件图标失败 {file_path}: {e}")
        return QIcon()

def wrap_button_text(text):  # 优化按钮文本显示，支持换行
    # 检查是换行符号
    for symbol in WRAP_SYMBOLS:
        if symbol in text:
            parts = text.split(symbol, 1)  # 只分割第一个出现的符号
            if len(parts) == 2:
                return parts[0] + symbol + '\n' + parts[1]

    # 判断文本类型并设置不同的最大长度
    def is_chinese_char(char):
        # 判断字符是否为中文字符
        return '\u4e00' <= char <= '\u9fff'

    chinese_count = sum(1 for char in text if is_chinese_char(char))
    total_chars = len(text)
    # 如果中文字符占比超过50%，认为是中文文本
    if chinese_count / total_chars > 0.5:
        max_length = MAX_TEXT_LENGTH - 2
    else:
        max_length = MAX_TEXT_LENGTH + 2
    if len(text) > max_length:
        # 寻找合适的换行位置（在空格或标点后）
        mid_point = len(text) // 2
        best_break = -1
        for i in range(mid_point, len(text)):
            if text[i] in [' ', '，', '。', '、', ',', '.', '-', '_']:
                best_break = i
                break
        if best_break == -1:
            for i in range(mid_point, 0, -1):
                if text[i] in [' ', '，', '。', '、', ',', '.', '-', '_']:
                    best_break = i
                    break
        if best_break != -1 and best_break < len(text) - 1:
            return text[:best_break + 1] + '\n' + text[best_break + 1:]
        else:
            # 强制在中间换行
            return text[:mid_point] + '\n' + text[mid_point:]
    return text

def DefMainWinSize():#处理主窗口属性
    global SCR_WIDTH, SCR_HEIGHT, MWIN_WIDTH, MWIN_HEIGHT,START_X, START_Y, HIDE_END_Y
    if SCR_WIDTH is None:
        ScrXY = QApplication.primaryScreen()
        ScrXYTrue = ScrXY.availableGeometry()
        SCR_WIDTH = ScrXYTrue.width()  # 获取屏幕可用宽度
        SCR_HEIGHT = ScrXYTrue.height()  # 获取屏幕可用高度
        initial_w= int(SCR_WIDTH * 0.8)  # 初始窗口宽度为屏幕宽度的80%
        initial_h = int(SCR_HEIGHT * 0.8)  # 初始窗口高度为屏幕高度的80%
        inner_w = initial_w - 4*sidelen #内部可用宽度
        inner_h = initial_h - 4*sidelen
        blocks_x = ceil(inner_w / SideLen)  #方块数量
        blocks_y = ceil(inner_h / SideLen)
        MWIN_WIDTH =  blocks_x*SideLen + 4*sidelen #主窗口最终宽度
        MWIN_HEIGHT = blocks_y*SideLen + 4*sidelen
        START_X = (SCR_WIDTH - MWIN_WIDTH)//2
        START_Y = (-MWIN_HEIGHT + int(MWIN_HEIGHT*0))
        HIDE_END_Y = START_Y

def Child_win(win_key,win_config,parent):#利用配置文件创建子窗口，读取为临时变量并传入子窗口类中
    win_name_temp = win_config.get(f"{win_key}_N", "未命名")
    win_geo_temp = win_config.get(f"{win_key}_Geo", [10, 10, 364, 364])
    btns_config = win_config.get("win_btn_data", {})
    btn_order = win_config.get("win_btn_order", [])
    btn_names_temp = []  # 按钮名称列表_临时
    btn_paths_temp = []  # 按钮路径列表_临时
    # 遍历所有按钮配置，按键名排序确保顺序一致
    for btn_key in btn_order:
        btn_config = btns_config[btn_key]
        btn_name = btn_config.get(f"{btn_key}_N", "")
        btn_path = btn_config.get(f"{btn_key}_Pth", "")
        if btn_name and btn_path:  # 只有当名称和路径都存在时才添加到列表
            btn_names_temp.append(btn_name)
            btn_paths_temp.append(btn_path)
    main_window = None
    if parent and hasattr(parent, 'parent') and parent.parent():
        main_window = parent.parent()
    sub_win = CtrlCTN(
        name=win_name_temp,
        parent=parent,
        btn_names=btn_names_temp,
        btn_paths=btn_paths_temp,
        win_key=win_key,
        win_geo=win_geo_temp,
    )
    return sub_win

def CR_Mwin():#主窗口
    DefMainWinSize()
    Mwin = QMainWindow()
    Mwin.resize(MWIN_WIDTH, MWIN_HEIGHT)#窗口大小
    Mwin.move(START_X,START_Y)#窗口位置
    Mwin.setWindowOpacity(0.75)
    Mwin.setWindowFlags(Mwin.windowFlags() | Qt.WindowType.WindowStaysOnTopHint)#置顶
    Mwin.setAttribute(Qt.WidgetAttribute.WA_QuitOnClose)
    flags = (
            Qt.WindowType.FramelessWindowHint |
            Qt.WindowType.WindowStaysOnTopHint |
            Qt.WindowType.Tool)#窗口属性
    Mwin.setWindowFlags(flags)
    icon_path = "data/Drawer.ico"
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
        icon_path = os.path.join(base_path, "data/Drawer.ico")
    tray_ICON = QSystemTrayIcon(QIcon(icon_path), Mwin)
    tray_menu = QMenu()
    exit_action = tray_menu.addAction("退出")
    exit_action.triggered.connect(QApplication.quit)  # 点击托盘退出菜单时退出程序
    tray_ICON.setContextMenu(tray_menu)
    tray_ICON.show()
    #创建QMdiArea中央控件到Mwin
    M_mdi_area = QMdiArea(Mwin)
    Mwin.setCentralWidget(M_mdi_area)
    #创建子窗口
    try:
        if os.path.exists('config.json'):
            with open('config.json', 'r', encoding='utf-8') as f:
                config = json.load(f)
        else:
            config = None
        if config and 'win_data' in config:
            win_order = config.get("win_order", [])
            win_data = config.get("win_data", {})
            for win_key in win_order:  # 按顺序创建窗口
                if win_key in win_data:
                    win_config = win_data[win_key]  # 获取单个窗口配置
                    sub_win = Child_win(win_key, win_config, M_mdi_area)  # 创建子窗口
                    sub_win.parent_mwin = Mwin
                    M_mdi_area.addSubWindow(sub_win)  # 添加到MDI区域
                    sub_win.show()  # 显示子窗口
        else:  # 如果配置文件不存在，创建默认窗口
            window_0 = CtrlCTN(name='未命名', parent=M_mdi_area)
            M_mdi_area.addSubWindow(window_0)
            window_0.show()
    except Exception as e:
        print(f"[ERROR] 读取配置文件失败: {e}")# 异常处理：配置文件读取或解析失败时
        window_0 = CtrlCTN(name='未命名', parent=M_mdi_area)# 出错时创建默认窗口，确保程序不会崩溃
        M_mdi_area.addSubWindow(window_0)
        window_0.show()
    Mwin.show()
    return Mwin

def Anim_AppearMwin(Mwin):#窗口滑出动画
    Anim = QPropertyAnimation(Mwin,b"geometry",parent=Mwin)
    Anim.setDuration(500);Anim.setEasingCurve(QEasingCurve.Type.InOutQuart)#动画持续时间和方式
    MwinXY = Mwin.geometry()#窗口初始位置
    ENXY = QRect(START_X, APPEAR_END_Y, MWIN_WIDTH, MWIN_HEIGHT)#结束位置
    Anim.setStartValue(MwinXY);Anim.setEndValue(ENXY)
    if AUTO_CAPS:
        Anim.finished.connect(lambda: _toggle_caps_lock_if_on())
    Anim.start()

def Anim_HideMwin(Mwin):
    Anim = QPropertyAnimation(Mwin, b"geometry",parent=Mwin)
    Anim.setDuration(500);Anim.setEasingCurve(QEasingCurve.Type.InOutQuart)#动画持续时间和方式
    MwinXY = Mwin.geometry()#窗口位置
    ENXY = QRect(START_X, HIDE_END_Y, MWIN_WIDTH, MWIN_HEIGHT)#结束位置
    Anim.setStartValue(MwinXY);Anim.setEndValue(ENXY)
    if AUTO_CAPS:
        Anim.finished.connect(lambda: _toggle_caps_lock_if_on())
    Anim.start()

def _toggle_caps_lock_if_on(): #检查大写锁定状态并关闭  # 增
    try:
        if ctypes.windll.user32.GetKeyState(0x14) & 0x0001:
            ctypes.windll.user32.keybd_event(0x14, 0x45, 0x0001, 0)
            ctypes.windll.user32.keybd_event(0x14, 0x45, 0x0001 | 0x0002, 0)
    except Exception as e:
        print(f"[ERROR] 关闭大写锁定失败: {e}")

def BD_kSC(Mwin):#快捷键触发的事件，隐藏动画或显示动画
    GL_Lsn = GlobalHotkeyListener(hotkey='<alt>+<caps_lock>')
    Mwin.hotkey_listener = GL_Lsn
    def toggle_window():
        try:
            if not GL_Lsn.TGL:
                Anim_AppearMwin(Mwin)
                GL_Lsn.TGL = True
            else:
                Anim_HideMwin(Mwin)
                GL_Lsn.TGL = False
        except Exception as e:
            print(f'[ERROR] 切换窗口状态失败: {e}')

    GL_Lsn.Evt.connect(toggle_window)
    GL_Lsn.start_listening()

def add_func_menu_button(Mwin):  # 在主窗口添加功能菜单按钮
    # 创建功能菜单按钮
    func_btn = QPushButton("功能菜单", Mwin)
    func_btn.setFixedSize(100, 30)
    btn_x = (Mwin.width() - func_btn.width()) // 2
    btn_y = Mwin.height() - func_btn.height() - 5
    func_btn.move(btn_x, btn_y)
    func_btn.show()
    # 创建功能菜单弹窗（QWidget实现）
    func_menu = QWidget(Mwin)
    func_menu.setWindowFlags(
        Qt.WindowType.FramelessWindowHint |
        Qt.WindowType.Popup |
        Qt.WindowType.WindowStaysOnTopHint
    )
    func_menu.setFixedSize(384, 384)
    func_menu.setWindowOpacity(0.8)
    func_menu.setStyleSheet("""
        QWidget {
            border: 2px solid #cccccc;
            border-radius: 8px;
        }
    """)
    func_menu.hide()
    # 创建布局和按钮容器
    layout = QGridLayout(func_menu)
    layout.setSpacing(16)
    layout.setContentsMargins(16,16,16,16)
    functions = [
        ("新建窗口", lambda: CRChildWin(Mwin)),
        ("刷新配置", lambda:  Ref_json(Mwin)),
        ("按钮名称", lambda: show_buttons_name(Mwin)),
        ("添加文件", lambda: AddFile(Mwin)),
        ("打开文件夹", lambda: open_RepoFolder()),
        ("设置", lambda: show_settings_dialog(Mwin)),
        ("退出", lambda: QApplication.quit())
    ]
    # 计算弹窗大小
    rows = (len(functions) + 3) // 4
    menu_width = 320
    menu_height = rows * (64 + 16) + 16
    func_menu.setFixedSize(menu_width, menu_height)
    for i, (name, func) in enumerate(functions):
        btn = MenuBtn(name, func, func_menu)
        row = i // 4
        col = i % 4
        layout.addWidget(btn, row, col)
    # 按钮点击事件 - 切换菜单显示/隐藏
    def toggle_func_menu():
        if func_menu.isVisible():
            func_menu.hide()
        else:
            btn_global_pos = func_btn.mapToGlobal(func_btn.rect().topLeft())
            menu_x = btn_global_pos.x() + (func_btn.width() - menu_width) // 2
            menu_y = btn_global_pos.y() - menu_height - 5
            func_menu.move(menu_x, menu_y)
            func_menu.show()
            func_menu.raise_()
    func_btn.clicked.connect(toggle_func_menu)
    return func_btn

def CRChildWin(Mwin):  # 创建新的子窗口
    config_path = "config.json"
    try:
        if os.path.exists(config_path):
            with open(config_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
        else:
            config = {"LstScanPth": "data/ExeLink", "win_order": [], "win_data": {}}
        # 查找下一个可用的窗口键名
        win_index = 0
        win_data = config.get("win_data", {})
        while f"Win_Win{win_index}" in win_data:
            win_index += 1
        win_key = f"Win_Win{win_index}"
        # 创建输入对话框获取窗口名称
        Anim_HideMwin(Mwin)
        dialog = MessageDialog(
            parent=Mwin,editable=True,default_text="",modal=True,
            title="新建子窗口",message="请输入子窗口名称:",placeholder="窗口名称",
            width=350, height=150)
        result = dialog.exec()  # 增
        if result != QDialog.DialogCode.Accepted:  # 用户点击取消或关闭对话框
            Anim_AppearMwin(Mwin)
            return  # 直接返回
        Anim_AppearMwin(Mwin)
        win_name = dialog.user_input
        if not win_name:
            win_name = f"未命名{win_index}"
        # 创建新的窗口配置
        if "win_data" not in config:
            config["win_data"] = {}
        if "win_order" not in config:
            config["win_order"] = []
        config["win_data"][win_key] = {
            "win_btn_order": [],
            "win_btn_data": {},
            f"{win_key}_N": win_name,
            f"{win_key}_Geo": [10, 10, 364, 364]
        }
        config["win_order"].append(win_key)
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=4)
        mdi_area = Mwin.centralWidget()
        sub_win = Child_win(win_key, config["win_data"][win_key], mdi_area)
        sub_win.parent_mwin = Mwin
        mdi_area.addSubWindow(sub_win)
        sub_win.show()
    except Exception as e:
        print(f"[ERROR] 创建新窗口失败: {e}")

def Ref_json(Mwin):  # 刷新配置文件函数
    # 读取当前配置中所有窗口下的按钮名称和路径为元组，将这些元组储存为列表list_JSON
    config_path = "config.json"
    list_JSON = []
    # 初始化 config 和 scan_path
    config = {"LstScanPth": "data/ExeLink", "win_order": [], "win_data": {}}
    scan_path = "data/ExeLink"
    try:
        if os.path.exists(config_path):
            with open(config_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
            scan_path = config.get("LstScanPth", "data/ExeLink")
            # 读取现有配置中的按钮数据
            win_data = config.get("win_data", {})
            for win_key, win_config in win_data.items():
                btns_config = win_config.get("win_btn_data", {})
                btn_order = win_config.get("win_btn_order", [])
                for btn_key in btn_order:  # 按顺序遍历按钮
                    if btn_key in btns_config:
                        btn_config = btns_config[btn_key]
                        name = btn_config.get(f"{btn_key}_N", "")
                        path = btn_config.get(f"{btn_key}_Pth", "") or btn_config.get(f"{btn_key}_Url", "")
                        if name and path:
                            list_JSON.append((name, path))
                        else:
                            print(f"-无效按钮: {btn_key} - 名称: '{name}', 路径: '{path}'")
    except Exception as e:
        print(f"[ERROR] 读取配置文件失败: {e}")
        # 继续执行，使用默认配置
    # 扫描储存快捷方式的文件夹，读取其中每个快捷方式的名称和路径储存为元组，将这些元组储存为列表list_FILE
    list_FILE = []
    try:
        link_path = Path(scan_path)
        if link_path.exists():
            Lnk_files = list(link_path.glob('*.lnk'))
            Url_files = list(link_path.glob('*.url'))
            all_files = Lnk_files + Url_files
            print(f'找到 {len(Lnk_files)} 个.lnk文件, {len(Url_files)} 个.url文件')
            for file in all_files:
                # 修改：直接使用文件路径，不解析.url文件内容
                file_info = QFileInfo(str(file))
                file_name = file_info.baseName()
                # 使用文件的绝对路径，与"复制文件地址"一致
                list_FILE.append((file_name, str(file.resolve())))
        else:
            print(f"[WARNING] 路径不存在: {scan_path}")
            return
    except Exception as e:
        print(f"[ERROR] 扫描快捷方式时出错: {e}")
        return
    # 处理两个列表，方法是以list_FILE为基准，去除其中与list_JSON相同的部分，得到list_FILE_inc_JSON
    list_FILE_inc_JSON = []
    for file_item in list_FILE:
        # 修改：只对比文件路径部分，忽略按钮名称
        item_exists = False
        file_name, file_path = file_item
        for json_item in list_JSON:
            json_name, json_path = json_item
            # 仅比较文件路径，忽略名称差异
            if file_path.strip().lower() == json_path.strip().lower():
                item_exists = True
                break
        if not item_exists:
            list_FILE_inc_JSON.append(file_item)
    # 把列表list_FILE_inc_JSON中的元组写入默认Win0的配置中
    if "win_data" not in config:
        config["win_data"] = {}
    if "Win_Win0" not in config["win_data"]:
        config["win_data"]["Win_Win0"] = {
            "win_btn_order": [],
            "win_btn_data": {},
            "Win_Win0_N": "未命名",
            "Win_Win0_Geo": [10, 10, 364, 364]
        }
    win0_config = config["win_data"]["Win_Win0"]
    win0_btns = win0_config.get("win_btn_data", {})
    win0_btn_order = win0_config.get("win_btn_order", [])
    # 找出当前最大的按钮索引
    max_index = -1
    for btn_key in win0_btns.keys():
        if btn_key.startswith("Win_Win0_Btn"):
            try:
                index = int(btn_key.replace("Win_Win0_Btn", ""))
                if index > max_index:
                    max_index = index
            except ValueError:
                pass
    # 添加新的按钮
    for name, path in list_FILE_inc_JSON:
        max_index += 1
        btn_key = f"Win_Win0_Btn{max_index}"
        # 修改：统一使用 _Pth 键名，不区分 lnk 或 url
        win0_btns[btn_key] = {
            f"{btn_key}_N": name,
            f"{btn_key}_Pth": path
        }
        win0_btn_order.append(btn_key)
    win0_config["win_btn_data"] = win0_btns
    win0_config["win_btn_order"] = win0_btn_order
    config["win_data"]["Win_Win0"] = win0_config
    # 确保Win_Win0在窗口顺序列表中
    if "win_order" not in config:
        config["win_order"] = []
    if "Win_Win0" not in config["win_order"]:
        config["win_order"].insert(0, "Win_Win0")
    # 保存更新后的配置
    try:
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=4)
    except Exception as e:
        print(f"[ERROR] 保存配置文件时出错: {e}")
        return
    mdi_area = Mwin.centralWidget()
    try:
        if os.path.exists(config_path):
            if reorganize_config_numbers(config):  # 传递config参数
                with open(config_path, 'w', encoding='utf-8') as f:
                    json.dump(config, f, ensure_ascii=False, indent=4)
                print("[INFO] 配置文件序号重整完成并保存")
                # 重整后再次刷新UI
                for sub_window in mdi_area.subWindowList():
                    if hasattr(sub_window, 'config_updated_signal'):
                        sub_window.config_updated_signal.emit()
    except Exception as e:
        print(f"[ERROR] 重整序号失败: {e}")

def show_buttons_name(Mwin): #在所有子窗口的按钮上显示名称覆盖层

    def create_overlay(parent_button, display_text): #创建覆盖层组件
        try:
            wrapped_text = wrap_button_text(display_text)
            # 创建覆盖层
            overlay = QLabel(wrapped_text, parent_button)
            overlay.setAlignment(Qt.AlignmentFlag.AlignCenter)
            overlay.setStyleSheet("""
                QLabel {
                    background-color: rgba(0, 0, 0, 180);
                    color: white;
                    border: 1px solid rgba(255, 255, 255, 100);
                    border-radius: 4px;
                    font-size: 9px;
                    padding: 2px;
                }
            """)
            overlay.setGeometry(0, 0, parent_button.width(), parent_button.height())
            overlay.show()
            QTimer.singleShot(5000, overlay.deleteLater)
        except Exception as e:
            print(f"[ERROR] 创建覆盖层失败: {e}")
    try:
        # 获取MDI区域
        mdi_area = Mwin.centralWidget()
        config_path = "config.json"
        with open(config_path, 'r', encoding='utf-8') as f:
            config = json.load(f)
        win_data = config.get("win_data", {})
        # 遍历所有子窗口
        for sub_window in mdi_area.subWindowList():
            if hasattr(sub_window, 'content_widget') and hasattr(sub_window, 'buttons'):
                win_key = sub_window.win_key
                win_config = win_data.get(win_key, {})
                btns_config = win_config.get("win_btn_data", {})
                btn_order = win_config.get("win_btn_order", [])
                # 遍历子窗口中的所有按钮
                for i, button in enumerate(sub_window.buttons):
                    if i < len(btn_order):
                        btn_key = btn_order[i]
                        btn_config = btns_config.get(btn_key, {})
                        btn_name = btn_config.get(f"{btn_key}_N", f"按钮{i}")
                        create_overlay(button, btn_name)
                    else:
                        create_overlay(button, f"按钮{i}")
    except Exception as e:
        print(f"[ERROR] 显示按钮名称失败: {e}")

def AddFile(Mwin):  # 打开文件选择器，创建快捷方式到data/ExeLink目录

    def _copy_url_file(source_path, target_dir, file_name):  # 复制.url文件到目标目录
        try:
            target_path = target_dir / f"{file_name}.url"
            # 如果文件名已存在，添加数字后缀
            counter = 1
            while target_path.exists():
                target_path = target_dir / f"{file_name} ({counter}).url"
                counter += 1
            shutil.copy2(source_path, target_path)
            print(f"[INFO] 已复制.url文件: {source_path} -> {target_path}")
        except Exception as e:
            print(f"[ERROR] 复制.url文件失败: {e}")
            raise

    def _create_lnk_shortcut(target_path, target_dir, file_name):  # 创建.lnk快捷方式
        try:
            # 创建快捷方式文件名
            lnk_name = f"{file_name}.lnk"
            lnk_path = target_dir / lnk_name
            # 如果文件名已存在，添加数字后缀
            counter = 1
            while lnk_path.exists():
                lnk_name = f"{file_name} ({counter}).lnk"
                lnk_path = target_dir / lnk_name
                counter += 1
            # 使用win32com创建快捷方式
            shell = win32com.client.Dispatch("WScript.Shell")
            shortcut = shell.CreateShortcut(str(lnk_path))
            shortcut.TargetPath = str(target_path)
            shortcut.WorkingDirectory = str(Path(target_path).parent)
            shortcut.Save()
            print(f"[INFO] 已创建.lnk快捷方式: {target_path} -> {lnk_path}")
        except Exception as e:
            print(f"[ERROR] 创建.lnk快捷方式失败: {e}")
            raise

    try:
        # 选择文件或文件夹
        file_path, _ = QFileDialog.getOpenFileName(
            Mwin,
            "选择要添加快捷方式的文件",
            "",
            "所有文件 (*.*)"
        )
        if not file_path:
            return  # 用户取消选择
        target_dir = Path("data/ExeLink")
        if not target_dir.exists():
            target_dir.mkdir(parents=True, exist_ok=True)
        file_name = Path(file_path).stem
        # 根据文件类型处理
        if file_path.lower().endswith('.url'):
            _copy_url_file(file_path, target_dir, file_name)
        else:
            _create_lnk_shortcut(file_path, target_dir, file_name)
        # 显示成功消息
        dialog = MessageDialog(
            parent=Mwin,editable=False,default_text="",modal=False,
            title="成功",message=f"已为 {file_name} 创建快捷方式",
            width=350,height=150,auto_close=1000)
        dialog.show()
        Ref_json(Mwin)
    except Exception as e:
        print(f"[ERROR] 添加文件失败: {e}")
        QMessageBox.warning(Mwin, "错误", f"添加文件失败: {e}")

def open_RepoFolder():  # 用文件资源管理器打开data\ExeLink文件夹
    try:
        folder_path = Path("data/ExeLink")
        if not folder_path.exists():
            folder_path.mkdir(parents=True, exist_ok=True)
        os.startfile(str(folder_path.resolve()))
    except Exception as e:
        print(f"[ERROR] 打开文件夹失败: {e}")

def show_settings_dialog(Mwin):  # 增
    try:
        Anim_HideMwin(Mwin)
        dialog = MessageDialog(
            parent=Mwin, editable=False, default_text="敬请期待", modal=True,
            title="设置", message="还没做完",
            width=350, height=150, auto_close=4000)
        def on_dialog_finished():
            Anim_AppearMwin(Mwin)
        dialog.finished.connect(on_dialog_finished)
        dialog.exec()
    except Exception as e:
        print(f"[ERROR] 显示设置对话框失败: {e}")
        Anim_AppearMwin(Mwin)
#visual.py
import ctypes
import json
import os
import shutil
import sys
from resources_rc import ResourceLoader
from math import ceil
from pathlib import Path

import win32api
import win32com.client
import win32con
import win32timezone
import win32event
import winerror
from PyQt6.QtCore import (QPropertyAnimation, QEasingCurve, QRect, Qt, QTimer, QTime, QDateTime, QObject,
                          pyqtSignal, QFileInfo, QSize, QProcess)
from PyQt6.QtGui import QIcon, QCursor, QFont, QPixmap ,QPainter
from PyQt6.QtWidgets import (QMainWindow, QApplication, QSystemTrayIcon, QMdiArea, QMdiSubWindow, QWidget, QVBoxLayout,
                             QLabel, QSizePolicy, QHBoxLayout, QPushButton, QToolButton, QGridLayout,
                             QScrollArea, QDialog, QLineEdit, QMenu, QMessageBox, QFileIconProvider,
                             QFileDialog, QCheckBox, QSpinBox, QComboBox, QSlider, QGroupBox, QLCDNumber)
from pynput.keyboard import HotKey, Listener

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
        state_key = f"{win_key}_State"
        if state_key not in win_config:
            win_config[state_key] = True
        btn_order = win_config.get("win_btn_order", [])
        btn_data = win_config.get("win_btn_data", {})
        for btn_key in btn_order:
            if btn_key not in btn_data:
                return False
    return True  # 所有检查通过返回True

def reorganize_config_numbers(config):#重整配置文件中窗口和按钮的序号，使其按顺序排列
    try:
        # === 窗口部分的重整 ===
        win_order = config.get("win_order", [])
        win_data = config.get("win_data", {})
        # 记录窗口列表长度
        win_count = len(win_order)
        # 生成新的窗口顺序列表
        new_win_order = [f"Win_Win{i}" for i in range(win_count)]
        # 创建新的窗口数据字典
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
            # 更新窗口状态键
            old_state_key = f"{old_win_key}_State"  # 改
            new_state_key = f"{new_win_key}_State"  # 改
            if old_state_key in win_config:  # 改
                win_config[new_state_key] = win_config.pop(old_state_key)
            # 更新窗口几何键
            old_geo_key = f"{old_win_key}_Geo"
            new_geo_key = f"{new_win_key}_Geo"
            if old_geo_key in win_config:
                win_config[new_geo_key] = win_config.pop(old_geo_key)
            # === 按钮部分的重整 ===
            btn_order = win_config.get("win_btn_order", [])
            btn_data = win_config.get("win_btn_data", {})
            # 记录按钮列表长度
            btn_count = len(btn_order)
            # 生成新的按钮顺序列表
            new_btn_order = [f"{new_win_key}_Btn{i}" for i in range(btn_count)]
            # 创建新的按钮数据字典
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
        return True
    except Exception as e:
        print(f"[ERROR] 重整配置文件序号失败: {e}")
        return False

# 全局变量定义
SCR_WIDTH = None  # 屏幕可用宽度（像素）
SCR_HEIGHT = None  # 屏幕可用高度（像素）
MWIN_WIDTH = None  # 主窗口宽度（像素）
MWIN_HEIGHT = None  # 主窗口高度（像素）
START_X = None  # 窗口初始X坐标（像素），水平居中
START_Y = None  # 窗口初始Y坐标（像素）
APPEAR_END_Y = 0   # 弹出动画结束Y坐标（像素）
HIDE_END_Y = None  # 隐藏动画结束Y坐标（像素）
SideLen = 64  # 按钮的边长（像素）
sidelen = 16  # 间隔的距离（像素）
MIN_SIZE  = sidelen*2 + SideLen  # 最小尺寸限制
XYCtrlCTN = (364, 364)  # 窗口默认大小
MAX_TEXT_LENGTH = 8  # 每行最大字符数
WRAP_SYMBOLS = ['：', ':', '-', ' ']  # 触发换行的符号

REST_TIMER = None  # 休息提示计时器
REST_ENABLED = False  # 休息提示功能开关
REST_INTERVAL = 30 # 休息间隔时长（分钟）
AUTO_HIDE = True  # 自动收起功能开关
WIN_KEY_ENABLED = False # Win键响应开关
AUTO_CAPS = True  # 自动关闭大写锁定开关
REST_PROMPT_LOOP = False  # 息息提示循环开关
HOTKEY = '<alt>+<caps_lock>' # 快捷键设置
OPACITY = 0.75  # 窗口透明度
ColorList = ["#000000", "#FFFFFF", "#1a1a1a", "#2d2d2d", "#404040", "#5a5a5a", "#747474"] # 主题颜色列表
WIN_RATIO = 0.8 # 主窗口大小
GRID_SNAP_ENABLED = True  # 网格吸附开关
GRID_CELL_SIZE = 16 # 网格单元大小
GRID_INFO = None # 存储网格信息的字典
REALIZED_SIZE = None # 显化尺寸
ICON_CACHE = {} # 图标缓存
try:
    ColorList_Dk = ["#000000", "#FFFFFF", "#1a1a1a", "#2d2d2d", "#404040", "#5a5a5a", "#747474"]
    ColorList_Lt = ["#FFFFFF", "#000000", "#cdcdcd", "#d7d7d7", "#e1e1e1", "#ebebeb", "#f5f5f5"]
    config_path = "config.json"
    if os.path.exists(config_path):
        with open(config_path, 'r', encoding='utf-8') as f:
            config = json.load(f)
        settings = config.get("settings", {})
        # 读取各配置项，如果不存在则使用默认值
        UI_SIZE_MULTIPLIERS = {"大": 3, "中": 2, "小": 1}
        ui_size_setting = settings.get("ui_size", "中")  # 默认中等尺寸
        UI_SCALE_FACTOR = UI_SIZE_MULTIPLIERS.get(ui_size_setting, 2)
        GRID_SNAP_ENABLED = settings.get("grid_snap", False)
        WIN_KEY_ENABLED = settings.get("win_key_enabled", False)
        AUTO_CAPS = settings.get("auto_caps", AUTO_CAPS)
        AUTO_HIDE = settings.get("auto_hide", AUTO_HIDE)
        HOTKEY = settings.get("hotkey", HOTKEY)
        REST_PROMPT_LOOP = settings.get("rest_prompt_loop", REST_PROMPT_LOOP)
        OPACITY = settings.get("opacity", OPACITY)
        WIN_RATIO = settings.get("win_ratio", WIN_RATIO)
        theme = settings.get("theme", "深色主题")
        ColorList = ColorList_Lt if theme == "浅色主题" else ColorList_Dk
        print(f'[INFO] 配置文件加载成功')
    else:
        print("[INFO] 配置文件不存在，使用默认设置")
except Exception as e:
    print(f"[ERROR] 加载配置文件失败: {e}")

class CtrlCTN(QMdiSubWindow):  #子窗口类
    config_updated_signal = pyqtSignal()
    window_folded_signal = pyqtSignal(str, object)
    window_expand_signal = pyqtSignal()
    def __init__(self,name='未分类',parent=None,btn_names=None,btn_paths=None,
                 win_key='Win_Win0',win_geo=None, win_state=True):
        super().__init__(parent)
        if win_geo and len(win_geo) == 4:
            x, y, width, height = win_geo
            self.setGeometry(x, y, width, height)
        else:
            self.setGeometry(sidelen, sidelen, XYCtrlCTN[0], XYCtrlCTN[1])
        if GRID_SNAP_ENABLED and GRID_INFO:
            current_geo = self.geometry()
            snapped_geo = self._snap_to_grid(current_geo)
            if snapped_geo != current_geo:
                self.setGeometry(snapped_geo)
        self.setWindowTitle(name)
        self.setWindowFlags(Qt.WindowType.CustomizeWindowHint |
                            Qt.WindowType.SubWindow |
                            Qt.WindowType.WindowTitleHint )
        sub_window_style = (f"QMdiSubWindow {{ "f"QMdiSubWindow::title {int(sidelen)};"
                            f"border: 2px solid {ColorList[6]} !important; background-color: {ColorList[2]};}}")
        self.setStyleSheet(sub_window_style)
        self.config_updated_signal.connect(self.ref_btns)
        self.window_expand_signal.connect(self._handle_expand_signal)
        self.now_Geoattr = None
        if GRID_SNAP_ENABLED and GRID_INFO:
            self.default_pos = (GRID_INFO['start_pos'][0], GRID_INFO['start_pos'][1])
        else:
            self.default_pos = (9,9)
        self.border_margin = sidelen
        self._is_resizing = False
        self.btn_names = btn_names or []
        self.btn_paths = btn_paths or []
        self.win_key = win_key
        self.win_geo = win_geo
        self.config_path = "config.json"
        self.parent_mwin = parent
        self.win_state = win_state
        #创建 QScrollArea 作为内容容器，支持滚动条
        self.scroll_area = QScrollArea(self)
        self.scroll_area.setWidgetResizable(True)  # 允许内容 widget 随滚动区域调整
        self.scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff) #垂直滚动条按需显示
        self.scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)#禁用横向滚动条
        self.content_widget = QWidget()
        self.scroll_area.setWidget(self.content_widget)
        #创建子窗口布局管理器
        self.layout = QGridLayout(self.content_widget)
        self.setWidget(self.scroll_area)
        self.layout.setHorizontalSpacing(sidelen)
        self.layout.setVerticalSpacing(sidelen)
        self.layout.setContentsMargins(sidelen, sidelen, sidelen, sidelen)
        margin = SideLen // 4 # 子窗口最小尺寸
        base_w = SideLen + margin * 2;base_h = base_w
        if GRID_INFO:
            cs = GRID_INFO['cell_size']
            min_w = ceil(base_w / cs) * cs
            min_h = ceil(base_h / cs) * cs
        else:min_w, min_h = base_w, base_h
        self.setMinimumSize(min_w, min_h)
        # 创建菜单按钮
        self.menu_btn = QPushButton("≡", self)
        self.menu_btn.setFixedSize(sidelen, sidelen)
        self.menu_btn.setObjectName("windowMenuBtn")
        window_menu_btn_style = f"""
        QPushButton#windowMenuBtn {{
            background-color: transparent;border: none;color: {ColorList[1]};
            font-weight: bold;font-size: {int(sidelen * 0.875)}px;}}
        QPushButton#windowMenuBtn:hover {{background-color: {ColorList[6]};border-radius: {int(sidelen * 0.125)}px;}}
        QPushButton#windowMenuBtn:pressed {{background-color: {ColorList[5]};}}"""
        self.menu_btn.setStyleSheet(window_menu_btn_style)
        # 创建功能菜单
        self.window_menu = QMenu(self)
        self.window_menu.setWindowOpacity(0.8)
        rename_action = self.window_menu.addAction("重命名窗口")
        rename_action.triggered.connect(self._rename_window)
        fold_action = self.window_menu.addAction("折叠窗口")
        fold_action.triggered.connect(self._fold_window)
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
        if not self.win_state:
            self.hide()
            QTimer.singleShot(100, lambda: self.window_folded_signal.emit(self.windowTitle(), self))

    def _adjust_menu_button_position(self):  # 调整菜单按钮位置到右上角
        try:
            border_width = int(sidelen * 0.125)
            btn_x = int(self.width() - self.menu_btn.width() - border_width ) # 额外4px内边距
            btn_y = int(border_width)
            self.menu_btn.move(btn_x, btn_y)
        except Exception as e:
            print(f"调整菜单按钮位置时出错: {e}")
    def _snap_to_grid(self, geometry):  # 将几何矩形吸附到最近的网格位置
        try:
            if not GRID_INFO:
                return geometry
            cell_size = GRID_INFO['cell_size']
            grid_start_x, grid_start_y = GRID_INFO['start_pos']
            legal_area = GRID_INFO['legal_area']
            # 计算在网格中的坐标
            grid_x = round((geometry.x() - grid_start_x) / cell_size)
            grid_y = round((geometry.y() - grid_start_y) / cell_size)
            # 计算网格中的尺寸（最小为1个网格单元）
            grid_width = max(1, round(geometry.width() / cell_size))
            grid_height = max(1, round(geometry.height() / cell_size))
            # 转换为实际像素坐标
            x = grid_start_x + grid_x * cell_size
            y = grid_start_y + grid_y * cell_size
            width = grid_width * cell_size
            height = grid_height * cell_size
            # 确保在合适区域内
            snapped_rect = QRect(x, y, width, height)
            if not legal_area.contains(snapped_rect):
                # 如果超出边界，调整到最近合适位置
                if x < legal_area.x():
                    x = legal_area.x()
                if y < legal_area.y():
                    y = legal_area.y()
                if x + width > legal_area.right():
                    x = legal_area.right() - width
                if y + height > legal_area.bottom():
                    y = legal_area.bottom() - height
                snapped_rect = QRect(x, y, width, height)
            return snapped_rect
        except Exception as e:
            print(f"[ERROR] 网格吸附失败: {e}")
            return geometry

    def _show_window_menu(self):  # 显示窗口菜单
        btn_global_pos = self.menu_btn.mapToGlobal(self.menu_btn.rect().bottomRight())
        self.window_menu.exec(btn_global_pos)
    def _rename_window(self): # 重命名窗口
        try:
            old_name = self.windowTitle()
            dialog = MessageDialog(
                parent=self, editable=True, default_text=old_name, modal=True,
                title="修改窗口名称", message="输入新的窗口名称")
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
    def _delete_window(self):  # 删除当前窗口，并将按钮配置移动到默认窗口中
        try:
            if self.win_key == "Win_Win0":
                dialog = MessageDialog(
                    parent=self,editable=False,default_text="",modal=True,
                    title="无法删除",message="默认窗口不能删除")
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

    def _fold_window(self):  # 折叠窗口功能
        self.hide()
        self.win_state = False
        try:
            config_path = "config.json"
            if os.path.exists(config_path):
                with open(config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
            else:
                config = {"LstScanPth": "data/ExeLink", "win_order": [], "win_data": {}}
            win_data = config.get("win_data", {})
            if self.win_key in win_data:
                state_key = f"{self.win_key}_State"
                win_data[self.win_key][state_key] = False
                config["win_data"] = win_data
                with open(config_path, 'w', encoding='utf-8') as f:
                    json.dump(config, f, ensure_ascii=False, indent=4)
        except Exception as e:
            print(f"[ERROR]更新配置失败: {e}")
        self.window_folded_signal.emit(self.windowTitle(), self)
    def _handle_expand_signal(self):  # 处理展开信号
        self.show()
        self.win_state = True
        try:
            config_path = "config.json"
            if os.path.exists(config_path):
                with open(config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
            else:
                config = {"LstScanPth": "data/ExeLink", "win_order": [], "win_data": {}}
            win_data = config.get("win_data", {})
            if self.win_key in win_data:
                state_key = f"{self.win_key}_State"
                win_data[self.win_key][state_key] = True
                config["win_data"] = win_data
                with open(config_path, 'w', encoding='utf-8') as f:
                    json.dump(config, f, ensure_ascii=False, indent=4)
        except Exception as e:
            print(f"[ERROR]更新配置失败: {e}")
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
            if GRID_SNAP_ENABLED and GRID_INFO:
                current_rect = self.now_Geoattr['rect']
                legal_area = GRID_INFO['legal_area']
                if not legal_area.contains(current_rect):
                    return 'F'  # 超出网格边界
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
                    if GRID_SNAP_ENABLED and GRID_INFO:
                        snapped_rect = self._snap_to_grid(self.now_Geoattr['rect'])
                        self.setGeometry(snapped_rect)
                    else:
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
            button.setFixedSize(SideLen, SideLen)
            if i < len(self.btn_paths):
                path = self.btn_paths[i]
                # 获取文件图标
                icon = get_file_icon(path)
                if not icon.isNull():
                    button.setIconSize(QSize(int(SideLen*0.75),int(SideLen*0.75)))
                    button.setIcon(icon)
                    button.setText("")
                else:
                    button.setText(wrap_button_text(buttonTEXT))
                button.clicked.connect(lambda checked, idx=i: self._open_button_file(idx))
                grid_button_style = f"""
                QPushButton {{background-color: transparent;border: none;border-radius: {int(sidelen * 0.25)}px;
                    padding: {int(sidelen * 0.125)}px;text-align: center;font-size:{int(sidelen * 0.625)}px;color: {ColorList[1]};}}
                QPushButton:hover {{background-color: {ColorList[4]};}}
                QPushButton:pressed {{background-color: {ColorList[5]};}}"""
                button.setStyleSheet(grid_button_style)
            self._setup_button_context_menu(button, i)
            row = i//maxcols
            col = i % maxcols
            layout.addWidget(button,row,col)
            buttons.append(button)
        return buttons

    def _open_button_file(self, button_index):  # 打开按钮对应路径
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
                return
            try:
                result = win32api.ShellExecute(0, "open", btn_path, None, None, win32con.SW_SHOWNORMAL)
                if result <= 32:
                    print(f"[WARNING] ShellExecute返回错误代码 {result} for {btn_path}")
                    os.startfile(btn_path)
            except Exception as e:
                print(f"[ERROR] 打开文件失败 {btn_path}: {e}")
            if AUTO_HIDE and self.parent_mwin:
                print('打开按钮对应文件')
                Mwin_ToggleState(self.parent_mwin)
        except Exception as e:
            print(f"[ERROR] 打开按钮文件失败: {e}")

    def _setup_button_context_menu(self, button, index):#为按钮设置右键菜单
        button.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        button.customContextMenuRequested.connect(
            lambda pos, btn=button, idx=index: self._show_button_context_menu(btn, idx, pos))
        button.installEventFilter(self)
    def _show_button_context_menu(self, button, index, position):#显示按钮的右键菜单
        # 创建菜单容器
        menu_container = QMenu(self)
        menu_container.setWindowFlags(
            Qt.WindowType.FramelessWindowHint |
            Qt.WindowType.Popup |
            Qt.WindowType.WindowStaysOnTopHint)
        menu_container.setWindowOpacity(0.8)
        button_context_menu_style = f"""
           QMenu {{background-color: {ColorList[2]};border: 1px solid {ColorList[6]};padding: {int(sidelen * 0.25)}px;
               font-size: {int(sidelen * 0.75)}px;}}
           QMenu::item {{padding: {int(sidelen * 0.5)}px {int(sidelen)}px;background-color: transparent;
               color: {ColorList[1]};min-width: {int(sidelen * 6)}px;}}
           QMenu::item:selected {{background-color: {ColorList[4]};border-radius: {int(sidelen * 0.25)}px;}}
           QMenu::separator {{height: 1px;background-color: {ColorList[5]};
               margin: {int(sidelen * 0.25)}px {int(sidelen * 0.5)}px;}}"""
        menu_container.setStyleSheet(button_context_menu_style)
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
    def _rename_button(self, button_index): # 重命名按钮
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
                title="修改按钮名称", message="输入新的按钮名称")
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
                                path_obj.unlink()
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
            dialog = MessageDialog(
                parent=self, editable=False, default_text=btn_name, modal=False,
                title="确认删除", message="确定要删除这个按钮吗？",
                auto_close=3000)
            def on_dialog_finished(result_code):
                # 只有明确点击取消按钮才不执行删除
                if result_code == QDialog.DialogCode.Rejected:
                    print("[INFO] 用户取消删除操作")
                    return
                perform_deletion()
            dialog.finished.connect(on_dialog_finished)
            dialog.show()

        except Exception as e:
            print(f"[ERROR] 删除按钮失败: {e}")
    def _move_button_to_window(self, button_index):  # 移动按钮到其他窗口
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
                Qt.WindowType.WindowStaysOnTopHint)
            menu_container.setWindowOpacity(0.7)
            move_menu_style = f"""
                QWidget {{background-color: {ColorList[2]};border: 1px solid {ColorList[5]};
                padding: {int(sidelen * 0.125)}px;}}
                QPushButton {{
                    background-color: {ColorList[2]};color: {ColorList[1]};border: none;
                    padding: {int(sidelen * 0.25)}px {int(sidelen * 0.25)}px; text-align: left;
                    font-size: {int(sidelen * 0.75)}px;min-width: {int(sidelen * 8)}px;min-height: {int(sidelen * 1.5)}px;}}
                QPushButton:hover {{background-color: {ColorList[4]};border-radius: {int(sidelen * 0.25)}px;}}"""
            menu_container.setStyleSheet(move_menu_style)
            layout = QVBoxLayout(menu_container)
            layout.setSpacing(0)
            layout.setContentsMargins(0, 0, 0, 0)
            # 添加窗口选项
            for win_key, win_name in other_windows.items():
                win_btn = QPushButton(win_name, menu_container)
                win_btn.setFixedSize(int(SideLen), int(SideLen*0.5))
                win_btn.clicked.connect(
                    lambda checked, target_win=win_key: self._perform_button_move(
                        config, btn_key, btn_config, target_win, menu_container))
                layout.addWidget(win_btn)
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
        available_width = max(MIN_SIZE, self.width() - 2 * sidelen - 4)  # -4为边框宽度
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
            if GRID_SNAP_ENABLED and GRID_INFO:
                current_rect = self.now_Geoattr['rect']
                snapped_rect = self._snap_to_grid(current_rect)
                if snapped_rect != current_rect:
                    self.setGeometry(snapped_rect)
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
            if GRID_SNAP_ENABLED and GRID_INFO:
                current_rect = self.now_Geoattr['rect']
                snapped_rect = self._snap_to_grid(current_rect)
                if snapped_rect != current_rect:
                    self.setGeometry(snapped_rect)
                    self._upd_now_Geoattr()
            self._apply_Gemotry()
            self.save_config()
        except Exception as e:
            print(f"ERROR in mouseReleaseEvent: {e}")

class GlobalHotkeyListener(QObject):
    Evt = pyqtSignal()
    WinKeyEvt = pyqtSignal()
    def __init__(self, hotkey: str = HOTKEY):
        super().__init__()
        self.hotkey = hotkey or HOTKEY
        self.hotkeys_listener = []
        self.listener_thread = None
        self.TGL = False

    def start_listening(self):
        try:
            def on_activate():
                self.Evt.emit()
            def on_win_key():
                if WIN_KEY_ENABLED:  # 检查开关状态
                    self.WinKeyEvt.emit()
            self.hotkeys_listener = [
                HotKey(HotKey.parse(self.hotkey), on_activate),
                HotKey(HotKey.parse('<cmd>'), on_win_key)
            ]
            def for_canonical(f):
                def wrapper(key):
                    canonical_key = self.listener_thread.canonical(key)
                    return f(canonical_key)
                return wrapper
            def on_press(key):
                for hotkey in self.hotkeys_listener:
                    hotkey.press(self.listener_thread.canonical(key))
            def on_release(key):
                for hotkey in self.hotkeys_listener:
                    hotkey.release(self.listener_thread.canonical(key))
            self.listener_thread = Listener(
                on_press=on_press,
                on_release=on_release)
            self.listener_thread.daemon = True
            self.listener_thread.start()
            print(f"[INFO] 全局快捷键监听已启动: {self.hotkey}")
        except Exception as e:
            print(f"[ERROR] 启动快捷键监听失败: {e}")

    def stop_listening(self):
        if self.listener_thread:
            self.listener_thread.stop()

class MessageDialog(QDialog): #用于统一管理项目中的各种通知和输入弹窗,通过控制输入框的可编辑性和默认内容实现不同功能
    # parent:父窗口引用|editable:输入框是否可编辑|default_text:输入框默认显示文本
    # title:窗口标题文本|message:输入框提示信息|placeholder:输入框占位符提示文本
    # style_sheet:样式表|width:对话框宽度|height:对话框高度|auto_close:自动关闭时间
    def __init__(self, parent, editable, default_text, modal,
                 title="提示", message="", placeholder="",
                 style_sheet="", width=SideLen*6, height=SideLen*4, auto_close=0):
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
        layout.setContentsMargins(int(sidelen*0.25),int(sidelen*0.25),int(sidelen*0.25),int(sidelen*0.25))
        layout.setSpacing(int(sidelen*0.5))
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
        default_style = f"""
        MessageDialog {{background-color: {ColorList[2]};color: {ColorList[2]};font-family: "Microsoft YaHei";
            font-size: {int(sidelen)}px;border: 1px solid {ColorList[5]};border-radius: {int(sidelen * 0.025)}px;}}
        MessageDialog QWidget#qt_calendar_navigationbar {{background-color: {ColorList[2]};}}
        MessageDialog::title {{background-color: {ColorList[2]};color: {ColorList[1]};font-size: {int(sidelen)}px;
            font-weight: bold; min-height:: {int(SideLen)}px;}}
        QLabel {{color: {ColorList[1]};padding: 4px 4px;font-size: {int(sidelen)}px;background-color: transparent;}}
        QLineEdit {{
            background-color: {ColorList[3]};color: {ColorList[1]};border: none;border-radius: 2px;
            padding: 6px 8px;margin: 4px 8px;font-size: {int(sidelen)}px;selection-background-color: {ColorList[4]};}}
        QLineEdit:focus {{background-color: {ColorList[4]};}}
        QLineEdit:read-only {{background-color: {ColorList[3]};color: {ColorList[6]};}}
        QPushButton {{background-color: {ColorList[3]};color: {ColorList[1]};border: 1px solid {ColorList[4]};
            border-radius: 2px;padding: 6px 12px;margin: 4px 2px;font-size: {int(sidelen)}px;min-width: 60px;}}
        QPushButton:hover {{background-color: {ColorList[4]};border: {int(sidelen * 0.0125)}px solid {ColorList[5]};}}
        QPushButton:pressed {{background-color: {ColorList[5]};}}
        QPushButton:focus {{outline: none;border: {int(sidelen * 0.0125)}px solid {ColorList[6]};}}"""
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

class SettingsWindow(QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent_mwin = parent
        self.settings = self._load_settings()
        if "hotkey" not in self.settings:
            self.settings["hotkey"] = HOTKEY
        self._setup_ui()
        settings_window_bg_style = f"""
        QMainWindow {{background-color: {ColorList[2]};color: {ColorList[1]};}}
        QWidget {{background-color: {ColorList[2]};color: {ColorList[1]};}}"""
        self.setStyleSheet(settings_window_bg_style)

    def _setup_ui(self): #设置窗口UI界面
        self.setWindowTitle("设置")
        self.setFixedSize(SideLen*10, SideLen*7)
        self.setWindowFlags(Qt.WindowType.Dialog | Qt.WindowType.WindowTitleHint |
                            Qt.WindowType.CustomizeWindowHint)
        self.setMaximumSize(int(SCR_WIDTH*0.8), int(SCR_HEIGHT*0.8))
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        # 主布局
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(sidelen, sidelen, sidelen, sidelen)
        main_layout.setSpacing(sidelen)
        # 分组布局
        Group_layout = QHBoxLayout()
        Group_layout.setSpacing(sidelen)
        Group_layout.addWidget(self._create_func_group()) # 功能组
        Group_layout.addWidget(self._create_appr_group()) # 外观组
        main_layout.addLayout(Group_layout)
        main_layout.addStretch(sidelen)
        main_layout.addWidget(self._create_button_area())

    def _create_func_group(self): # 创建功能组
        group_box = QGroupBox("功能设置")
        func_group_border_style = f"""
        QGroupBox {{font-weight: bold;font-size: {int(sidelen)}px;
            padding-top: {int(sidelen*0.5)}px;margin-top: {int(sidelen*0.5)}px;color: {ColorList[1]};
            background-color: {ColorList[2]};border: 2px solid {ColorList[5]};border-radius: {int(sidelen*0.5)}px;}}
        QGroupBox::title {{subcontrol-origin: margin;left: {int(sidelen*0.5)}px;
            padding: {int(-sidelen*0.5)} {int(sidelen*0.5)}px {int(-sidelen*0.5)} {int(sidelen*0.5)}px;color: {ColorList[1]};}}"""
        group_box.setStyleSheet(func_group_border_style)
        layout = QVBoxLayout(group_box)
        layout.setSpacing(int(sidelen))
        layout.setContentsMargins(sidelen, sidelen, sidelen, sidelen)
        group_box.setMaximumHeight(SideLen*7)
        layout.addWidget(self._set_auto_caps_lock_off())
        layout.addWidget(self._set_auto_collapse_Mwin())
        layout.addWidget(self._set_win_key_response())
        layout.addWidget(self._set_grid_snap_setting())
        layout.addWidget(self._set_custom_hotkey())
        return group_box
    def _create_appr_group(self): # 创建外观组
        group_box = QGroupBox("外观设置")
        appr_group_border_style = f"""
       QGroupBox {{font-weight: bold;font-size: {int(sidelen)}px;
            padding-top: {int(sidelen*0.5)}px;margin-top: {int(sidelen*0.5)}px;color: {ColorList[1]};
            background-color: {ColorList[2]};border: 2px solid {ColorList[5]};border-radius: {int(sidelen*0.5)}px;}}
        QGroupBox::title {{subcontrol-origin: margin;left: {int(sidelen*0.5)}px;
            padding: {int(-sidelen*0.5)} {int(sidelen*0.5)}px {int(-sidelen*0.5)} {int(sidelen*0.5)}px;color: {ColorList[1]};}}"""
        group_box.setStyleSheet(appr_group_border_style)
        layout = QVBoxLayout(group_box)
        layout.setSpacing(int(sidelen))
        layout.setContentsMargins(sidelen, sidelen, sidelen, sidelen)
        group_box.setMaximumHeight(SideLen * 7)
        layout.addWidget(self._set_opacity_level())
        layout.addWidget(self._set_example_combo_setting())
        layout.addWidget(self._set_ui_size_setting())
        layout.addWidget(self._set_example_slider_setting())
        layout.addStretch(1)
        return group_box

    def _set_auto_caps_lock_off(self): #创建自动关闭大写锁定设置
        widget = QWidget()
        layout = QHBoxLayout(widget)
        auto_caps_widget_style = f"QWidget {{ background-color: {ColorList[3]}; padding: {int(sidelen * 0.5)}px; }}"
        widget.setStyleSheet(auto_caps_widget_style)
        layout.setContentsMargins(0, 0, 0, 0)
        self.auto_caps_checkbox = QCheckBox("自动关闭大写锁定")
        self.auto_caps_checkbox.setChecked(self.settings.get("auto_caps", AUTO_CAPS))
        self.auto_caps_checkbox.stateChanged.connect(lambda state: print(f"自动关闭大写锁定: {'启用' if state else '禁用'}"))
        checkbox_full_style = f"""
                QCheckBox {{font-size: {int(sidelen)}px;color: {ColorList[1]};}}
                QCheckBox::indicator {{width: {int(sidelen)}px;height: {int(sidelen)}px;}}
                QCheckBox::indicator:unchecked {{border: 1px solid {ColorList[6]};background-color: {ColorList[2]};}}
                QCheckBox::indicator:checked {{border: 1px solid {ColorList[6]};background-color: {ColorList[1]};}}"""
        self.auto_caps_checkbox.setStyleSheet(checkbox_full_style)
        layout.addWidget(self.auto_caps_checkbox)
        tip_label = QLabel("?")
        tip_label.setToolTip("在使用快捷键后会自动关闭大写锁定")
        tip_label.setStyleSheet(" QLabel { color: #666666; font-weight: bold; padding: 2px 6px; }")
        layout.addWidget(tip_label)
        layout.addStretch(1)
        return widget
    def _set_auto_collapse_Mwin(self): #自动收起主窗口
        widget = QWidget()
        layout = QHBoxLayout(widget)
        auto_hide_widget_style = f"QWidget {{ background-color: {ColorList[3]}; padding: {int(sidelen * 0.5)}px; }}"
        widget.setStyleSheet(auto_hide_widget_style)
        layout.setContentsMargins(0, 0, 0, 0)
        self.auto_hide_checkbox = QCheckBox("自动收起主窗口")
        self.auto_hide_checkbox.setChecked(self.settings.get("auto_hide", AUTO_HIDE))
        self.auto_hide_checkbox.stateChanged.connect(lambda state: print(f"自动隐藏: {'启用' if state else '禁用'}"))
        checkbox_full_style = f"""
            QCheckBox {{font-size: {int(sidelen)}px;color: {ColorList[1]};}}
            QCheckBox::indicator {{width: {int(sidelen)}px;height: {int(sidelen)}px;}}
            QCheckBox::indicator:unchecked {{border: 1px solid {ColorList[6]};background-color: {ColorList[2]};}}
            QCheckBox::indicator:checked {{border: 1px solid {ColorList[6]};background-color: {ColorList[1]};}}"""
        self.auto_hide_checkbox.setStyleSheet(checkbox_full_style)
        layout.addWidget(self.auto_hide_checkbox)
        tip_label = QLabel("?")
        tip_label.setToolTip("点击窗口中的任意一个按钮后会自动隐藏主窗口")
        tip_label.setStyleSheet(" QLabel { color: #666666; font-weight: bold; padding: 2px 6px; }")
        layout.addWidget(tip_label)
        layout.addStretch(1)
        return widget
    def _set_win_key_response(self):  # Win键响应设置 #增
        widget = QWidget()
        layout = QHBoxLayout(widget)
        win_key_widget_style = f"QWidget {{ background-color: {ColorList[3]}; padding: {int(sidelen * 0.5)}px; }}"
        widget.setStyleSheet(win_key_widget_style)
        layout.setContentsMargins(0, 0, 0, 0)
        self.win_key_checkbox = QCheckBox("响应Win键")
        self.win_key_checkbox.setChecked(self.settings.get("win_key_enabled", False))
        self.win_key_checkbox.stateChanged.connect(lambda state: print(f"Win键响应: {'启用' if state else '禁用'}"))
        checkbox_full_style = f"""
            QCheckBox {{font-size: {int(sidelen)}px;color: {ColorList[1]};}}
            QCheckBox::indicator {{width: {int(sidelen)}px;height: {int(sidelen)}px;}}
            QCheckBox::indicator:unchecked {{border: 1px solid {ColorList[6]};background-color: {ColorList[2]};}}
            QCheckBox::indicator:checked {{border: 1px solid {ColorList[6]};background-color: {ColorList[1]};}}"""
        self.win_key_checkbox.setStyleSheet(checkbox_full_style)
        layout.addWidget(self.win_key_checkbox)
        tip_label = QLabel("?")
        tip_label.setToolTip("按下Win键也会显示/隐藏主窗口")
        tip_label.setStyleSheet(" QLabel { color: #666666; font-weight: bold; padding: 2px 6px; }")
        layout.addWidget(tip_label)
        layout.addStretch(1)
        return widget
    def _set_grid_snap_setting(self):  # 网格吸附设置
        widget = QWidget()
        layout = QHBoxLayout(widget)
        grid_snap_widget_style = f"QWidget {{ background-color: {ColorList[3]}; padding: {int(sidelen * 0.5)}px; }}"
        widget.setStyleSheet(grid_snap_widget_style)
        layout.setContentsMargins(0, 0, 0, 0)
        self.grid_snap_checkbox = QCheckBox("启用网格吸附")
        self.grid_snap_checkbox.setChecked(self.settings.get("grid_snap", GRID_SNAP_ENABLED))
        self.grid_snap_checkbox.stateChanged.connect(lambda state: print(f"网格吸附: {'启用' if state else '禁用'}"))
        checkbox_full_style = f"""
            QCheckBox {{font-size: {int(sidelen)}px;color: {ColorList[1]};}}
            QCheckBox::indicator {{width: {int(sidelen)}px;height: {int(sidelen)}px;}}
            QCheckBox::indicator:unchecked {{border: 1px solid {ColorList[6]};background-color: {ColorList[2]};}}
            QCheckBox::indicator:checked {{border: 1px solid {ColorList[6]};background-color: {ColorList[1]};}}"""
        self.grid_snap_checkbox.setStyleSheet(checkbox_full_style)
        layout.addWidget(self.grid_snap_checkbox)
        tip_label = QLabel("?")
        tip_label.setToolTip("子窗口组件会自动对齐到网格")
        tip_label.setStyleSheet(" QLabel { color: #666666; font-weight: bold; padding: 2px 6px; }")
        layout.addWidget(tip_label)
        layout.addStretch(1)
        return widget
    def _set_custom_hotkey(self): #设置快捷键
        # 修饰键映射字典
        self.modifier_keys = {
            "Ctrl": "<ctrl>",
            "Alt": "<alt>",
            "Shift": "<shift>",
        }
        # 普通键映射字典
        self.normal_keys = {
            "F1": "<f1>", "F2": "<f2>", "F3": "<f3>", "F5": "<f5>",
            "F6": "<f6>", "F7": "<f7>", "F8": "<f8>", "F9": "<f9>", "F10": "<f10>",
            "F11": "<f11>", "F12": "<f12>",
            "A": "a", "B": "b", "C": "c", "D": "d", "E": "e", "F": "f", "G": "g",
            "H": "h", "I": "i", "J": "j", "K": "k", "L": "l", "M": "m", "N": "n",
            "O": "o", "P": "p", "Q": "q", "R": "r", "S": "s", "T": "t", "U": "u",
            "V": "v", "W": "w", "X": "x", "Y": "y", "Z": "z",
            "0": "0", "1": "1", "2": "2", "3": "3", "4": "4", "5": "5", "6": "6",
            "7": "7", "8": "8", "9": "9",
            "空格": "<space>", "回车": "<enter>", "退格": "<backspace>",
            "插入": "<insert>", "Home": "<home>", "Page Up": "<page_up>",
            "Page Down": "<page_down>", "End": "<end>",
            "左箭头": "<left>", "右箭头": "<right>", "上箭头": "<up>", "下箭头": "<down>",
            "Caps Lock": "<caps_lock>", "Scroll Lock": "<scroll_lock>",
            "Num Lock": "<num_lock>", "Pause": "<pause>",
            "Print Screen": "<print_screen>", "菜单键": "<menu>",
            "小键盘0": "<num_0>", "小键盘1": "<num_1>", "小键盘2": "<num_2>",
            "小键盘3": "<num_3>", "小键盘4": "<num_4>", "小键盘5": "<num_5>",
            "小键盘6": "<num_6>", "小键盘7": "<num_7>", "小键盘8": "<num_8>",
            "小键盘9": "<num_9>",
            "小键盘*": "<num_multiply>", "小键盘+": "<num_add>",
            "小键盘-": "<num_subtract>", "小键盘.": "<num_decimal>",
            "小键盘/": "<num_divide>", "小键盘回车": "<num_enter>"
        }
        def _show_modifier_menu(button):  # 显示修饰键菜单
            menu = QMenu(self)
            menu.setWindowOpacity(0.8)
            modifier_menu_style = f"""
            QMenu {{background-color: {ColorList[2]};border: 1px solid {ColorList[5]};}}
            QMenu::item {{padding: 8px 16px;color: {ColorList[1]};}}
            QMenu::item:selected {{background-color: {ColorList[4]};}}"""
            menu.setStyleSheet(modifier_menu_style)
            for display_name, key_value in self.modifier_keys.items():
                action = menu.addAction(display_name)
                action.triggered.connect(
                    lambda checked, name=display_name,
                           btn=button: _update_modifier_button(name, btn))
            btn_pos = button.mapToGlobal(button.rect().bottomLeft())
            menu.exec(btn_pos)

        def _show_normal_key_menu(button):  # 显示普通键菜单
            menu = QMenu(self)
            menu.setWindowOpacity(0.8)
            normal_key_menu_style = f"""
            QMenu {{background-color: {ColorList[2]};border: 1px solid {ColorList[5]};}}
            QMenu::item {{padding: 8px 16px;color: {ColorList[1]};}}
            QMenu::item:selected {{background-color: {ColorList[4]};}}"""
            menu.setStyleSheet(normal_key_menu_style)
            categories = {
                "功能键": ["F1", "F2", "F3", "F5", "F6", "F7", "F8", "F9", "F10", "F11", "F12"],
                "字母键": ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M",
                           "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"],
                "数字键": ["0", "1", "2", "3", "4", "5", "6", "7", "8", "9"],
                "修饰键": ["Space", "Enter", "Tab", "Backspace", "Insert", "Home", "Page Up", "Page Down", "End"],
                "方向键": ["左箭头", "右箭头", "上箭头", "下箭头"],
                "锁定键": ["Caps Lock", "Scroll Lock", "Num Lock", "Pause"],
                "其他键": ["Print Screen", "菜单键"],
                "小键盘": ["小键盘0", "小键盘1", "小键盘2", "小键盘3", "小键盘4", "小键盘5",
                           "小键盘6", "小键盘7", "小键盘8", "小键盘9", "小键盘*", "小键盘+",
                           "小键盘-", "小键盘.", "小键盘/", "小键盘回车"]}
            for category, keys in categories.items():
                submenu = menu.addMenu(category)
                submenu_style = f"""
                QMenu {{background-color: {ColorList[2]};border: 1px solid {ColorList[5]};}}
                QMenu::item {{padding: 8px 16px;color: {ColorList[1]};}}
                QMenu::item:selected {{background-color: {ColorList[4]};}}"""
                submenu.setStyleSheet(submenu_style)
                for key_name in keys:
                    action = submenu.addAction(key_name)
                    action.triggered.connect(
                        lambda checked, name=key_name,
                               btn=button: _update_normal_key_button(name, btn))
            btn_pos = button.mapToGlobal(button.rect().bottomLeft())
            menu.exec(btn_pos)

        def _update_modifier_button(display_name, button):  # 更新修饰键按钮
            button.setText(display_name)
            _update_hotkey_display()

        def _update_normal_key_button(display_name, button):  # 更新普通键按钮
            button.setText(display_name)
            _update_hotkey_display()

        def _update_hotkey_display():  # 更新快捷键显示
            modifier_display = self.modifier_button.text()
            normal_key_display = self.normal_key_button.text()
            if modifier_display != "选择修饰键" and normal_key_display != "选择按键":
                modifier_value = self.modifier_keys.get(modifier_display, "")
                normal_value = self.normal_keys.get(normal_key_display, "")
                if modifier_value and normal_value:
                    self.current_hotkey = f"{modifier_value}+{normal_value}"

        def _parse_current_hotkey():  # 解析当前快捷键
            current_hotkey = self.settings.get("hotkey", HOTKEY)
            clean_hotkey = current_hotkey.strip()
            parts = clean_hotkey.replace('<', '').replace('>', '').split('+')
            # 查找修饰键
            modifier_display = "选择修饰键"
            if len(parts) >= 2:
                modifier_part = parts[0].lower()
                for display, value in self.modifier_keys.items():
                    clean_value = value.replace('<', '').replace('>', '')  # 改
                    if modifier_part == clean_value:
                        modifier_display = display
                        break
            # 查找普通键
            normal_key_display = "选择按键"
            if parts:
                normal_key_part = parts[-1].lower()
                for display, value in self.normal_keys.items():
                    clean_value = value.replace('<', '').replace('>', '')  # 改
                    if normal_key_part == clean_value:
                        normal_key_display = display
                        break
            return modifier_display, normal_key_display

        widget = QWidget()
        layout = QHBoxLayout(widget)
        hotkey_widget_style = f"QWidget {{ background-color: {ColorList[3]}; padding: {int(sidelen * 0.5)}px; }}"
        widget.setStyleSheet(hotkey_widget_style)
        layout.setContentsMargins(0, 0, 0, 0)
        hotkey_label = QLabel("快捷键:")
        hotkey_label_style = f"color: {ColorList[1]}; font-size: {int(sidelen)}px;"
        hotkey_label.setStyleSheet(hotkey_label_style)
        layout.addWidget(hotkey_label)
        # 修饰键按钮
        hotkey_button_style = f"""
        QPushButton {{background-color: {ColorList[3]};color: {ColorList[1]};border: 1px solid {ColorList[5]};
            ;padding: {int(sidelen*0.25)}px {int(sidelen*0.5)}px;font-size: {int(sidelen*0.8)}px;}}
        QPushButton:hover {{background-color: {ColorList[4]};}}
        QPushButton:pressed {{background-color: {ColorList[5]};}}"""
        self.modifier_button = QPushButton()
        modifier_display, normal_display = _parse_current_hotkey()
        self.modifier_button.setText(modifier_display)
        self.modifier_button.clicked.connect(lambda: _show_modifier_menu(self.modifier_button))
        self.modifier_button.setStyleSheet(hotkey_button_style)
        layout.addWidget(self.modifier_button)
        # 加号标签
        plus_label = QLabel("+")
        plus_label.setStyleSheet(f"color: {ColorList[1]}; font-size: {int(sidelen)}px;")
        layout.addWidget(plus_label)
        # 普通键按钮
        self.normal_key_button = QPushButton()
        self.normal_key_button.setText(normal_display)
        self.normal_key_button.clicked.connect(lambda: _show_normal_key_menu(self.normal_key_button))
        self.normal_key_button.setStyleSheet(hotkey_button_style)
        layout.addWidget(self.normal_key_button)
        # 初始化当前快捷键
        _update_hotkey_display()
        tip_label = QLabel("?")
        tip_label.setToolTip("设置显示/隐藏主窗口的全局快捷键\n点击按钮选择修饰键和按键")
        tip_label.setStyleSheet(" QLabel { color: #666666; font-weight: bold; padding: 2px 6px; }")
        layout.addWidget(tip_label)
        layout.addStretch(1)
        return widget

    def _set_opacity_level(self): #透明度设置
        widget = QWidget()
        layout = QHBoxLayout(widget)
        opacity_widget_style = f"QWidget {{ background-color: {ColorList[3]}; padding: {int(sidelen*0.5)}px; }}"
        widget.setStyleSheet(opacity_widget_style)
        layout.setContentsMargins(0, 0, 0, 0)
        opacity_label = QLabel("不透明度:")
        opacity_label_full_style = f"color: {ColorList[1]}; font-size: {int(sidelen)}px;"
        opacity_label.setStyleSheet(opacity_label_full_style)
        layout.addWidget(opacity_label)
        self.opacity_spinbox = QSpinBox()
        self.opacity_spinbox.setRange(10, 100)
        self.opacity_spinbox.setSuffix("%")
        default_opacity = int(OPACITY * 100)
        current_opacity = self.settings.get("opacity", OPACITY)
        self.opacity_spinbox.setValue(int(current_opacity * 100))
        self.opacity_spinbox.valueChanged.connect(lambda value: print(f"透明度设置为: {value}%"))
        spinbox_style = f"""
        QSpinBox {{background-color: {ColorList[3]};color: {ColorList[1]};border: 1px solid {ColorList[5]};
            border-radius: 3px;padding: 2px 4px;font-size: {int(sidelen*0.8)}px;}}
        QSpinBox::up-button, QSpinBox::down-button {{background-color: {ColorList[4]};border: 1px solid {ColorList[5]};}}
        QSpinBox::up-button:hover, QSpinBox::down-button:hover {{background-color: {ColorList[5]};}}"""
        self.opacity_spinbox.setStyleSheet(spinbox_style)
        layout.addWidget(self.opacity_spinbox)
        tip_label = QLabel("?")
        tip_label.setToolTip("调整主窗口的透明度，范围10%-100%")
        tip_label.setStyleSheet(" QLabel { color: #666666; font-weight: bold; padding: 2px 6px; }")
        layout.addWidget(tip_label)
        layout.addStretch(1)
        return widget
    def _set_example_combo_setting(self): # 主题下拉框设置
        widget = QWidget()
        layout = QHBoxLayout(widget)
        combo_widget_style = f"QWidget {{ background-color: {ColorList[3]}; padding: {int(sidelen*0.5)}px; }}"
        widget.setStyleSheet(combo_widget_style)
        layout.setContentsMargins(0, 0, 0, 0)
        combo_label = QLabel("主题:")
        combo_label_style = f"color: {ColorList[1]}; font-size: {int(sidelen)}px;"
        combo_label.setStyleSheet(combo_label_style)
        layout.addWidget(combo_label)
        self.theme_combo = QComboBox()
        self.theme_combo.addItems(["深色主题", "浅色主题"])
        current_theme = "深色主题" if ColorList == ColorList_Dk else "浅色主题"
        self.theme_combo.setCurrentText(current_theme)
        self.theme_combo.currentTextChanged.connect(lambda text: print(f"主题选择: {text}"))
        combo_box_style = f"""
        QComboBox {{background-color: {ColorList[3]};color: {ColorList[1]};border: 1px solid {ColorList[5]};
            padding: {int(sidelen*0.25)}px {int(sidelen)}px;font-size: {int(sidelen*0.8)}px;}}
        QComboBox::drop-down {{border: none;background-color: {ColorList[4]};}}
        QComboBox::down-arrow {{color: {ColorList[1]};}}
        QComboBox QAbstractItemView {{background-color: {ColorList[3]};color: {ColorList[1]};
            border: 1px solid {ColorList[5]};selection-background-color: {ColorList[4]};}}"""
        self.theme_combo.setStyleSheet(combo_box_style)
        layout.addWidget(self.theme_combo)
        tip_label = QLabel("?")
        tip_label.setToolTip("选择应用程序的主题样式\n重启生效")
        tip_label.setStyleSheet(" QLabel { color: #666666; font-weight: bold; padding: 2px 6px; }")
        layout.addWidget(tip_label)
        layout.addStretch(1)
        return widget
    def _set_ui_size_setting(self):  # UI尺寸设置
        widget = QWidget()
        layout = QHBoxLayout(widget)
        ui_size_widget_style = f"QWidget {{ background-color: {ColorList[3]}; padding: {int(sidelen * 0.5)}px; }}"
        widget.setStyleSheet(ui_size_widget_style)
        layout.setContentsMargins(0, 0, 0, 0)
        ui_size_label = QLabel("界面尺寸:")
        ui_size_label_style = f"color: {ColorList[1]}; font-size: {int(sidelen)}px;"
        ui_size_label.setStyleSheet(ui_size_label_style)
        layout.addWidget(ui_size_label)
        self.ui_size_combo = QComboBox()
        self.ui_size_combo.addItems(["大", "中", "小"])
        current_ui_size = self.settings.get("ui_size", "中")
        self.ui_size_combo.setCurrentText(current_ui_size)
        self.ui_size_combo.currentTextChanged.connect(lambda text: print(f"界面尺寸选择: {text}"))
        combo_box_style = f"""
        QComboBox {{background-color: {ColorList[3]};color: {ColorList[1]};border: 1px solid {ColorList[5]};
            padding: {int(sidelen*0.25)}px {int(sidelen)}px;font-size: {int(sidelen*0.8)}px;}}
        QComboBox::drop-down {{border: none;background-color: {ColorList[4]};}}
        QComboBox::down-arrow {{color: {ColorList[1]};}}
        QComboBox QAbstractItemView {{background-color: {ColorList[3]};color: {ColorList[1]};
            border: 1px solid {ColorList[5]};selection-background-color: {ColorList[4]};}}"""
        self.ui_size_combo.setStyleSheet(combo_box_style)
        layout.addWidget(self.ui_size_combo)
        tip_label = QLabel("?")
        tip_label.setToolTip("设置界面元素的尺寸大小\n重启后生效")
        tip_label.setStyleSheet(" QLabel { color: #666666; font-weight: bold; padding: 2px 6px; }")
        layout.addWidget(tip_label)
        layout.addStretch(1)
        return widget
    def _set_example_slider_setting(self): # 创建主窗口大小设置
        widget = QWidget()
        layout = QHBoxLayout(widget)
        slider_widget_style = f"QWidget {{ background-color: {ColorList[3]}; padding: {int(sidelen*0.5)}px; }}"
        widget.setStyleSheet(slider_widget_style)
        layout.setContentsMargins(0, 0, 0, 0)
        slider_label = QLabel("窗口大小:")
        slider_label_style = f"color: {ColorList[1]}; font-size: {int(sidelen)}px;"
        slider_label.setStyleSheet(slider_label_style)
        layout.addWidget(slider_label)
        self.win_ratio_slider = QSlider(Qt.Orientation.Horizontal)
        self.win_ratio_slider.setRange(2, 9)
        self.win_ratio_slider.setValue(int(WIN_RATIO * 10))
        self.win_ratio_slider.valueChanged.connect(lambda value: print(f"窗口占比: {value}0%"))
        slider_style = f"""
        QSlider::groove:horizontal {{border: 2px solid {ColorList[5]};height: {int(sidelen*0.5)}px;
            ;background-color: {ColorList[3]};border-radius: {int(sidelen*0.125)}px;}}
        QSlider::handle:horizontal {{
            background-color: {ColorList[1]};border: 2px solid {ColorList[5]};
            width: {int(sidelen)}px;margin: {int(-sidelen*0.25)}px 0;border-radius: {int(sidelen*0.25)}px;}}
        QSlider::handle:horizontal:hover {{background-color: {ColorList[6]};}}"""
        self.win_ratio_slider.setStyleSheet(slider_style)
        self.win_ratio_slider.setFixedWidth(int(sidelen * 6))
        layout.addWidget(self.win_ratio_slider, 1)
        value_label = QLabel(f"{int(WIN_RATIO * 100)}%")
        value_label.setFixedWidth(20)
        self.win_ratio_slider.valueChanged.connect(lambda value: value_label.setText(f"{value}0%"))
        value_label_style = f"color: {ColorList[1]}; background-color: transparent; font-size: {int(sidelen)}px;"
        value_label.setStyleSheet(value_label_style)
        layout.addWidget(value_label)
        tip_label = QLabel("?")
        tip_label.setToolTip("设置主窗口相对于屏幕的尺寸比例\n重启后生效")
        tip_label.setStyleSheet(" QLabel { color: #666666; font-weight: bold; padding: 2px 6px; }")
        layout.addWidget(tip_label)
        layout.addStretch(1)
        return widget

    def _create_button_area(self): # 创建按钮区域
        widget = QWidget()
        widget.setFixedHeight(int(sidelen*3))
        layout = QHBoxLayout(widget)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(1)
        settings_button_style = f"""
        QPushButton {{background-color: {ColorList[4]};color: {ColorList[1]};border: 1px solid {ColorList[5]};
            border-radius: 3px;padding: {int(sidelen*0.25)}px {int(sidelen*0.5)}px;
            font-size: {int(sidelen)}px;min-width: {int(sidelen*5)}px;max-width: {int(sidelen*1.5)}px;min-height: {int(sidelen*1.5)}px;max-height: {int(sidelen*1.5)}px;}}
        QPushButton:hover {{background-color: {ColorList[5]};}}
        QPushButton:pressed {{background-color: {ColorList[6]};}}"""
        self.apply_restart_btn = QPushButton("应用并重启")
        self.apply_restart_btn.clicked.connect(self._apply_and_restart)
        self.apply_restart_btn.setStyleSheet(settings_button_style)
        layout.addWidget(self.apply_restart_btn)
        self.apply_btn = QPushButton("应用")
        self.apply_btn.clicked.connect(self._apply_settings)
        self.apply_btn.setStyleSheet(settings_button_style)
        layout.addWidget(self.apply_btn)
        self.cancel_btn = QPushButton("取消")
        self.cancel_btn.clicked.connect(self.close)
        self.cancel_btn.setStyleSheet(settings_button_style)
        layout.addWidget(self.cancel_btn)
        self.ok_btn = QPushButton("确定")
        self.ok_btn.clicked.connect(self._ok_settings)
        self.ok_btn.setStyleSheet(settings_button_style)
        layout.addWidget(self.ok_btn)
        return widget
    def _load_settings(self): # 从配置文件加载设置
        try:
            config_path = "config.json"
            if os.path.exists(config_path):
                with open(config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                return config.get("settings", {})
        except Exception as e:
            print(f"[ERROR] 加载设置失败: {e}")
        return {}
    def _save_settings(self): # 保存设置到配置文件
        try:
            config_path = "config.json"
            config = {}
            if os.path.exists(config_path):
                with open(config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
            # 获取当前设置的快捷键
            current_hotkey = getattr(self, 'current_hotkey', None)
            if not current_hotkey:
                # 如果没有设置新的快捷键，使用原来的或默认值
                current_hotkey = self.settings.get("hotkey", HOTKEY)
            config["settings"] = {
                "auto_caps": self.auto_caps_checkbox.isChecked(),
                "auto_hide": self.auto_hide_checkbox.isChecked(),
                "win_key_enabled": self.win_key_checkbox.isChecked(),
                "grid_snap": self.grid_snap_checkbox.isChecked(),
                "opacity": self.opacity_spinbox.value() / 100.0,
                "win_ratio": self.win_ratio_slider.value() / 10.0,
                "hotkey": current_hotkey,
                "theme": self.theme_combo.currentText(),
                "ui_size": self.ui_size_combo.currentText(),
            }
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=4)
            return True
        except Exception as e:
            print(f"[ERROR] 保存设置失败: {e}")
            return False
    def _apply_settings(self): # 应用设置
        if self._save_settings():
            # 更新全局变量
            global AUTO_CAPS, AUTO_HIDE, HOTKEY, OPACITY, WIN_RATIO, GRID_SNAP_ENABLED
            AUTO_CAPS = self.auto_caps_checkbox.isChecked()
            AUTO_HIDE = self.auto_hide_checkbox.isChecked()
            WIN_KEY_ENABLED = self.win_key_checkbox.isChecked()
            GRID_SNAP_ENABLED = self.grid_snap_checkbox.isChecked()
            current_hotkey = getattr(self, 'current_hotkey', None)
            if current_hotkey:
                HOTKEY = current_hotkey
            OPACITY = self.opacity_spinbox.value() / 100.0
            WIN_RATIO = self.win_ratio_slider.value() / 10.0
            if self.theme_combo.currentText() == "浅色主题":
                ColorList = ColorList_Lt
            else:
                ColorList = ColorList_Dk
            if self.parent_mwin:
                self.parent_mwin.setWindowOpacity(OPACITY)
                if current_hotkey:
                    restart_hotkey_listener(self.parent_mwin)
            print("[INFO] 设置已应用")

    def _ok_settings(self):
        self._apply_settings()
        self.close()
    def _apply_and_restart(self):
        try:
            self._save_settings()
            # 停止热键监听器并等待线程死亡
            if hasattr(self.parent_mwin, 'hotkeys_listener') and self.parent_mwin.hotkeys_listener:
                self.parent_mwin.hotkeys_listener.stop_listening()
                listener = self.parent_mwin.hotkeys_listener
                if hasattr(listener, 'listener_thread') and listener.listener_thread:
                    listener.listener_thread.join(timeout=3.0)
            QApplication.closeAllWindows()
            # 释放互斥锁
            global _app_mutex
            if _app_mutex:
                win32api.CloseHandle(_app_mutex)
                _app_mutex = None
            import os
            import sys
            if getattr(sys, 'frozen', False):
                exe_path = sys.executable
                os.startfile(exe_path)
            else:
                import subprocess
                pythonw = sys.executable.replace("python.exe", "pythonw.exe")
                if not os.path.exists(pythonw):
                    pythonw = sys.executable
                subprocess.Popen([pythonw, "-m", "WinDrawer"],
                                 cwd=os.path.dirname(sys.argv[0]) or os.getcwd(),
                                 creationflags=subprocess.CREATE_NO_WINDOW)
            QTimer.singleShot(400, lambda: os._exit(0))
        except Exception as e:
            print(f"[RESTART] 重启异常: {e}")
            QTimer.singleShot(100, lambda: os._exit(0))
def show_settings_dialog(Mwin):
    try:
        settings_window = SettingsWindow(Mwin)
        screen_geometry = QApplication.primaryScreen().availableGeometry()
        window_geometry = settings_window.frameGeometry()
        # 计算居中位置
        x = max(0, (screen_geometry.width() - window_geometry.width()) // 2)
        y = max(0, (screen_geometry.height() - window_geometry.height()) // 2)
        settings_window.move(x, y)
        settings_window.show()
    except Exception as e:
        print(f"[ERROR] 显示设置对话框失败: {e}")
def _setup_icon_cache(): # 初始化图标缓存
    global ICON_CACHE
    try:
        theme_suffix = "Dk" if ColorList == ColorList_Lt else "Lt"
        icon_names = [
            "AddWindow", "CollapseWindow", "Settings", "ReloadConfig",
            "Exit", "AddFile", "FileName", "OpenFolder"
        ]
        for icon_name in icon_names:
            resource_path = f"icons/{theme_suffix}/data/QPixmap/{theme_suffix}/{icon_name}_{theme_suffix}.png"
            icon = ResourceLoader.get_icon(resource_path)
            if not icon.isNull():
                ICON_CACHE[icon_name] = icon
            else:
                print(f"[WARNING] 图标资源不存在: {resource_path}")
    except Exception as e:
        print(f"[ERROR] 初始化图标缓存失败: {e}")
def get_file_icon(file_path):  # 使用获取文件图标
    try:
        try:
            import pythoncom
            pythoncom.CoInitialize()  # 初始化 COM
        except:
            pass
        original_icon = QIcon()
        if file_path.lower().endswith('.lnk'):
            import pythoncom
            from win32com.shell import shell, shellcon
            shortcut = pythoncom.CoCreateInstance(
                shell.CLSID_ShellLink,
                None,
                pythoncom.CLSCTX_INPROC_SERVER,
                shell.IID_IShellLink
            )
            shortcut.QueryInterface(pythoncom.IID_IPersistFile).Load(file_path)
            # 获取目标路径
            target_path = shortcut.GetPath(0)[0]
            # 如果目标文件是.bat文件，直接获取快捷方式图标
            if target_path.lower().endswith('.bat'):
                file_info = QFileInfo(file_path)
            else:
                file_info = QFileInfo(target_path)
            icon_provider = QFileIconProvider()
            original_icon = icon_provider.icon(file_info)
        elif file_path.lower().endswith('.url'):
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                # 查找 IconFile 属性
                for line in content.split('\n'):
                    if line.lower().startswith('iconfile='):
                        icon_path = line.split('=', 1)[1].strip()
                        if icon_path and Path(icon_path).exists():
                            original_icon = QIcon(icon_path)
                            break
                if original_icon.isNull():
                    file_info = QFileInfo(file_path)
                    icon_provider = QFileIconProvider()
                    original_icon = icon_provider.icon(file_info)
            except Exception as e:
                file_info = QFileInfo(file_path)
                icon_provider = QFileIconProvider()
                original_icon = icon_provider.icon(file_info)
                print(f"[ERROR] 获取URL图标失败 {file_path}: {e}")
        else:
            file_info = QFileInfo(file_path)
            icon_provider = QFileIconProvider()
            original_icon = icon_provider.icon(file_info)
        if not original_icon.isNull():
            icon_size = int(SideLen * 0.75)
            available_sizes = original_icon.availableSizes()
            final_pixmap = QPixmap(icon_size, icon_size)
            final_pixmap.fill(Qt.GlobalColor.transparent)
            painter = QPainter(final_pixmap)
            painter.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform)
            # 选择一个合适的源尺寸进行缩放
            if available_sizes:
                best_size = available_sizes[0]
                for size in available_sizes:
                    if abs(size.width() - icon_size) < abs(best_size.width() - icon_size):
                        best_size = size
                source_pixmap = original_icon.pixmap(best_size)
            else:
                source_pixmap = original_icon.pixmap(icon_size, icon_size)
            scaled_pixmap = source_pixmap.scaled(icon_size, icon_size, Qt.AspectRatioMode.KeepAspectRatio,
                                                 Qt.TransformationMode.SmoothTransformation)
            x_offset = (icon_size - scaled_pixmap.width()) // 2
            y_offset = (icon_size - scaled_pixmap.height()) // 2
            painter.drawPixmap(x_offset, y_offset, scaled_pixmap)
            painter.end()
            return QIcon(scaled_pixmap)
        else:
            return QIcon()
    except Exception as e:
        print(f"[ERROR] 获取文件图标失败 {file_path}: {e}")
        return QIcon()

def wrap_button_text(text):  # 优化按钮文本显示，支持换行
    # 检查是换行符号
    for symbol in WRAP_SYMBOLS:
        if symbol in text:
            parts = text.split(symbol, 1)
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
    global SCR_WIDTH, SCR_HEIGHT, MWIN_WIDTH, MWIN_HEIGHT,START_X, START_Y, HIDE_END_Y, GRID_INFO
    global sidelen, SideLen, MIN_SIZE, XYCtrlCTN, GRID_CELL_SIZE
    if SCR_WIDTH is None:
        ScrXY = QApplication.primaryScreen()
        ScrXYTrue = ScrXY.geometry()
        SCR_WIDTH = ScrXYTrue.width()  # 获取屏幕可用宽度
        SCR_HEIGHT = ScrXYTrue.height()  # 获取屏幕可用高度
        quantum_base = SCR_WIDTH * 0.025
        quotient = quantum_base // 8
        remainder = quantum_base % 8
        QUANTUM_SIZE = int(quotient * 8) if remainder < 4 else int((quotient + 1) * 8) # 量子尺寸
        try:# 读取界面尺寸设置，确定缩放倍数
            config_path = "config.json"
            if os.path.exists(config_path):
                with open(config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                ui_size_setting = config.get("settings", {}).get("ui_size", "中")  # 默认中
            else:ui_size_setting = "中"
        except:ui_size_setting = "中"
        if ui_size_setting == "大":SCALE_FACTOR = 3
        elif ui_size_setting == "中":SCALE_FACTOR = 2
        else:SCALE_FACTOR = 1
        MANIFEST_SIZE = QUANTUM_SIZE * SCALE_FACTOR# 计算显化尺寸
        sidelen = MANIFEST_SIZE // 4  # 间隔距离 #改
        SideLen = MANIFEST_SIZE  # 按钮边长（4倍显化尺寸） #改
        MIN_SIZE = sidelen * 2 + SideLen  # 最小尺寸限制 #改
        XYCtrlCTN = (MIN_SIZE*6, MIN_SIZE*6)  # 窗口默认大小 #改
        GRID_CELL_SIZE = MANIFEST_SIZE // 2  # 网格单元大小（2倍显化尺寸） #改
        initial_w= int(SCR_WIDTH * WIN_RATIO)  # 窗口宽度
        initial_h = int(SCR_HEIGHT * WIN_RATIO)  # 窗口高度
        inner_w = initial_w - 4*sidelen #内部可用宽度
        inner_h = initial_h - 4*sidelen
        blocks_x = ceil(inner_w / SideLen)  #方块数量
        blocks_y = ceil(inner_h / SideLen)
        MWIN_WIDTH =  blocks_x*SideLen + 4*sidelen #主窗口最终宽度
        MWIN_HEIGHT = blocks_y*SideLen + 4*sidelen
        START_X = (SCR_WIDTH - MWIN_WIDTH)//2
        START_Y = (-MWIN_HEIGHT + int(MWIN_HEIGHT*0))
        HIDE_END_Y = START_Y
        illegal_margins = (16, 16, 16, 32)  # 左、上、右、下禁区边界
        available_width = MWIN_WIDTH - (illegal_margins[0] + illegal_margins[2])
        available_height = MWIN_HEIGHT - (illegal_margins[1] + illegal_margins[3])
        # 计算网格单元数量
        grid_cols = available_width // GRID_CELL_SIZE
        grid_rows = available_height // GRID_CELL_SIZE
        # 计算实际网格区域尺寸（8px对齐）
        grid_area_width = grid_cols * GRID_CELL_SIZE
        grid_area_height = grid_rows * GRID_CELL_SIZE
        # 计算网格起始位置（居中）
        grid_start_x = illegal_margins[0] + (available_width - grid_area_width) // 2
        grid_start_y = illegal_margins[1] + (available_height - grid_area_height) // 2
        GRID_INFO = {
            'cell_size': GRID_CELL_SIZE,  # 单个网格单元的大小（8px）
            'grid_size': (grid_cols, grid_rows),  # 网格的列数和行数
            'start_pos': (grid_start_x, grid_start_y),  # 网格区域左上角的起始坐标
            'area_size': (grid_area_width, grid_area_height),  # 网格区域的总宽度和高度
            'legal_area': QRect(grid_start_x, grid_start_y, grid_area_width, grid_area_height),  # 合适网格区域的矩形
            'illegal_margins': illegal_margins  # 禁区边界：左、上、右、下边距
        }
        _setup_icon_cache()

def Child_win(win_key,win_config,parent):#利用配置文件创建子窗口，读取为临时变量并传入子窗口类中
    win_name_temp = win_config.get(f"{win_key}_N", "未命名")
    win_geo_temp = win_config.get(f"{win_key}_Geo", [10, 10, 364, 364])
    win_state_temp = win_config.get(f"{win_key}_State", True)
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
        win_state=win_state_temp
    )
    return sub_win
class MainWindow(QMainWindow):
    forward_fold_signal = pyqtSignal(str, object)
    def __init__(self):
        super().__init__()
    def closeEvent(self, event):  # 主窗口关闭事件
        try:
            # 停止热键监听
            if hasattr(self, 'hotkeys_listener') and self.hotkeys_listener:
                self.hotkeys_listener.stop_listening()
            # 清理子窗口
            mdi_area = self.centralWidget()
            if mdi_area:
                for sub_window in mdi_area.subWindowList():
                    if hasattr(sub_window, 'config_save_timer'):
                        sub_window.config_save_timer.stop()
            event.accept()
        except Exception as e:
            print(f"[ERROR] 关闭清理失败: {e}")
            event.accept()
def CR_Mwin():  # 主窗口
    DefMainWinSize()
    Mwin = MainWindow()
    Mwin.resize(MWIN_WIDTH, MWIN_HEIGHT)#窗口大小
    Mwin.move(START_X,START_Y)#窗口位置
    Mwin.setWindowOpacity(OPACITY)
    Mwin.setWindowFlags(Mwin.windowFlags() | Qt.WindowType.WindowStaysOnTopHint)#置顶
    Mwin.setAttribute(Qt.WidgetAttribute.WA_QuitOnClose)
    flags = (
            Qt.WindowType.FramelessWindowHint |
            Qt.WindowType.WindowStaysOnTopHint |
            Qt.WindowType.Tool)#窗口属性
    Mwin.setWindowFlags(flags)
    main_window_style = f"QMainWindow {{ background-color: {ColorList[2]}; }}"
    Mwin.setStyleSheet(main_window_style)
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
    add_tool_window_button(Mwin)
    try:#创建子窗口
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
                    sub_win.window_folded_signal.connect(Mwin.forward_fold_signal)
                    M_mdi_area.addSubWindow(sub_win)  # 添加到MDI区域
        else:  # 如果配置文件不存在，创建默认窗口
            window_0 = CtrlCTN(name='未命名', parent=M_mdi_area)
            sub_win.window_folded_signal.connect(Mwin.forward_fold_signal)
            M_mdi_area.addSubWindow(window_0)
            window_0.show()
    except Exception as e:
        print(f"[ERROR] 读取配置文件失败: {e}")# 异常处理：配置文件读取或解析失败时
        window_0 = CtrlCTN(name='未命名', parent=M_mdi_area)# 出错时创建默认窗口，确保程序不会崩溃
        sub_win.window_folded_signal.connect(Mwin.forward_fold_signal)
        M_mdi_area.addSubWindow(window_0)
        window_0.show()
    create_rest_prompt_window(Mwin)
    Mwin.forward_fold_signal = pyqtSignal(str, object)
    Mwin.show()
    return Mwin

def Mwin_ToggleState(Mwin): ##切换窗口状态
    def _toggle_caps_lock_if_on():  # 检查大写锁定状态并关闭
        try:
            if ctypes.windll.user32.GetKeyState(0x14) & 0x0001:
                ctypes.windll.user32.keybd_event(0x14, 0x45, 0x0001, 0)
                ctypes.windll.user32.keybd_event(0x14, 0x45, 0x0001 | 0x0002, 0)
        except Exception as e:
            print(f"[ERROR] 关闭大写锁定失败: {e}")
    def Anim_AppearMwin(Mwin):  # 窗口滑出动画
        Anim = QPropertyAnimation(Mwin, b"geometry", parent=Mwin)
        Anim.setDuration(500);
        Anim.setEasingCurve(QEasingCurve.Type.InOutQuart)  # 动画持续时间和方式
        MwinXY = Mwin.geometry()  # 窗口初始位置
        ENXY = QRect(START_X, APPEAR_END_Y, MWIN_WIDTH, MWIN_HEIGHT)  # 结束位置
        Anim.setStartValue(MwinXY);
        Anim.setEndValue(ENXY)
        if AUTO_CAPS:
            Anim.finished.connect(lambda: _toggle_caps_lock_if_on())
        Anim.start()
        return Anim
    def Anim_HideMwin(Mwin):
        Anim = QPropertyAnimation(Mwin, b"geometry", parent=Mwin)
        Anim.setDuration(500);
        Anim.setEasingCurve(QEasingCurve.Type.InOutQuart)  # 动画持续时间和方式
        MwinXY = Mwin.geometry()  # 窗口位置
        ENXY = QRect(START_X, HIDE_END_Y, MWIN_WIDTH, MWIN_HEIGHT)  # 结束位置
        Anim.setStartValue(MwinXY);
        Anim.setEndValue(ENXY)
        if AUTO_CAPS:
            Anim.finished.connect(lambda: _toggle_caps_lock_if_on())
        Anim.start()
        return Anim
    if hasattr(Mwin, 'hotkeys_listener') and Mwin.hotkeys_listener:
        if not Mwin.hotkeys_listener.TGL:  # 当前是隐藏状态
            Anim_AppearMwin(Mwin)
            Mwin.hotkeys_listener.TGL = True  # 更新状态
        else:  # 当前是显示状态
            Anim_HideMwin(Mwin)
            Mwin.hotkeys_listener.TGL = False  # 更新状态
def BD_kSC(Mwin):#快捷键触发的事件，隐藏动画或显示动画
    GL_Lsn = GlobalHotkeyListener(hotkey=HOTKEY)
    Mwin.hotkeys_listener = GL_Lsn
    def toggle_window():
        try:
            if not GL_Lsn.TGL:
                Mwin_ToggleState(Mwin)
                GL_Lsn.TGL = True
            else:
                Mwin_ToggleState(Mwin)
                GL_Lsn.TGL = False
        except Exception as e:
            print(f'[ERROR] 切换窗口状态失败: {e}')

    def toggle_window_win_key():  # Win键响应
        try:
            if not GL_Lsn.TGL:
                Mwin_ToggleState(Mwin)
                GL_Lsn.TGL = True
            else:
                Mwin_ToggleState(Mwin)
                GL_Lsn.TGL = False
        except Exception as e:
            print(f'[ERROR] Win键切换窗口状态失败: {e}')

    GL_Lsn.Evt.connect(toggle_window)
    GL_Lsn.WinKeyEvt.connect(toggle_window_win_key)
    GL_Lsn.start_listening()
def restart_hotkey_listener(Mwin):  # 重启快捷键监听
    try:
        if hasattr(Mwin, 'hotkeys_listener') and Mwin.hotkeys_listener:
            Mwin.hotkeys_listener.stop_listening()
        BD_kSC(Mwin)
    except Exception as e:
        print(f"[ERROR] 重启快捷键监听失败: {e}")

def add_tool_window_button(Mwin):  # 在主窗口添加功能菜单按钮
    class MenuBtn(QToolButton):  # 菜单按钮类
        def __init__(self, func_name,func,icon_name,parent=None):  # 初始化菜单按钮
            super().__init__(parent)
            self.func = func
            self.setFixedSize(SideLen, SideLen)
            self.clicked.connect(self._on_click)
            self.setText(func_name)
            self._setup_icon(icon_name)
            self.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextUnderIcon)
            menu_button_style = f"""
            QToolButton {{
                background-color: transparent;border: 1px solid transparent;border-radius: {int(sidelen * 0.5)}px;color: {ColorList[1]};
                font-size: {int(sidelen * 0.5)}px;padding: 2px;}}
            QToolButton:hover {{background-color: {ColorList[4]};}}
            QToolButton:pressed {{background-color: {ColorList[5]};}}"""
            self.setStyleSheet(menu_button_style)

        def _setup_icon(self, icon_name): # 设置图标函数
            try:
                icon = ICON_CACHE.get(icon_name)
                if icon and not icon.isNull():
                    self.setIcon(icon)
                    self.setIconSize(QSize(int(SideLen * 0.75), int(SideLen * 0.75)))
            except Exception as e:
                print(f"[ERROR] 加载图标失败 {icon_name}: {e}")

        def _on_click(self):  # 按钮点击事件处理
            try:
                self.func()
            except Exception as e:
                print(f"[ERROR] 执行功能时出错: {e}")
    # 创建功能菜单按钮
    functions = [
        ("新建窗口", lambda: CRChildWin(Mwin), "AddWindow"),
        ("刷新配置", lambda: Ref_json(Mwin), "ReloadConfig"),
        ("按钮名称", lambda: show_buttons_name(Mwin), "FileName"),
        ("添加文件", lambda: AddFile(Mwin), "AddFile"),
        ("打开文件夹", lambda: open_RepoFolder(), "OpenFolder"),
        ("设置", lambda: show_settings_dialog(Mwin), "Settings"),
        ("退出", lambda: QApplication.quit(), "Exit")
    ]#功能按钮
    folded_window_buttons = []#被折叠窗口的按钮
    menu_geometry_cache = None
    def create_folded_window_button(window_name, window_ref):  # 创建被折叠窗口的对应按钮
        def show_folded_window():
            window_ref.window_expand_signal.emit()
            for i, (name, ref, _) in enumerate(folded_window_buttons):
                if ref == window_ref:  # 通过引用匹配
                    folded_window_buttons.pop(i)
                    break
            create_all_buttons()
        button_tuple = (window_name, show_folded_window, "CollapseWindow")
        folded_window_buttons.append((window_name, window_ref, button_tuple))
        create_all_buttons()

    def update_all_buttons():  # 合并按钮列表
        try:
            all_buttons = functions.copy()  # 基础功能按钮
            # 添加折叠窗口按钮
            for _, _, button_tuple in folded_window_buttons:
                all_buttons.append(button_tuple)
            return all_buttons
        except Exception as e:
            print(f"[ERROR] 更新所有按钮列表失败: {e}")
            return functions.copy()

    def calculate_menu_geometry():  # 计算菜单几何属性
        try:
            all_buttons = update_all_buttons()
            menu_width = MWIN_WIDTH - 2 * sidelen
            max_cols = max(1, (menu_width - 2 * sidelen) // (SideLen + sidelen))
            rows = (len(all_buttons) + max_cols - 1) // max_cols
            menu_height = rows * (SideLen + sidelen) + 2 * sidelen
            btn_global_pos = func_btn.mapToGlobal(func_btn.rect().topLeft())
            menu_x = btn_global_pos.x() + (func_btn.width() - menu_width) // 2
            menu_y = btn_global_pos.y() - menu_height - sidelen // 2
            geometry = (menu_x, menu_y, menu_width, menu_height, max_cols)
            print(f"[DEBUG] 计算菜单几何属性: 位置({menu_x},{menu_y}) 大小({menu_width}x{menu_height})")
            return geometry
        except Exception as e:
            print(f"[ERROR] 计算菜单几何属性失败: {e}")
            return None

    def create_all_buttons():  # 移除现有按钮并创建新的按钮
        try:
            for i in reversed(range(layout.count())):
                item = layout.itemAt(i)
                if item.widget():
                    item.widget().setParent(None)
            geometry = calculate_menu_geometry()
            if not geometry:
                return
            menu_x, menu_y, menu_width, menu_height, max_cols = geometry
            func_menu.setFixedSize(menu_width, menu_height)
            func_menu.move(menu_x, menu_y)
            all_buttons = update_all_buttons()
            for i, (name, func, icon_name) in enumerate(all_buttons):
                btn = MenuBtn(name, func, icon_name, func_menu)
                row = i // max_cols
                col = i % max_cols
                layout.addWidget(btn, row, col)
            # 更新几何属性缓存
            nonlocal menu_geometry_cache
            menu_geometry_cache = geometry
            print(f"[DEBUG] 创建所有按钮完成，缓存几何属性")
        except Exception as e:
            print(f"[ERROR] 创建所有按钮失败: {e}")

    def control_button_event():  # 控制按钮事件
        try:
            if func_menu.isVisible():
                func_menu.hide()
            else:
                if menu_geometry_cache:
                    _, _, menu_width, menu_height, _ = menu_geometry_cache
                    btn_global_pos = func_btn.mapToGlobal(func_btn.rect().topLeft())
                    menu_x = btn_global_pos.x() + (func_btn.width() - menu_width) // 2
                    menu_y = btn_global_pos.y() - menu_height - sidelen // 2
                    func_menu.move(menu_x, menu_y)
                    func_menu.show()
                    func_menu.raise_()
                else:
                    create_all_buttons()
                    func_menu.show()
                    func_menu.raise_()
        except Exception as e:
            print(f"[ERROR] 控制按钮事件失败: {e}")

    func_btn = QPushButton("⮝", Mwin)
    func_btn.setFixedSize(MWIN_WIDTH, int(sidelen))
    btn_x = 0
    btn_y = Mwin.height() - int(sidelen)
    func_btn.move(btn_x, btn_y)
    func_btn.show()
    # 创建功能菜单弹窗（QWidget实现）
    func_menu = QWidget(Mwin)
    func_menu.setWindowFlags(
        Qt.WindowType.FramelessWindowHint |
        Qt.WindowType.Popup |
        Qt.WindowType.WindowStaysOnTopHint)
    func_menu.setWindowOpacity(0.8)
    func_menu_style = (f"QWidget{{border: 2px solid {ColorList[6]}; background-color: {ColorList[2]}; }}")
    func_menu.setStyleSheet(func_menu_style)
    func_menu.hide()
    def update_button_symbol(is_expanded):
        func_btn.setText("⮟" if is_expanded else "⮝")
    func_menu.showEvent = lambda event:update_button_symbol(True)
    func_menu.hideEvent = lambda event:update_button_symbol(False)
    # 创建布局和按钮容器
    layout = QGridLayout(func_menu)
    layout.setSpacing(sidelen)
    layout.setContentsMargins(sidelen,sidelen,sidelen,sidelen)
    func_btn.clicked.connect(control_button_event)
    Mwin.forward_fold_signal.connect(create_folded_window_button)
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
        Mwin_ToggleState(Mwin)
        dialog = MessageDialog(
            parent=Mwin,editable=True,default_text="",modal=True,
            title="新建子窗口",message="请输入子窗口名称:",placeholder="窗口名称")
        result = dialog.exec()
        if result != QDialog.DialogCode.Accepted:  # 用户点击取消或关闭对话框
            Mwin_ToggleState(Mwin)
            return  # 直接返回
        Mwin_ToggleState(Mwin)
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
            f"{win_key}_State": True,
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
    # 处理两个列表，以list_FILE为基准，去除其中与list_JSON相同的部分，得到list_FILE_inc_JSON
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
            button_overlay_style = f"""
                QLabel {{
                    background-color: rgba(0, 0, 0, 180);color: {ColorList[1]};
                    border: 1px solid rgba(255, 255, 255, 100);border-radius: 4px;
                    font-size: {int(sidelen * 0.5)}px;  /* 使用sidelen控制字体大小 */
                    padding: {int(sidelen * 0.25)}px;font-weight: bold;}}"""
            overlay.setStyleSheet(button_overlay_style)
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
            auto_close=1000)
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

def create_rest_prompt_window(Mwin): # 创建响铃提示窗口
    time_display = [None]
    next_prompt_display = [None]
    rest_timer = QTimer()
    rest_timer.setSingleShot(True)

    def _snap_rest_window_to_grid(geometry):  # 时钟组件的网格吸附方法
        try:
            if not GRID_SNAP_ENABLED or not GRID_INFO:
                return geometry
            cell_size = GRID_INFO['cell_size']
            grid_start_x, grid_start_y = GRID_INFO['start_pos']
            legal_area = GRID_INFO['legal_area']
            # 计算在网格中的坐标
            grid_x = round((geometry.x() - grid_start_x) / cell_size)
            grid_y = round((geometry.y() - grid_start_y) / cell_size)
            # 计算网格中的尺寸（最小为1个网格单元）
            grid_width = max(1, round(geometry.width() / cell_size))
            grid_height = max(1, round(geometry.height() / cell_size))
            # 转换为实际像素坐标
            x = grid_start_x + grid_x * cell_size
            y = grid_start_y + grid_y * cell_size
            width = grid_width * cell_size
            height = grid_height * cell_size
            # 确保在合适区域内
            snapped_rect = QRect(x, y, width, height)
            if not legal_area.contains(snapped_rect):
                # 如果超出边界，调整到最近合适位置
                if x < legal_area.x():
                    x = legal_area.x()
                if y < legal_area.y():
                    y = legal_area.y()
                if x + width > legal_area.right():
                    x = legal_area.right() - width
                if y + height > legal_area.bottom():
                    y = legal_area.bottom() - height
                snapped_rect = QRect(x, y, width, height)
            return snapped_rect
        except Exception as e:
            print(f"[ERROR] 时钟组件网格吸附失败: {e}")
            return geometry
    def _check_rest_window_boundary(geometry):  # 检查时钟组件是否在合适边界内
        try:
            if not GRID_INFO:
                return True
            legal_area = GRID_INFO['legal_area']
            return legal_area.contains(geometry)
        except Exception as e:
            print(f"[ERROR] 时钟组件边界检查失败: {e}")
            return False
    def _apply_rest_window_geometry():  # 应用时钟组件的几何属性（包含网格吸附）
        try:
            if not rest_window:
                return
            current_geo = rest_window.geometry()
            # 应用网格吸附
            if GRID_SNAP_ENABLED and GRID_INFO:
                snapped_geo = _snap_rest_window_to_grid(current_geo)
                if snapped_geo != current_geo:
                    rest_window.setGeometry(snapped_geo)
            # 边界检查
            if not _check_rest_window_boundary(rest_window.geometry()):
                # 如果超出边界，移动到默认位置
                saved_geo = rest_settings.get("geometry",[SideLen, SideLen, (int(SideLen*4)), (int(SideLen*2))])
                rest_window.setGeometry(*saved_geo)
        except Exception as e:
            print(f"[ERROR] 应用时钟组件几何属性失败: {e}")

    def load_rest_settings():  # 加载响铃提示设置
        try:
            if getattr(sys, 'frozen', False):
                base_path = os.path.dirname(sys.executable)
                config_path = os.path.join(base_path, "config.json")
            else:
                config_path = "config.json"
            if os.path.exists(config_path):
                with open(config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                # 确保settings字段存在并读取rest_prompt
                rest_config = config.get("rest_prompt", {})  # 改
                return {
                    "geometry": rest_config.get("geometry",[SideLen, SideLen, (int(SideLen*4)), (int(SideLen*2))]),
                    "interval": rest_config.get("interval", 10)
                }
        except Exception as e:
            print(f"[ERROR] 加载响铃提示设置失败: {e}")
        return {"geometry": [100, 100, (int(SCR_WIDTH * 0.24)), (int(SCR_HEIGHT * 0.2))], "interval": 10}
    def save_rest_settings(geometry=None, interval=None):  # 保存提示间隔设置
        try:
            if getattr(sys, 'frozen', False):
                base_path = os.path.dirname(sys.executable)
                config_path = os.path.join(base_path, "config.json")
            else:
                config_path = "config.json"
            config = {}
            if os.path.exists(config_path):
                with open(config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
            if "rest_prompt" not in config:
                config["rest_prompt"] = {}
            if geometry:
                config["rest_prompt"]["geometry"] = geometry
            if interval is not None:
                config["rest_prompt"]["interval"] = interval
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=4)
        except Exception as e:
            print(f"[ERROR] 保存响铃提示设置失败: {e}")
    def update_time_display(): # 更新时间显示
        nonlocal time_display
        try:
            current_time = QTime.currentTime()
            time_str = current_time.toString("hh:mm:ss")
            formatted_time = f"{time_str[0]} {time_str[1]}:{time_str[3]} {time_str[4]}:{time_str[6]} {time_str[7]}"
            if time_display[0]:
                time_display[0].display(formatted_time)
        except Exception as e:
            print(f"[ERROR] 更新时间显示失败: {e}")
    def show_rest_alert():  #显示休息提醒
        try:
            if Mwin.isHidden() or Mwin.y() < 0:
                Anim_AppearMwin(Mwin)
            # 显示提醒消息
            def delayed_show_dialog():
                dialog = MessageDialog(
                    parent=Mwin, editable=False, default_text="", modal=False,
                    title="响铃提醒", message="铃声已响起！",
                    auto_close=5000)
                dialog.show()
            QTimer.singleShot(500, delayed_show_dialog)
            QTimer.singleShot(5000, lambda: Mwin_ToggleState(Mwin) if not Mwin.isHidden() else None)
            # 如果启用循环，重新设置定时器
            if REST_PROMPT_LOOP:
                if hasattr(rest_timer, 'last_interval'):
                    setup_rest_timer(rest_timer.last_interval // (60 * 1000))
                else:
                    setup_rest_timer(5)
                update_next_prompt_display()
            else:
                rest_timer.stop()
                update_next_prompt_display()
        except Exception as e:
            print(f"[ERROR] 显示休息提醒失败: {e}")
    def clear_rest_timer():  # 清除休息定时器
        try:
            rest_timer.stop()
            update_next_prompt_display()  # 更新显示
            print("[INFO] 下次响铃已取消")
        except Exception as e:
            print(f"[ERROR] 清除定时器失败: {e}")
    def update_next_prompt_display():  # 更新下次提示时间显示
        nonlocal next_prompt_display
        try:
            if rest_timer.isActive():
                # 计算下次提示时间
                remaining_time = rest_timer.remainingTime()
                next_time = QDateTime.currentDateTime().addMSecs(remaining_time)
                time_str = next_time.toString("hh:mm:ss")
                formatted_time = time_str
            else:
                formatted_time = "--:--:--"
            if next_prompt_display[0]:
                next_prompt_display[0].display(formatted_time)
        except Exception as e:
            print(f"[ERROR] 更新下次提示时间显示失败: {e}")
    def setup_rest_timer(minutes=None):  # 设置休息定时器
        try:
            rest_timer.stop()
            if minutes is not None and minutes > 0:
                interval = minutes * 60 * 1000
                rest_timer.start(interval)
                rest_timer.last_interval = interval
                update_next_prompt_display()
        except Exception as e:
            print(f"[ERROR] 设置定时器失败: {e}")
    def setup_next_prompt():  #设置下次提示时间
        try:
            default_interval = rest_settings.get("interval", 10)  # 改
            dialog = MessageDialog(
                parent=Mwin, editable=True, default_text=str(default_interval), modal=True,
                title="设置响铃间隔", message="请输入提示间隔时间（分钟）:",)
            result = dialog.exec()
            if result == QDialog.DialogCode.Accepted:
                minutes_text = dialog.user_input.strip()
                if minutes_text:
                    try:
                        minutes = int(minutes_text)
                        if minutes > 0:
                            setup_rest_timer(minutes)
                            save_rest_settings(interval=minutes)
                        else:
                            raise ValueError("时间必须大于0")
                    except ValueError as e:
                        error_dialog = MessageDialog(
                            parent=Mwin, editable=False, default_text="", modal=True,
                            title="输入错误", message="请输入有效的正整数")
                        error_dialog.exec()
        except Exception as e:
            print(f"[ERROR] 设置下次提示时间失败: {e}")
    def toggle_loop_prompt(state): # 切换循环提示设置
        global REST_PROMPT_LOOP
        REST_PROMPT_LOOP = bool(state)

    rest_settings = load_rest_settings()
    time_timer = QTimer()
    time_timer.timeout.connect(update_time_display)
    time_timer.start(1000)
    rest_timer.timeout.connect(show_rest_alert)
    try:
        mdi_area = Mwin.centralWidget()
        if not isinstance(mdi_area, QMdiArea):
            return
        # 创建休息提示子窗口
        rest_window = QMdiSubWindow(mdi_area)
        rest_window.setWindowTitle("简易时钟")
        saved_geo = rest_settings.get("geometry",[SideLen, SideLen, (int(SideLen*4)), (int(SideLen*2))])
        rest_window.setGeometry(*saved_geo)
        rest_window.setMinimumSize((int(SideLen*4)), (int(SideLen*2)))
        rest_window.setWindowFlags(
            Qt.WindowType.CustomizeWindowHint |
            Qt.WindowType.SubWindow |
            Qt.WindowType.WindowTitleHint)
        rest_window_style = f"QMdiSubWindow {{ border: 2px solid {ColorList[6]} !important; background-color: {ColorList[2]}; color: {ColorList[1]}; }}"
        rest_window.setStyleSheet(rest_window_style)
        def on_rest_window_geometry_changed():  # 自动保存几何属性
            if rest_window and rest_window.isVisible():
                current_geo = rest_window.geometry()
                if GRID_SNAP_ENABLED and GRID_INFO:
                    snapped_geo = _snap_rest_window_to_grid(current_geo)
                    if snapped_geo != current_geo:
                        rest_window.setGeometry(snapped_geo)
                        current_geo = snapped_geo
                geo_list = [current_geo.x(), current_geo.y(),
                            current_geo.width(), current_geo.height()]
                save_rest_settings(geometry=geo_list)
                rest_window.update()
                rest_window.repaint()
        rest_window.geometry_timer = QTimer()
        rest_window.geometry_timer.setSingleShot(True)
        rest_window.geometry_timer.timeout.connect(on_rest_window_geometry_changed)
        def rest_window_resize_event(event):
            # 正确调用父类的resizeEvent
            QMdiSubWindow.resizeEvent(rest_window, event)
            rest_window.geometry_timer.start(500)
            current_geo = rest_window.geometry()
            if not _check_rest_window_boundary(current_geo):
                _apply_rest_window_geometry()
            # 网格吸附
            elif GRID_SNAP_ENABLED and GRID_INFO:
                snapped_geo = _snap_rest_window_to_grid(current_geo)
                if snapped_geo != current_geo:
                    rest_window.setGeometry(snapped_geo)
            # 强制刷新样式
            rest_window.style().unpolish(rest_window)
            rest_window.style().polish(rest_window)
            rest_window.update()
        def rest_window_move_event(event):
            # 正确调用父类的moveEvent
            QMdiSubWindow.moveEvent(rest_window, event)
            rest_window.geometry_timer.start(500)
            current_geo = rest_window.geometry()
            if not _check_rest_window_boundary(current_geo):
                _apply_rest_window_geometry()
            # 网格吸附
            elif GRID_SNAP_ENABLED and GRID_INFO:
                snapped_geo = _snap_rest_window_to_grid(current_geo)
                if snapped_geo != current_geo:
                    rest_window.setGeometry(snapped_geo)
            # 强制刷新样式
            rest_window.style().unpolish(rest_window)
            rest_window.style().polish(rest_window)
            rest_window.update()
        rest_window.resizeEvent = rest_window_resize_event
        rest_window.moveEvent = rest_window_move_event
        # 创建中心widget
        central_widget = QWidget()
        rest_window.setWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.setSpacing(10)
        layout.setContentsMargins(20, 20, 20, 20)
        # 创建时间显示标签
        time_display[0] = QLCDNumber()
        print(f"[DEBUG] time_display已赋值: {time_display[0] is not None}")
        time_display[0].setDigitCount(11)
        time_display[0].setSegmentStyle(QLCDNumber.SegmentStyle.Flat)
        time_display_style = f"QLCDNumber {{ color: {ColorList[1]}; background-color: transparent; border: none; }}"
        time_display[0].setStyleSheet(time_display_style)
        time_display[0].setFixedSize(int(SideLen*3), int(SideLen))
        layout.addWidget(time_display[0], alignment=Qt.AlignmentFlag.AlignCenter)
        update_time_display()
        # 创建按钮水平布局
        controls_layout = QHBoxLayout()
        setup_button = QPushButton("下次响铃提示：")
        setup_button.setFixedSize(int(sidelen*6),int(sidelen*1.5))
        setup_button_style = f"""
        QPushButton {{background-color: {ColorList[3]};color: {ColorList[1]};border: 
            1px solid {ColorList[5]};
            font-size: {int(sidelen*0.8)}px;}}
        QPushButton:hover {{background-color: {ColorList[4]};}}
        QPushButton:pressed {{background-color: {ColorList[5]};}}"""
        setup_button.setStyleSheet(setup_button_style)
        setup_button.clicked.connect(setup_next_prompt)
        controls_layout.addWidget(setup_button)
        next_prompt_display[0] = QLCDNumber()
        next_prompt_display[0].setDigitCount(8)
        next_prompt_display[0].setSegmentStyle(QLCDNumber.SegmentStyle.Flat)
        next_prompt_style = f"QLCDNumber {{ color: {ColorList[1]}; background-color: transparent; border: none; }}"
        next_prompt_display[0].setStyleSheet(next_prompt_style)
        next_prompt_display[0].setFixedSize(int(sidelen*6), int(sidelen*2))
        next_prompt_display[0].setToolTip("下次响铃提示时间\n点击取消提示")  # 悬停提示
        next_prompt_display[0].mousePressEvent = lambda event: clear_rest_timer()  # 点击事件
        controls_layout.addWidget(next_prompt_display[0])
        # 创建循环提示复选框
        loop_checkbox = QCheckBox("循环")
        loop_checkbox.setChecked(REST_PROMPT_LOOP)
        loop_checkbox_style = f"""
        QCheckBox {{color: {ColorList[1]};font-size: {int(sidelen*0.8)}px;spacing: {int(sidelen*0.125)}px;
            border: 3px solid {ColorList[5]};padding: {int(sidelen*0.125)}px {int(sidelen*0.125)}px;background-color: {ColorList[2]};}}
        QCheckBox::indicator {{width: {int(sidelen*0.8)}px;height: {int(sidelen*0.8)}px;}}
        QCheckBox::indicator:unchecked {{border: 1px solid {ColorList[6]};background-color: {ColorList[2]};}}
        QCheckBox::indicator:checked {{border: 1px solid {ColorList[6]};background-color: {ColorList[6]};}}"""
        loop_checkbox.setStyleSheet(loop_checkbox_style)
        loop_checkbox.stateChanged.connect(toggle_loop_prompt)
        controls_layout.addStretch(1)
        controls_layout.addWidget(loop_checkbox)
        layout.addLayout(controls_layout)
        update_next_prompt_display()
        mdi_area.addSubWindow(rest_window)
        rest_window.show()
        rest_window.time_timer = time_timer
        rest_window.rest_timer = rest_timer
        return rest_window
    except Exception as e:
        print(f"[ERROR] 创建简易时钟窗口失败: {e}")
        return None

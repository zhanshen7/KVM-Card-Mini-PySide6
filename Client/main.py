import os
import sys
import tempfile
import time
from typing import ClassVar

import hid_def
import pythoncom
import pyWinhook as pyHook
import qdarktheme
import win32api
import win32con
import yaml
from default import default_config
from loguru import logger
from PySide6.QtCore import QPoint, QSize, Qt, QTimer, QTranslator, Signal
from PySide6.QtGui import QCursor, QGuiApplication, QIcon, QPixmap, QSurfaceFormat
from PySide6.QtMultimedia import (
    QAudioFormat,
    QAudioSink,
    QAudioSource,
    QCamera,
    QMediaCaptureSession,
    QMediaDevices,
)
from PySide6.QtMultimediaWidgets import QVideoWidget
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QDialog,
    QDialogButtonBox,
    QLabel,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QWidget,
)
from ui import (
    device_setup_dialog_ui,
    main_ui,
)

kb_buffer = [1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
mouse_buffer = [2, 0, 0, 0, 0, 0, 0, 0, 0]
mouse_buffer_rel = [7, 0, 0, 0, 0, 0, 0, 0, 0]
PATH = os.path.dirname(os.path.abspath(__file__))
ARGV_PATH = os.path.dirname(os.path.abspath(sys.argv[0]))

if not os.path.exists(os.path.join(ARGV_PATH, "config.yaml")):
    with open(os.path.join(ARGV_PATH, "config.yaml"), "w", encoding="utf-8") as f:
        f.write(default_config)

translation = True
try:
    with open(os.path.join(ARGV_PATH, "config.yaml"), "r", encoding="utf-8") as load_f:
        config = yaml.safe_load(load_f)["config"]
        translation = config["translation"]
except (KeyError, OSError, TypeError, UnicodeError, yaml.YAMLError):
    pass


def str_bool(b) -> str:
    if not translation:
        return str(b)
    return "启用" if b else "禁用"


class NullWriter:
    def write(self, data):
        return len(data)

    def flush(self):
        pass


null_writer = NullWriter()

# 屏蔽控制台输出
if sys.argv[-1] != "debug":
    sys.stdout = null_writer
    sys.stderr = null_writer

    logger.remove()
else:
    hid_def.set_verbose(True)


def load_icon(name) -> QIcon:
    return QIcon(f"{PATH}/icons/24_dark/{name}.png")


def load_pixmap(name) -> QPixmap:
    return QPixmap(f"{PATH}/icons/24_dark/{name}.png")


class MyPushButton(QPushButton):
    def setPixmap(self, pixmap):
        icon = QIcon(pixmap)
        self.setIcon(icon)
        self.setIconSize(QSize(18, 18))


class MyDeviceSetupDialog(QDialog, device_setup_dialog_ui.Ui_Dialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setupUi(self)
        self.setWindowModality(Qt.WindowModality.NonModal)
        self.setLayoutDirection(Qt.LayoutDirection.LeftToRight)
        self.buttonBox.setOrientation(Qt.Orientation.Horizontal)
        self.buttonBox.setStandardButtons(
            QDialogButtonBox.StandardButton.Cancel | QDialogButtonBox.StandardButton.Ok
        )
        self.setWindowFlags(
            Qt.WindowType.CustomizeWindowHint | Qt.WindowType.WindowCloseButtonHint
        )
        self.setWindowFlags(self.windowFlags() | Qt.WindowType.WindowStaysOnTopHint)


class MyMainWindow(QMainWindow, main_ui.Ui_MainWindow):
    _hid_signal = Signal(list)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setupUi(self)
        self.setFocusPolicy(Qt.FocusPolicy.StrongFocus)

        self.camera = None
        self.capture_session = None
        self.camera_opened = False
        self.camera_info = None
        self.audio_source = None
        self.audio_sink = None
        self.audio_source_device = None
        self.audio_sink_device = None
        self.audio_pending = bytearray()
        self.audio_frame_size = 0
        self.audio_max_pending = 0
        self.device_connected = False

        # 子窗口
        self.device_setup_dialog = MyDeviceSetupDialog()
        # 导入外部数据
        try:
            with open(
                os.path.join(PATH, "Data", "keyboard_scancode2hid.yml"),
                encoding="utf-8",
            ) as load_f:
                self.keyboard_scancode2hid = yaml.safe_load(load_f)
            with open(
                os.path.join(ARGV_PATH, "config.yaml"), encoding="utf-8"
            ) as load_f:
                self.configfile = yaml.safe_load(load_f)
            self.config = self.configfile["config"]
            self.video_config = self.configfile["video_config"]
            self.audio_config = self.configfile["audio_config"]
            self.config.setdefault("mouse_report_freq", 125)
            self.audio_config.setdefault("buffer_ms", 40)
            self.relative_mouse_speed = self.config["relative_mouse_speed"]
            if self.config["mouse_report_freq"] != 0:
                self.mouse_report_interval = 1000 / self.config["mouse_report_freq"]
                self.dynamic_mouse_report_interval = False
            else:
                self.mouse_report_interval = 10
                self.dynamic_mouse_report_interval = True
        except (
            KeyError,
            OSError,
            TypeError,
            UnicodeError,
            ValueError,
            yaml.YAMLError,
        ) as error:
            QMessageBox.critical(
                self,
                self.tr("Error"),
                f"Import config error:\n {error}\n\n"
                + self.tr(
                    "Check the config.yaml and restart the program\nor delete the config.yaml to reset the config file."
                ),
            )
            sys.exit(1)
        # 加载配置文件
        self.status = {
            "fullscreen": False,
            "topmost": False,
            "mouse_capture": False,
            "hide_cursor": False,
            "init_ok": False,
            "screen_height": 0,
            "relative_mouse": False,
        }

        # 获取显示器分辨率大小
        screen = QGuiApplication.primaryScreen()
        self.status["screen_height"] = screen.availableGeometry().height()

        # 窗口图标
        self.setWindowIcon(QIcon(f"{PATH}/icons/icon.ico"))
        self.device_setup_dialog.setWindowIcon(load_icon("import"))

        # 状态栏图标
        self.statusbar_lock_btn = MyPushButton()
        self.statusbar_lock_btn.setPixmap(load_pixmap("lock"))
        self.statusbar_lock_btn.setToolTip(self.tr("Lock remote screen"))
        self.statusbar_screenshot_btn = MyPushButton()
        self.statusbar_screenshot_btn.setPixmap(load_pixmap("capture"))
        self.statusbar_screenshot_btn.setToolTip(self.tr("Screenshot"))
        self.statusbar_btn5 = MyPushButton()
        self.statusbar_btn5.setPixmap(load_pixmap("hook-off"))

        self.statusbar_icon1 = MyPushButton()
        self.statusbar_icon2 = MyPushButton()
        self.statusbar_icon3 = MyPushButton()
        self.statusbar_icon1.setPixmap(load_pixmap("video-off"))
        self.statusbar_icon2.setPixmap(load_pixmap("keyboard-off"))
        self.statusbar_icon3.setPixmap(load_pixmap("mouse-off"))
        self.statusbar_fullscreen_btn = MyPushButton()
        self.statusbar_fullscreen_btn.setPixmap(load_pixmap("fullscreen"))
        self.statusbar_fullscreen_btn.setToolTip(self.tr("Fullscreen"))

        self.statusBar().setStyleSheet("padding: 0px;")
        self.statusBar().addPermanentWidget(self.statusbar_lock_btn)
        self.statusBar().addPermanentWidget(self.statusbar_screenshot_btn)
        self.statusBar().addPermanentWidget(self.statusbar_btn5)
        self.statusBar().addPermanentWidget(self.statusbar_icon2)
        self.statusBar().addPermanentWidget(self.statusbar_icon3)
        self.statusBar().addPermanentWidget(self.statusbar_icon1)
        self.statusBar().addPermanentWidget(self.statusbar_fullscreen_btn)

        self.statusbar_lock_btn.clicked.connect(self.lock_remote_screen)
        self.statusbar_screenshot_btn.clicked.connect(self.capture_screenshot)
        self.statusbar_btn5.clicked.connect(self.system_hook_func)
        self.statusbar_icon1.clicked.connect(self.device_config)
        self.statusbar_icon2.clicked.connect(lambda: self.reset_keymouse(4))
        self.statusbar_icon3.clicked.connect(self.toggle_mouse_capture)
        self.statusbar_fullscreen_btn.clicked.connect(self.fullscreen_func)

        # 菜单栏图标
        self.action_video_devices.setIcon(load_icon("import"))
        self.action_video_device_connect.setIcon(load_icon("video"))
        self.action_video_device_disconnect.setIcon(load_icon("video-off"))
        self.actionMinimize.setIcon(load_icon("window-minimize"))
        self.actionexit.setIcon(load_icon("window-close"))
        self.actionReload_MCU.setIcon(load_icon("reload"))
        self.actionReload_Key_Mouse.setIcon(load_icon("reload"))
        self.action_fullscreen.setIcon(load_icon("fullscreen"))
        self.action_Resize_window.setIcon(load_icon("resize"))
        self.actionKeep_ratio.setIcon(load_icon("ratio"))
        self.actionResetKeyboard.setIcon(load_icon("reload"))
        self.actionResetMouse.setIcon(load_icon("reload"))
        self.actionCapture_mouse.setIcon(load_icon("mouse"))
        self.actionRelease_mouse.setIcon(load_icon("mouse-off"))
        self.actionKeep_on_top.setIcon(load_icon("topmost"))
        self.actionHide_cursor.setIcon(load_icon("cursor"))
        self.actionSystem_hook.setIcon(load_icon("hook"))
        self.actionRelative_mouse.setIcon(load_icon("relative"))

        if self.video_config["keep_aspect_ratio"]:
            self.set_checked(self.actionKeep_ratio, True)

        # 初始化监视器
        self.videoWidget = QVideoWidget()
        self.videoWidget.setAttribute(Qt.WidgetAttribute.WA_OpaquePaintEvent)
        self.takeCentralWidget()
        self.setCentralWidget(self.videoWidget)
        self.videoWidget.setMouseTracking(True)
        video_surface = self.videoWidget.findChild(QWidget)
        if video_surface is not None:
            video_surface.setMouseTracking(True)
        self.videoWidget.hide()

        self.disconnect_label = QLabel()
        self.disconnect_label.setPixmap(load_pixmap("disconnected"))
        self.disconnect_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.disconnect_label.setMouseTracking(True)
        self.takeCentralWidget()
        self.setCentralWidget(self.disconnect_label)
        self.disconnect_label.show()

        # 按键绑定
        self.action_video_device_connect.triggered.connect(
            lambda: self.set_device(True)
        )
        self.action_video_device_disconnect.triggered.connect(
            lambda: self.set_device(False)
        )
        self.action_video_devices.triggered.connect(self.device_config)
        self.actionReload_Key_Mouse.triggered.connect(lambda: self.reset_keymouse(4))
        self.actionMinimize.triggered.connect(self.showMinimized)
        self.actionexit.triggered.connect(self.close)

        self.device_setup_dialog.comboBox.currentIndexChanged.connect(
            self.update_device_info
        )
        self.device_setup_dialog.comboBox_2.currentTextChanged.connect(
            self.update_pixel_formats
        )
        self.device_setup_dialog.comboBox_3.currentTextChanged.connect(
            self.update_frame_rates
        )

        self.action_fullscreen.triggered.connect(self.fullscreen_func)
        self.action_Resize_window.triggered.connect(self.resize_window_func)
        self.actionKeep_ratio.triggered.connect(self.keep_ratio_func)
        self.actionKeep_on_top.triggered.connect(self.topmost_func)

        self.actionRelease_mouse.triggered.connect(self.release_mouse)
        self.actionCapture_mouse.triggered.connect(self.capture_mouse)
        self.actionResetKeyboard.triggered.connect(lambda: self.reset_keymouse(1))
        self.actionResetMouse.triggered.connect(lambda: self.reset_keymouse(3))
        self.actionReload_MCU.triggered.connect(lambda: self.reset_keymouse(2))

        self.actionHide_cursor.triggered.connect(self.hide_cursor_func)
        self.actionSystem_hook.triggered.connect(self.system_hook_func)
        self.actionRelative_mouse.triggered.connect(self.relative_mouse_func)

        self.device_setup_dialog.checkBoxAudio.setChecked(
            self.audio_config["audio_support"]
        )
        self.device_setup_dialog.checkBoxAudio.stateChanged.connect(
            self.audio_checkbox_switch
        )
        self.audio_checkbox_switch()

        # 设置聚焦方式
        self.statusbar_lock_btn.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.statusbar_screenshot_btn.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.statusbar_btn5.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.statusbar_icon1.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.statusbar_icon2.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.statusbar_icon3.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.statusbar_fullscreen_btn.setFocusPolicy(Qt.FocusPolicy.NoFocus)

        self.mouse_action_timer = QTimer(self)
        self.mouse_action_timer.timeout.connect(self.mouse_action_timeout)
        self.check_device_timer = QTimer(self)
        self.check_device_timer.timeout.connect(self.check_device_status)
        self.check_device_timer.start(1000)

        self.reset_keymouse(4)

        # self.setMouseTracking(True)

        self.mouse_scroll_timer = QTimer(self)
        self.mouse_scroll_timer.timeout.connect(self.mouse_scroll_stop)

        self._new_mouse_report = 0
        self._last_mouse_pos = None
        self.rel_x = 0
        self.rel_y = 0
        self._mouse_report_timer = QTimer(self)
        self._mouse_report_timer.timeout.connect(self.mouse_report_timeout)
        self._mouse_report_timer.start(self.mouse_report_interval)
        self._hid_signal.connect(self.hid_report)

        self.hook_state = False
        self.hook_manager = pyHook.HookManager()
        self.hook_manager.KeyDown = self.hook_keyboard_down_event
        self.hook_manager.KeyUp = self.hook_keyboard_up_event
        self.pythoncom_timer = QTimer(self)
        self.pythoncom_timer.timeout.connect(lambda: pythoncom.PumpWaitingMessages())
        self.hook_pressed_keys = []
        self._local_screenshot_trigger = False

        self.status["init_ok"] = True

        self._restore_track = False

        self.camera_list_inited = False
        if self.video_config["auto_connect"]:
            self.device_setup_dialog.checkBoxAutoConnect.setChecked(True)
            QTimer.singleShot(1000, lambda: self.set_device(True, center=True))

    code_remap: ClassVar[dict[str, int]] = {
        "Rcontrol": 0x011D,
        "Rmenu": 0x0138,
        "Lwin": 0x015B,
        "Rwin": 0x015C,
    }

    def hook_keyboard_down_event(self, event):
        logger.debug(f"Hook: {event.Key} {event.ScanCode}")
        if (
            self._local_screenshot_trigger
            and getattr(event, "KeyID", None) == win32con.VK_SNAPSHOT
        ):
            return True
        if event.Key in self.code_remap:
            scan_code = self.code_remap[event.Key]
        else:
            scan_code = event.ScanCode
        if scan_code not in self.hook_pressed_keys:
            self.hook_pressed_keys.append(scan_code)
            self.keyPress(scan_code)
        return False

    def hook_keyboard_up_event(self, event):
        if (
            self._local_screenshot_trigger
            and getattr(event, "KeyID", None) == win32con.VK_SNAPSHOT
        ):
            return True
        if event.Key in self.code_remap:
            scan_code = self.code_remap[event.Key]
        else:
            scan_code = event.ScanCode
        self.keyRelease(scan_code)
        try:
            self.hook_pressed_keys.remove(scan_code)
        except ValueError:
            pass
        return False

    def toggle_mouse_capture(self):
        if self.status["mouse_capture"]:
            self.release_mouse()
            self.statusBar().showMessage(self.tr("Mouse capture off"))
        else:
            self.capture_mouse()

    def set_checked(self, attr, state):
        font = attr.font()
        font.setBold(state)
        attr.setFont(font)
        if attr.isCheckable():
            # attr.setChecked(bold)
            text = attr.text().replace(" ·", "")
            if state:
                text += " ·"
            attr.setText(text)

    def save_config(self):
        # 保存配置文件
        with open(os.path.join(ARGV_PATH, "config.yaml"), "w", encoding="utf-8") as f:
            yaml.dump(self.configfile, f)

    def audio_checkbox_switch(self):
        if self.device_setup_dialog.checkBoxAudio.isChecked():
            self.device_setup_dialog.comboBox_4.show()
            self.device_setup_dialog.comboBox_5.show()
            self.device_setup_dialog.comboBox_audio_buffer.show()
            self.device_setup_dialog.label_4.show()
            self.device_setup_dialog.label_5.show()
            self.device_setup_dialog.label_audio_buffer.show()
            self.device_setup_dialog.label_7.show()
            self.device_setup_dialog.setMaximumHeight(330)
            self.device_setup_dialog.setMinimumHeight(330)
            self.update_audio_devices()
        else:
            self.device_setup_dialog.comboBox_4.hide()
            self.device_setup_dialog.comboBox_5.hide()
            self.device_setup_dialog.comboBox_audio_buffer.hide()
            self.device_setup_dialog.label_4.hide()
            self.device_setup_dialog.label_5.hide()
            self.device_setup_dialog.label_audio_buffer.hide()
            self.device_setup_dialog.label_7.hide()
            self.device_setup_dialog.setMaximumHeight(230)
            self.device_setup_dialog.setMinimumHeight(230)
        self.device_setup_dialog.adjustSize()

    def update_audio_devices(self):
        self.device_setup_dialog.comboBox_4.clear()
        self.device_setup_dialog.comboBox_5.clear()
        self.device_setup_dialog.comboBox_4.addItem("Default")
        self.device_setup_dialog.comboBox_5.addItem("Default")
        in_devices = QMediaDevices.audioInputs()
        out_devices = QMediaDevices.audioOutputs()
        devices = ["Default"]
        for i in in_devices:
            self.device_setup_dialog.comboBox_4.addItem(i.description())
            devices.append(i.description())
        if self.audio_config["audio_device_in"] in devices:
            self.device_setup_dialog.comboBox_4.setCurrentText(
                self.audio_config["audio_device_in"]
            )
        else:
            self.device_setup_dialog.comboBox_4.setCurrentIndex(0)
        devices = ["Default"]
        for i in out_devices:
            self.device_setup_dialog.comboBox_5.addItem(i.description())
            devices.append(i.description())
        if self.audio_config["audio_device_out"] in devices:
            self.device_setup_dialog.comboBox_5.setCurrentText(
                self.audio_config["audio_device_out"]
            )
        else:
            self.device_setup_dialog.comboBox_5.setCurrentIndex(0)
        self.update_audio_buffer_options()

    def update_audio_buffer_options(self):
        buffer_ms = int(self.audio_config.get("buffer_ms", 40))
        combo_box = self.device_setup_dialog.comboBox_audio_buffer
        combo_box.blockSignals(True)
        combo_box.clear()
        for option in (20, 40, 60, 120):
            combo_box.addItem(f"{option} ms", option)
        if combo_box.findData(buffer_ms) < 0:
            combo_box.addItem(f"{buffer_ms} ms", buffer_ms)
        combo_box.setCurrentIndex(combo_box.findData(buffer_ms))
        combo_box.blockSignals(False)

    # 弹出采集卡设备设置窗口，并打开采集卡设备
    def device_config(self):
        self.device_setup_dialog.comboBox.clear()
        cameras = QMediaDevices.videoInputs()
        remember_name = self.video_config["device_name"]
        # self.video_config["device_name"] = ""
        devices = []
        for camera in cameras:
            self.device_setup_dialog.comboBox.addItem(camera.description())
            devices.append(camera.description())
        self.camera_list_inited = True
        if remember_name in devices:
            self.device_setup_dialog.comboBox.setCurrentText(remember_name)
            self.update_device_info()
            resolution_str = (
                str(self.video_config["resolution_X"])
                + "x"
                + str(self.video_config["resolution_Y"])
            )
            self.device_setup_dialog.comboBox_2.setCurrentText(resolution_str)
            self.device_setup_dialog.comboBox_3.setCurrentText(
                self.video_config["format"]
            )
            self.update_frame_rates()
            self._select_frame_rate(self.video_config.get("frame_rate"))
        else:
            self.device_setup_dialog.comboBox.setCurrentIndex(0)
            self.update_device_info()
            self.device_setup_dialog.comboBox_2.setCurrentIndex(0)
            self.device_setup_dialog.comboBox_3.setCurrentIndex(0)
            self.update_frame_rates()
            try:
                self.video_config["resolution_X"] = (
                    self.device_setup_dialog.comboBox_2.currentText().split("x")[0]
                )
                self.video_config["resolution_Y"] = (
                    self.device_setup_dialog.comboBox_2.currentText().split("x")[1]
                )
            except IndexError:
                self.video_config["resolution_X"] = 0
                self.video_config["resolution_Y"] = 0
            self.video_config["format"] = (
                self.device_setup_dialog.comboBox_3.currentText()
            )
            self.video_config["frame_rate"] = (
                self.device_setup_dialog.comboBox_frame_rate.currentData()
            )

        if self.device_setup_dialog.checkBoxAudio.isChecked():
            self.update_audio_devices()

        wm_pos = self.geometry()
        wm_size = self.size()
        self.device_setup_dialog.move(
            wm_pos.x() + wm_size.width() // 2 - self.device_setup_dialog.width() // 2,
            wm_pos.y() + wm_size.height() // 2 - self.device_setup_dialog.height() // 2,
        )
        # 如果选择设备
        ret = self.device_setup_dialog.exec()

        if not ret:
            return
        try:
            self.video_config["device_name"] = (
                self.device_setup_dialog.comboBox.currentText()
            )
            self.video_config["resolution_X"] = int(
                self.device_setup_dialog.comboBox_2.currentText().split("x")[0]
            )
            self.video_config["resolution_Y"] = int(
                self.device_setup_dialog.comboBox_2.currentText().split("x")[1]
            )
            self.video_config["format"] = (
                self.device_setup_dialog.comboBox_3.currentText()
            )
            self.video_config["frame_rate"] = float(
                self.device_setup_dialog.comboBox_frame_rate.currentData()
            )

            if self.device_setup_dialog.checkBoxAudio.isChecked():
                self.audio_config["audio_device_in"] = (
                    self.device_setup_dialog.comboBox_4.currentText()
                )
                self.audio_config["audio_device_out"] = (
                    self.device_setup_dialog.comboBox_5.currentText()
                )
                self.audio_config["buffer_ms"] = int(
                    self.device_setup_dialog.comboBox_audio_buffer.currentData()
                )
        except (IndexError, TypeError, ValueError):
            self.video_alert(self.tr("Selected invalid device"))
            return
        logger.debug(self.video_config)
        try:
            self.set_device(True, center=True)
            self.video_config["auto_connect"] = (
                self.device_setup_dialog.checkBoxAutoConnect.isChecked()
            )
            self.audio_config["audio_support"] = (
                self.device_setup_dialog.checkBoxAudio.isChecked()
            )
            self.save_config()
        except (OSError, yaml.YAMLError) as error:
            logger.error(f"Unable to save configuration: {error}")

    # 获取采集卡分辨率
    def update_device_info(self):
        previous_resolution = self.device_setup_dialog.comboBox_2.currentText()
        self.device_setup_dialog.comboBox_2.clear()
        self.device_setup_dialog.comboBox_3.clear()
        self.device_setup_dialog.comboBox_frame_rate.clear()
        cameras = QMediaDevices.videoInputs()
        if not self.camera_list_inited:
            for camera in cameras:
                if camera.description() == self.video_config["device_name"]:
                    self.camera_info = camera
                    break
            else:
                logger.error(self.tr("Target video device not found"))
                self.camera_info = None
                return
        else:
            try:
                self.camera_info = cameras[
                    self.device_setup_dialog.comboBox.currentIndex()
                ]
            except IndexError:
                self.camera_info = None
                return
        res_list = []
        for i in self.camera_info.videoFormats():
            resolutions_str = f"{i.resolution().width()}x{i.resolution().height()}"
            if resolutions_str not in res_list:
                res_list.append(resolutions_str)
                self.device_setup_dialog.comboBox_2.addItem(resolutions_str)
        if previous_resolution in res_list:
            self.device_setup_dialog.comboBox_2.setCurrentText(previous_resolution)
        self.update_pixel_formats()

    def update_pixel_formats(self):
        if self.camera_info is None:
            return
        try:
            width, height = map(
                int, self.device_setup_dialog.comboBox_2.currentText().split("x")
            )
        except ValueError:
            return
        previous_format = self.device_setup_dialog.comboBox_3.currentText()
        formats = []
        for camera_format in self.camera_info.videoFormats():
            if (
                camera_format.resolution().width() == width
                and camera_format.resolution().height() == height
            ):
                pixel_format = camera_format.pixelFormat().name.split("_")[1]
                if pixel_format not in formats:
                    formats.append(pixel_format)
        self.device_setup_dialog.comboBox_3.blockSignals(True)
        self.device_setup_dialog.comboBox_3.clear()
        self.device_setup_dialog.comboBox_3.addItems(formats)
        if previous_format in formats:
            self.device_setup_dialog.comboBox_3.setCurrentText(previous_format)
        self.device_setup_dialog.comboBox_3.blockSignals(False)
        self.update_frame_rates()

    def update_frame_rates(self):
        if self.camera_info is None:
            return
        try:
            width, height = map(
                int, self.device_setup_dialog.comboBox_2.currentText().split("x")
            )
        except ValueError:
            return
        pixel_format = self.device_setup_dialog.comboBox_3.currentText()
        frame_rates = sorted(
            {
                camera_format.maxFrameRate()
                for camera_format in self.camera_info.videoFormats()
                if camera_format.resolution().width() == width
                and camera_format.resolution().height() == height
                and camera_format.pixelFormat().name.split("_")[1] == pixel_format
            },
            reverse=True,
        )
        self.device_setup_dialog.comboBox_frame_rate.blockSignals(True)
        self.device_setup_dialog.comboBox_frame_rate.clear()
        for frame_rate in frame_rates:
            self.device_setup_dialog.comboBox_frame_rate.addItem(
                f"{frame_rate:.1f} FPS", frame_rate
            )
        self.device_setup_dialog.comboBox_frame_rate.blockSignals(False)

    def _select_frame_rate(self, frame_rate):
        if frame_rate is None:
            return
        frame_rate_index = self.device_setup_dialog.comboBox_frame_rate.findData(
            float(frame_rate)
        )
        if frame_rate_index >= 0:
            self.device_setup_dialog.comboBox_frame_rate.setCurrentIndex(
                frame_rate_index
            )

    def _show_disconnected(self):
        self.takeCentralWidget()
        self.setCentralWidget(self.disconnect_label)
        self.videoWidget.hide()
        self.disconnect_label.show()
        self.setWindowTitle("USB KVM Client")

    def camera_error_occurred(self, error, message):
        if self.camera is None:
            return
        device_name = (
            self.camera_info.description()
            if self.camera_info is not None
            else self.tr("Unknown device")
        )
        error_s = f"Device: {device_name}\nReturned: {message or error}\n\n" + self.tr(
            "Device disconnected"
        )
        self.camera_info = None
        self.device_event_handle("video_error")
        self._show_disconnected()
        QTimer.singleShot(0, self._teardown_media)
        QMessageBox.critical(self, self.tr("Device Error"), error_s)

    def capture_screenshot(self):
        self._local_screenshot_trigger = True
        win32api.keybd_event(win32con.VK_SNAPSHOT, 0, 0, 0)
        win32api.keybd_event(win32con.VK_SNAPSHOT, 0, win32con.KEYEVENTF_KEYUP, 0)
        QTimer.singleShot(100, self._clear_local_screenshot_trigger)
        self.statusBar().showMessage(self.tr("Local Print Screen triggered"))

    def lock_remote_screen(self):
        if not self.device_connected:
            self.statusBar().showMessage(self.tr("Keyboard Mouse connect error"))
            return
        self.update_kb_hid(0xE3, True)
        self.update_kb_hid(0x0F, True)
        self.update_kb_hid(0x0F, False)
        self.update_kb_hid(0xE3, False)
        self.statusBar().showMessage(self.tr("Remote screen locked"))

    def _clear_local_screenshot_trigger(self):
        self._local_screenshot_trigger = False

    @staticmethod
    def _audio_format_description(audio_format):
        return (
            f"{audio_format.sampleRate()} Hz, "
            f"{audio_format.channelCount()} channel(s), "
            f"{audio_format.sampleFormat().name}"
        )

    def _common_audio_format(self):
        """Return a format accepted natively by both selected audio devices."""
        candidates = [
            self.audio_in_device.preferredFormat(),
            self.audio_out_device.preferredFormat(),
        ]
        for sample_rate in (48000, 44100):
            for channel_count in (2, 1):
                for sample_format in (
                    QAudioFormat.SampleFormat.Int16,
                    QAudioFormat.SampleFormat.Float,
                ):
                    audio_format = QAudioFormat()
                    audio_format.setSampleRate(sample_rate)
                    audio_format.setChannelCount(channel_count)
                    audio_format.setSampleFormat(sample_format)
                    candidates.append(audio_format)
        for audio_format in candidates:
            if (
                audio_format.isValid()
                and self.audio_in_device.isFormatSupported(audio_format)
                and self.audio_out_device.isFormatSupported(audio_format)
            ):
                return audio_format
        return None

    def _pump_audio(self):
        """Move PCM from capture to playback without changing its format."""
        if self.audio_source_device is None or self.audio_sink_device is None:
            return
        data = self.audio_source_device.readAll().data()
        if data:
            self.audio_pending.extend(data)
        if len(self.audio_pending) > self.audio_max_pending:
            dropped = len(self.audio_pending) - self.audio_max_pending
            dropped -= dropped % self.audio_frame_size
            if dropped:
                del self.audio_pending[:dropped]
                logger.warning(f"Audio buffer overflow: dropped {dropped} bytes")
        while self.audio_pending:
            written = self.audio_sink_device.write(bytes(self.audio_pending))
            if written <= 0:
                break
            del self.audio_pending[:written]

    def _start_direct_audio(self):
        audio_format = self.audio_format
        if audio_format is None:
            return False
        buffer_ms = max(20, min(int(self.audio_config.get("buffer_ms", 40)), 120))
        self.audio_frame_size = audio_format.bytesPerFrame()
        buffer_size = (
            audio_format.sampleRate() * self.audio_frame_size * buffer_ms // 1000
        )
        self.audio_max_pending = buffer_size * 2
        self.audio_source = QAudioSource(self.audio_in_device, audio_format, self)
        self.audio_sink = QAudioSink(self.audio_out_device, audio_format, self)
        self.audio_source.setBufferSize(buffer_size)
        self.audio_sink.setBufferSize(buffer_size)
        self.audio_source_device = self.audio_source.start()
        self.audio_sink_device = self.audio_sink.start()
        if self.audio_source_device is None or self.audio_sink_device is None:
            logger.error(
                f"Unable to start direct audio: input={self.audio_source.error().name}, "
                f"output={self.audio_sink.error().name}"
            )
            self._stop_direct_audio()
            return False
        self.audio_source_device.readyRead.connect(self._pump_audio)
        logger.info(
            "Direct audio started: "
            f"{self.audio_in_device.description()} -> "
            f"{self.audio_out_device.description()} "
            f"({self._audio_format_description(audio_format)}, {buffer_ms} ms)"
        )
        return True

    def _stop_direct_audio(self):
        if self.audio_source is not None:
            self.audio_source.stop()
            self.audio_source.deleteLater()
        if self.audio_sink is not None:
            self.audio_sink.stop()
            self.audio_sink.deleteLater()
        self.audio_source = None
        self.audio_sink = None
        self.audio_source_device = None
        self.audio_sink_device = None
        self.audio_pending.clear()

    def _teardown_media(self):
        """Stop and dispose of all video and direct-PCM audio objects."""
        self._stop_direct_audio()

        camera = self.camera
        capture_session = self.capture_session
        self.camera = None
        self.capture_session = None
        self.camera_opened = False

        if capture_session is not None:
            capture_session.deleteLater()
        if camera is not None:
            camera.stop()
            camera.deleteLater()

    # 初始化指定配置视频设备
    def setup_device(self):
        if self.camera_info is None:
            self.update_device_info()
            if self.camera_info is None:
                self.video_alert(self.tr("Target video device not found"))
                return False
        self.camera = QCamera(self.camera_info)
        matching_formats = [
            camera_format
            for camera_format in self.camera_info.videoFormats()
            if camera_format.resolution().width() == self.video_config["resolution_X"]
            and camera_format.resolution().height() == self.video_config["resolution_Y"]
            and camera_format.pixelFormat().name.split("_")[1]
            == self.video_config["format"]
        ]
        if not matching_formats:
            self._teardown_media()
            self.video_alert(
                self.tr("Unsupported combination of resolution and format")
            )
            return False
        selected_frame_rate = self.video_config.get("frame_rate")
        if selected_frame_rate is not None:
            matching_formats = [
                camera_format
                for camera_format in matching_formats
                if abs(camera_format.maxFrameRate() - float(selected_frame_rate)) < 0.01
            ] or matching_formats
        self.camera.setCameraFormat(
            max(
                matching_formats, key=lambda camera_format: camera_format.maxFrameRate()
            )
        )

        if self.device_setup_dialog.checkBoxAudio.isChecked():
            in_devices = QMediaDevices.audioInputs()
            out_devices = QMediaDevices.audioOutputs()
            in_device_name = self.audio_config["audio_device_in"]
            out_device_name = self.audio_config["audio_device_out"]
            if in_device_name == "Default":
                in_device = QMediaDevices.defaultAudioInput()
            else:
                for i in in_devices:
                    if i.description() == in_device_name:
                        in_device = i
                        break
                else:
                    in_device = None
            if out_device_name == "Default":
                out_device = QMediaDevices.defaultAudioOutput()
            else:
                for i in out_devices:
                    if i.description() == out_device_name:
                        out_device = i
                        break
                else:
                    out_device = None
            if in_device is None or out_device is None:
                self._teardown_media()
                self.video_alert(self.tr("Audio device not found"))
                return False
            self.audio_in_device = in_device
            self.audio_out_device = out_device
            self.audio_format = self._common_audio_format()
            if self.audio_format is None:
                self._teardown_media()
                self.video_alert(
                    self.tr("Audio devices do not share a supported PCM format")
                )
                return False

        self.camera.errorOccurred.connect(self.camera_error_occurred)
        self.camera.start()
        if not self.camera.isActive():
            self._teardown_media()
            self.video_alert(self.tr("Video device connect failed"))
            return False

        self.capture_session = QMediaCaptureSession()
        self.capture_session.setCamera(self.camera)
        self.capture_session.setVideoOutput(self.videoWidget)

        if self.device_setup_dialog.checkBoxAudio.isChecked():
            if not self._start_direct_audio():
                self._teardown_media()
                self.video_alert(self.tr("Audio stream could not be opened"))
                return False
            logger.debug("Audio device ok")
        return True

    # 视频设备错误提示
    def video_alert(self, s):
        QMessageBox.critical(self, self.tr("Video Error"), s)

    # 启用和禁用视频设备
    def set_device(self, state, center=False):
        if state:
            self._teardown_media()
            if not self.setup_device():
                self._show_disconnected()
                return
            if not self.status["fullscreen"]:
                self.resize_window_func(center=center)
            if self.camera is None:
                self._show_disconnected()
                return
            fps = self.camera.cameraFormat().maxFrameRate()
            self.device_event_handle("video_ok")
            self.takeCentralWidget()
            self.setCentralWidget(self.videoWidget)
            self.disconnect_label.hide()
            self.videoWidget.show()
            self.setWindowTitle(
                f"USB KVM Client - {self.video_config['resolution_X']}x{self.video_config['resolution_Y']} @ {fps:.1f}"
            )
            if self.dynamic_mouse_report_interval and fps > 0:
                self.mouse_report_interval = 1000 / fps
                self._mouse_report_timer.setInterval(round(self.mouse_report_interval))
        else:
            if (
                self.camera is None
                and self.capture_session is None
                and self.audio_source is None
                and self.audio_sink is None
            ):
                return
            self._teardown_media()
            self.device_event_handle("video_close")
            self._show_disconnected()

    # 捕获鼠标功能
    def capture_mouse(self):
        self.status["mouse_capture"] = True
        self.statusbar_icon3.setPixmap(load_pixmap("mouse"))
        self.statusBar().showMessage(
            self.tr("Mouse capture on (Press Right-Alt to release)")
        )

    # 释放鼠标功能
    def release_mouse(self):
        self.status["mouse_capture"] = False
        self._clear_mouse_state()
        if self.device_connected:
            self._hid_signal.emit(mouse_buffer.copy())
            self._hid_signal.emit(mouse_buffer_rel.copy())
        self.statusbar_icon3.setPixmap(load_pixmap("mouse-off"))

    def _clear_mouse_state(self):
        mouse_buffer[2:] = [0] * (len(mouse_buffer) - 2)
        mouse_buffer_rel[2:] = [0] * (len(mouse_buffer_rel) - 2)
        self.rel_x = 0
        self.rel_y = 0
        self._last_mouse_pos = None
        self._new_mouse_report = 0

    # 通过视频设备分辨率调整窗口大小
    def resize_window_func(self, center=True):
        if self.status["fullscreen"]:
            return
        if self.status["screen_height"] - self.video_config["resolution_Y"] < 100:
            self.showNormal()
            self.resize(
                int(
                    self.status["screen_height"]
                    * (2 / 3)
                    * self.video_config["resolution_X"]
                    / (self.video_config["resolution_Y"] + 66)
                ),
                int(self.status["screen_height"] * (2 / 3)),
            )
            self.showMaximized()
        else:
            self.showNormal()
            self.resize(
                self.video_config["resolution_X"],
                self.video_config["resolution_Y"] + 66,
            )
            if center:
                qr = self.frameGeometry()
                cp = QGuiApplication.primaryScreen().availableGeometry().center()
                qr.moveCenter(cp)
                self.move(qr.topLeft())
        if self.video_config["keep_aspect_ratio"]:
            self.videoWidget.setAspectRatioMode(Qt.AspectRatioMode.KeepAspectRatio)
        else:
            self.videoWidget.setAspectRatioMode(Qt.AspectRatioMode.IgnoreAspectRatio)

    def keep_ratio_func(self):
        self.video_config["keep_aspect_ratio"] = not self.video_config[
            "keep_aspect_ratio"
        ]
        if self.video_config["keep_aspect_ratio"]:
            self.videoWidget.setAspectRatioMode(Qt.AspectRatioMode.KeepAspectRatio)
            self.set_checked(self.actionKeep_ratio, True)
        else:
            self.videoWidget.setAspectRatioMode(Qt.AspectRatioMode.IgnoreAspectRatio)
            self.set_checked(self.actionKeep_ratio, False)
        if not self.status["fullscreen"]:
            self.resize(self.width(), self.height() + 1)
        self.statusBar().showMessage(
            self.tr("Keep aspect ratio: ")
            + str_bool(self.video_config["keep_aspect_ratio"])
        )
        self.save_config()

    # 重置键盘鼠标
    def reset_keymouse(self, s):
        if s == 1:  # keyboard
            for i in range(1, len(kb_buffer)):
                kb_buffer[i] = 0
            hidinfo = hid_def.hid_report(kb_buffer)
            if hidinfo == 1 or hidinfo == 4:
                self.device_event_handle("hid_error")
            elif hidinfo == 0:
                self.device_event_handle("hid_ok")
        elif s == 2:  # MCU
            hidinfo = hid_def.hid_report([4, 0])
            if hidinfo == 1 or hidinfo == 4:
                self.device_event_handle("hid_error")
            elif hidinfo == 0:
                self.device_event_handle("hid_ok")
            self.statusbar_icon2.setPixmap(load_pixmap("keyboard-off"))
            self.status["mouse_capture"] = False
            self.statusbar_icon3.setPixmap(load_pixmap("mouse-off"))
        elif s == 3:  # mouse
            self._clear_mouse_state()
            hidinfos = (
                hid_def.hid_report(mouse_buffer),
                hid_def.hid_report(mouse_buffer_rel),
            )
            if any(hidinfo in (1, 4) for hidinfo in hidinfos):
                self.device_event_handle("hid_error")
            elif any(hidinfo == 0 for hidinfo in hidinfos):
                self.device_event_handle("hid_ok")
        elif s == 4:  # hid
            hid_code = hid_def.init_usb(hid_def.vendor_id, hid_def.usage_page)
            if hid_code == 0:
                self.device_event_handle("hid_init_ok")
            else:
                self.device_event_handle("hid_init_error")

    # 设备事件处理
    def device_event_handle(self, s):
        if s == "hid_error":
            self.statusBar().showMessage(
                self.tr("Keyboard Mouse connect error, try to <Reload Key/Mouse>")
            )
            self.statusbar_icon2.setPixmap(load_pixmap("keyboard-off"))
            self.status["mouse_capture"] = False
            self.statusbar_icon3.setPixmap(load_pixmap("mouse-off"))
            self.device_connected = False
            self.check_device_timer.stop()
        elif s == "video_error":
            self.statusBar().showMessage(self.tr("Video device error"))
            self.statusbar_icon1.setPixmap(load_pixmap("video-off"))
            self.camera_opened = False
            self.check_device_timer.start(1000)
        elif s == "video_close":
            self.statusBar().showMessage(self.tr("Video device close"))
            self.statusbar_icon1.setPixmap(load_pixmap("video-off"))
            self.camera_opened = False
            self.statusbar_icon3.setPixmap(load_pixmap("mouse-off"))
            self.status["mouse_capture"] = False
            self.check_device_timer.start(1000)
        elif s == "hid_init_error":
            self.statusBar().showMessage(self.tr("Keyboard Mouse initialization error"))
            self.statusbar_icon2.setPixmap(load_pixmap("keyboard-off"))
            self.device_connected = False
            self.check_device_timer.stop()
        elif s == "hid_init_ok":
            self.statusBar().showMessage(self.tr("Keyboard Mouse initialization done"))
            self.statusbar_icon2.setPixmap(load_pixmap("keyboard"))
            self.device_connected = True
            self.check_device_timer.start(1000)
        elif s == "hid_ok":
            self.statusbar_icon2.setPixmap(load_pixmap("keyboard"))
            self.device_connected = True
            self.check_device_timer.start(1000)
        elif s == "video_ok":
            self.statusBar().showMessage(self.tr("Video device connected"))
            self.statusbar_icon1.setPixmap(load_pixmap("video"))
            self.status["mouse_capture"] = True
            self.statusbar_icon3.setPixmap(load_pixmap("mouse"))
            self.camera_opened = True
            self.check_device_timer.stop()
        elif s == "device_disconnect":
            self.statusBar().showMessage(self.tr("Device disconnect"))
            self.statusbar_icon2.setPixmap(load_pixmap("keyboard-off"))
            self.statusbar_icon3.setPixmap(load_pixmap("mouse-off"))
            self.status["mouse_capture"] = False
            self.device_connected = False
            self.check_device_timer.stop()
        elif s == "video_disconnect":
            self.statusBar().showMessage(self.tr("Device disconnect"))
            self.statusbar_icon1.setPixmap(load_pixmap("video-off"))
            self.camera_opened = False
            self.check_device_timer.stop()

    # 检查连接状态
    def check_device_status(self):
        if self.device_connected and not hid_def.check_connection():
            self.device_event_handle("device_disconnect")
        # if self.camera_opened:
        #     if self.camera.availability() != QMultimedia.Available:
        #         self.device_event_handle("video_disconnect")

    def system_hook_func(self):
        self.set_system_hook(not self.hook_state)

    def set_system_hook(self, enabled):
        if self.hook_state == enabled:
            return
        self.hook_state = enabled
        self.set_checked(self.actionSystem_hook, enabled)
        self.statusBar().showMessage(self.tr("System hook: ") + str_bool(enabled))
        if enabled:
            self.pythoncom_timer.start(5)
            self.hook_manager.HookKeyboard()
            self.statusbar_btn5.setPixmap(load_pixmap("hook"))
        else:
            self.hook_manager.UnhookKeyboard()
            self.pythoncom_timer.stop()
            self.statusbar_btn5.setPixmap(load_pixmap("hook-off"))

    def relative_mouse_func(self):
        self.status["relative_mouse"] = not self.status["relative_mouse"]
        self.set_checked(self.actionRelative_mouse, self.status["relative_mouse"])
        message = self.tr("Relative mouse: ") + str_bool(self.status["relative_mouse"])
        if self.status["relative_mouse"]:
            message += self.tr(" (Press Right-Alt to release mouse)")
        self.statusBar().showMessage(message)

    # 全屏幕切换
    def fullscreen_func(self):
        self.status["fullscreen"] = not self.status["fullscreen"]
        if self.status["fullscreen"]:
            if not self.config["fullscreen_alert_showed"]:
                alert = QMessageBox(self)
                alert.setWindowTitle(self.tr("Fullscreen"))
                alert.setText(
                    self.tr("Press F11 to toggle fullscreen")
                    + self.tr("\nStay cursor at left top corner to show toolbar")
                )
                alert.addButton(
                    self.tr("I know it, don't show again"),
                    QMessageBox.ButtonRole.AcceptRole,
                )
                alert.exec()
                self.config["fullscreen_alert_showed"] = True
                self.save_config()
            self.set_system_hook(True)
            self.set_hide_cursor(True)
            self.showFullScreen()
            self.action_fullscreen.setChecked(True)
            self.action_Resize_window.setEnabled(False)
            self.statusBar().hide()
            self.menuBar().hide()
            self.set_checked(self.action_fullscreen, True)
        else:
            self.set_system_hook(False)
            self.set_hide_cursor(False)
            self.showNormal()
            self.action_fullscreen.setChecked(False)
            self.action_Resize_window.setEnabled(True)
            self.statusBar().show()
            self.menuBar().show()
            self.set_checked(self.action_fullscreen, False)

    # 隐藏指针
    def hide_cursor_func(self):
        self.set_hide_cursor(not self.status["hide_cursor"])

    def set_hide_cursor(self, enabled):
        self.status["hide_cursor"] = enabled
        self.set_checked(self.actionHide_cursor, enabled)
        if self.status["mouse_capture"]:
            self.setCursor(
                Qt.CursorShape.BlankCursor
                if enabled or self.status["relative_mouse"]
                else Qt.CursorShape.ArrowCursor
            )
        self.statusBar().showMessage(
            self.tr("Hide cursor when capture mouse: ") + str_bool(enabled)
        )

    # 保持窗口在最前
    def topmost_func(self):
        self.status["topmost"] = not self.status["topmost"]
        self.setWindowFlags(
            Qt.WindowType.WindowStaysOnTopHint
            if self.status["topmost"]
            else Qt.WindowType.Widget
        )
        self.show()
        self.statusBar().showMessage(
            self.tr("Window always on top: ") + str_bool(self.status["topmost"])
        )
        self.set_checked(self.actionKeep_on_top, self.status["topmost"])

    # 窗口失焦事件
    def changeEvent(self, event):
        super().changeEvent(event)
        if not getattr(self, "status", {}).get("init_ok"):
            return
        if not self.isActiveWindow():  # 窗口失去焦点时重置键盘，防止卡键
            self.reset_keymouse(1)

    def mouseButton_to_int(self, s: Qt.MouseButton):
        if s == Qt.MouseButton.LeftButton:
            return 1
        elif s == Qt.MouseButton.RightButton:
            return 2
        elif s == Qt.MouseButton.MiddleButton:
            return 4
        elif s == Qt.MouseButton.XButton1:
            return 8
        elif s == Qt.MouseButton.XButton2:
            return 16
        else:
            return 0

    # 鼠标按下事件
    _last_click_time = 0

    def mousePressEvent(self, event):
        if (
            not self.status["mouse_capture"]
            and self.device_connected
            and event.button() == Qt.MouseButton.LeftButton
            and self.camera_opened
        ):
            if time.perf_counter() - self._last_click_time < 0.2:
                self.capture_mouse()
            else:
                self._last_click_time = time.perf_counter()
                self.statusBar().showMessage(self.tr("Double click to capture mouse"))
        if not self.status["mouse_capture"]:
            return
        if not self.status["relative_mouse"]:
            buffer = mouse_buffer
        else:
            buffer = mouse_buffer_rel
        buffer[2] = buffer[2] | self.mouseButton_to_int(event.button())
        self._hid_signal.emit(buffer)

    # 鼠标松开事件
    def mouseReleaseEvent(self, event):
        if not self.status["mouse_capture"]:
            return
        if not self.status["relative_mouse"]:
            buffer = mouse_buffer
        else:
            buffer = mouse_buffer_rel
        buffer[2] = buffer[2] ^ self.mouseButton_to_int(event.button())
        if buffer[2] < 0 or buffer[2] > 7:
            buffer[2] = 0
        self._hid_signal.emit(buffer)

    # 鼠标滚动事件
    def wheelEvent(self, event):
        if not self.status["mouse_capture"]:
            return
        if not self.status["relative_mouse"]:
            buffer = mouse_buffer
            bit = 7
        else:
            buffer = mouse_buffer_rel
            bit = 5
        if event.angleDelta().y() == 120:
            buffer[bit] = 0x01
        elif event.angleDelta().y() == -120:
            buffer[bit] = 0xFF
        else:
            buffer[bit] = 0
        self._hid_signal.emit(buffer)
        if self.mouse_scroll_timer.isActive():
            self.mouse_scroll_timer.stop()
        self.mouse_scroll_timer.start(100)

    def mouse_scroll_stop(self):
        self.mouse_scroll_timer.stop()
        if not self.status["relative_mouse"]:
            buffer = mouse_buffer
            bit = 7
        else:
            buffer = mouse_buffer_rel
            bit = 5
        buffer[bit] = 0
        self._hid_signal.emit(buffer)

    def mouse_action_timeout(self):
        if self.mouse_action_target == "menuBar":
            self.menuBar().show()
        elif self.mouse_action_target == "statusBar":
            self.statusBar().show()
        self.mouse_action_timer.stop()

    def hid_report(self, buf: list[int]):
        hidinfo = hid_def.hid_report(buf)
        if hidinfo == 1 or hidinfo == 4:
            self.device_event_handle("hid_error")

    def mouseMoveEvent(self, event):
        p = event.position().toPoint()
        x, y = p.x(), p.y()
        in_menu_bar_hotspot = False
        in_statusbar_hotspot = False
        if self.status["fullscreen"]:
            hotspot_width = 200
            hotspot_height = 60
            in_menu_bar_hotspot = x < hotspot_width and y < hotspot_height
            in_statusbar_hotspot = (
                x >= self.width() - hotspot_width
                and y >= self.height() - hotspot_height
            )
            if in_menu_bar_hotspot or in_statusbar_hotspot:
                if (
                    in_menu_bar_hotspot
                    and self.menuBar().isHidden()
                    and not self.mouse_action_timer.isActive()
                ):
                    self.mouse_action_target = "menuBar"
                    self.mouse_action_timer.start(500)
                elif (
                    in_statusbar_hotspot
                    and self.statusBar().isHidden()
                    and not self.mouse_action_timer.isActive()
                ):
                    self.mouse_action_target = "statusBar"
                    self.mouse_action_timer.start(500)
            else:
                if not self.menuBar().isHidden():
                    self.menuBar().hide()
                if not self.statusBar().isHidden():
                    self.statusBar().hide()
                if self.mouse_action_timer.isActive():
                    self.mouse_action_timer.stop()
        if not (self.status["mouse_capture"]):
            self.setCursor(Qt.CursorShape.ArrowCursor)
            return
        if in_menu_bar_hotspot or in_statusbar_hotspot:
            self.setCursor(Qt.CursorShape.ArrowCursor)
            self._last_mouse_pos = None
            return
        elif self.status["hide_cursor"] or self.status["relative_mouse"]:
            self.setCursor(Qt.CursorShape.BlankCursor)
        else:
            self.setCursor(Qt.CursorShape.ArrowCursor)
        if not self.status["relative_mouse"]:
            self._last_mouse_pos = None
            if not self.camera_opened:
                x_res = self.disconnect_label.width()
                y_res = self.disconnect_label.height()
                width = self.disconnect_label.width()
                height = self.disconnect_label.height()
                # x_pos = self.disconnect_label.pos().x()
                y_pos = self.disconnect_label.pos().y()
            else:
                x_res = self.video_config["resolution_X"]
                y_res = self.video_config["resolution_Y"]
                width = self.videoWidget.width()
                height = self.videoWidget.height()
                # x_pos = self.videoWidget.pos().x()
                y_pos = self.videoWidget.pos().y()
            x_diff = 0
            y_diff = 0
            if self.video_config["keep_aspect_ratio"]:
                cam_scale = y_res / x_res
                finder_scale = height / width
                if finder_scale > cam_scale:
                    x_diff = 0
                    y_diff = height - width * cam_scale
                elif finder_scale < cam_scale:
                    x_diff = width - height / cam_scale
                    y_diff = 0
            x_hid = (x - x_diff / 2) / (width - x_diff)
            y_hid = (y - y_diff / 2 - y_pos) / (height - y_diff)
            x_hid = max(min(x_hid, 1), 0)
            y_hid = max(min(y_hid, 1), 0)
            x_hid = int(x_hid * 0x7FFF)
            y_hid = int(y_hid * 0x7FFF)
            mouse_buffer[3] = x_hid & 0xFF
            mouse_buffer[4] = x_hid >> 8
            mouse_buffer[5] = y_hid & 0xFF
            mouse_buffer[6] = y_hid >> 8
            self._new_mouse_report = 1
        else:
            middle_pos = self.mapToGlobal(QPoint(self.width() // 2, self.height() // 2))
            mouse_pos = QCursor.pos()
            if self._last_mouse_pos is not None:
                self.rel_x += (
                    mouse_pos.x() - self._last_mouse_pos.x()
                ) * self.relative_mouse_speed
                self.rel_y += (
                    mouse_pos.y() - self._last_mouse_pos.y()
                ) * self.relative_mouse_speed
                self._new_mouse_report = 2
                self._last_mouse_pos = mouse_pos
                if (
                    abs(mouse_pos.x() - middle_pos.x()) > 25
                    or abs(mouse_pos.y() - middle_pos.y()) > 25
                ):
                    QCursor.setPos(middle_pos)
                    self._last_mouse_pos = middle_pos
            else:
                self._last_mouse_pos = middle_pos
                QCursor.setPos(middle_pos)

    def leaveEvent(self, event):
        if getattr(self, "status", {}).get("fullscreen"):
            self.setCursor(Qt.CursorShape.ArrowCursor)
            self._last_mouse_pos = None
        super().leaveEvent(event)

    def mouse_report_timeout(self):
        if self._new_mouse_report == 1:
            self._hid_signal.emit(mouse_buffer)
        elif self._new_mouse_report == 2:
            x_hid = max(min(round(self.rel_x), 127), -127)
            y_hid = max(min(round(self.rel_y), 127), -127)
            self.rel_x -= x_hid
            self.rel_y -= y_hid
            mouse_buffer_rel[3] = x_hid & 0xFF
            mouse_buffer_rel[4] = y_hid & 0xFF
            self._hid_signal.emit(mouse_buffer_rel)
            mouse_buffer_rel[3] = 0
            mouse_buffer_rel[4] = 0
        self._new_mouse_report = 0

    scan_to_b2: ClassVar[dict[int, int]] = {
        0x001D: 1,  # Left Control
        0x002A: 2,  # Left Shift
        0x0038: 4,  # Left Alt
        0x015B: 8,  # Left GUI
        0x011D: 16,  # Right Control
        0x0036: 32,  # Right Shift
        0x0138: 64,  # Right Alt
        0x015C: 128,  # Right GUI
    }

    hid_to_b2: ClassVar[dict[int, int]] = {
        0xE0: 1,  # Left Control
        0xE1: 2,  # Left Shift
        0xE2: 4,  # Left Alt
        0xE3: 8,  # Left GUI
        0xE4: 16,  # Right Control
        0xE5: 32,  # Right Shift
        0xE6: 64,  # Right Alt
        0xE7: 128,  # Right GUI
    }

    def update_kb(self, scancode: int, state: bool):
        if state:
            if scancode in self.scan_to_b2:
                kb_buffer[2] |= self.scan_to_b2[scancode]
            else:
                scancode2hid = self.keyboard_scancode2hid.get(scancode, 0)
                if scancode2hid == 0:
                    if scancode != 256:
                        logger.warning(f"scancode2hid not found: {scancode}")
                    return
                for i in range(4, 10):
                    if kb_buffer[i] == scancode2hid:
                        return
                    if kb_buffer[i] == 0:
                        kb_buffer[i] = scancode2hid
                        break
                else:
                    logger.warning("Buffer overflow")
        else:
            if scancode in self.scan_to_b2:
                kb_buffer[2] &= ~self.scan_to_b2[scancode]
            else:
                scancode2hid = self.keyboard_scancode2hid.get(scancode, 0)
                if scancode2hid == 0:
                    if scancode != 256:
                        logger.warning(f"scancode2hid not found: {scancode}")
                    return
                for i in range(4, 10):
                    if kb_buffer[i] == scancode2hid:
                        kb_buffer[i] = 0
                        break
                else:
                    logger.warning("Key not found in buffer")
        if not self.device_connected:
            return 0
        self._hid_signal.emit(kb_buffer)
        return 0

    def update_kb_hid(self, hid: int, state: bool):
        if state:
            if hid in self.hid_to_b2:
                kb_buffer[2] |= self.hid_to_b2[hid]
            else:
                for i in range(4, 10):
                    if kb_buffer[i] == hid:
                        return
                    if kb_buffer[i] == 0:
                        kb_buffer[i] = hid
                        break
                else:
                    logger.warning("Buffer overflow")
        else:
            if hid in self.hid_to_b2:
                kb_buffer[2] &= ~self.hid_to_b2[hid]
            else:
                for i in range(4, 10):
                    if kb_buffer[i] == hid:
                        kb_buffer[i] = 0
                        break
                else:
                    logger.warning("Key not found in buffer")
        self._hid_signal.emit(kb_buffer)
        return 0

    # 键盘按下事件
    def keyPressEvent(self, event):
        if self._local_screenshot_trigger and event.key() == Qt.Key.Key_Print:
            return
        if event.isAutoRepeat():
            return
        if event.key() == Qt.Key.Key_F11:
            self.fullscreen_func()
            return
        self.keyPress(event.nativeScanCode())

    def keyPress(self, scancode: int):
        if scancode == 0x0057:  # F11
            self.fullscreen_func()
            return
        if scancode == 0x0138:  # Right Alt
            self.release_mouse()
            self.statusBar().showMessage(self.tr("Mouse capture off"))
            return
        self.update_kb(scancode, True)

    # 键盘松开事件
    def keyReleaseEvent(self, event):
        if event.isAutoRepeat():
            return
        if self._local_screenshot_trigger and event.key() == Qt.Key.Key_Print:
            return
        self.keyRelease(event.nativeScanCode())

    def keyRelease(self, scancode: int):
        if scancode == 0x0057:  # F11 is reserved for fullscreen
            return
        if scancode == 0x0138:  # Right Alt is reserved for mouse release
            return
        self.update_kb(scancode, False)

    def closeEvent(self, event):
        if self.config.get("confirm_before_close", True):
            close_dialog = QMessageBox(self)
            close_dialog.setIcon(QMessageBox.Icon.Question)
            close_dialog.setWindowTitle(self.tr("Exit"))
            close_dialog.setText(self.tr("Are you sure you want to close the window?"))
            close_dialog.setStandardButtons(
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            close_dialog.setDefaultButton(QMessageBox.StandardButton.No)
            do_not_ask_again = QCheckBox(self.tr("Don't ask again"), close_dialog)
            close_dialog.setCheckBox(do_not_ask_again)
            if close_dialog.exec() != QMessageBox.StandardButton.Yes:
                event.ignore()
                return
            if do_not_ask_again.isChecked():
                self.config["confirm_before_close"] = False
                try:
                    self.save_config()
                except (OSError, yaml.YAMLError) as error:
                    logger.error(
                        f"Unable to save close confirmation preference: {error}"
                    )
        self.mouse_action_timer.stop()
        self.mouse_scroll_timer.stop()
        self.check_device_timer.stop()
        self._mouse_report_timer.stop()
        self.set_system_hook(False)
        if self.device_connected:
            self.reset_keymouse(1)
            self.reset_keymouse(3)
        self.set_device(False)
        super().closeEvent(event)

    def focusInEvent(self, event):
        if self._restore_track:
            self._restore_track = False
            self.status["mouse_capture"] = True
        if self.hook_state:
            self.pythoncom_timer.start(5)
            self.hook_manager.HookKeyboard()
            self.statusbar_btn5.setPixmap(load_pixmap("hook"))
        super().focusInEvent(event)

    def focusOutEvent(self, event):
        if self.status["mouse_capture"]:
            self.status["mouse_capture"] = False
            self._restore_track = True
        if self.hook_state:
            self.hook_manager.UnhookKeyboard()
            self.pythoncom_timer.stop()
            self.statusbar_btn5.setPixmap(load_pixmap("hook-off"))
        super().focusOutEvent(event)


def clear_splash():
    if "NUITKA_ONEFILE_PARENT" in os.environ:
        splash_filename = os.path.join(
            tempfile.gettempdir(),
            f"onefile_{int(os.environ['NUITKA_ONEFILE_PARENT'])}_splash_feedback.tmp",
        )
        if os.path.exists(splash_filename):
            os.unlink(splash_filename)


def main():
    argv = [
        *sys.argv,
        "-platform",
        "windows:darkmode=2",
        "--style",
        "Windows",
    ]  # or "Fusion" ?
    surface_format = QSurfaceFormat.defaultFormat()
    surface_format.setSwapInterval(0)
    QSurfaceFormat.setDefaultFormat(surface_format)
    app = QApplication(argv)
    translator = QTranslator(app)
    if translation and translator.load(os.path.join(PATH, "trans_cn.qm")):
        app.installTranslator(translator)
    translator2 = QTranslator(app)
    if translation and translator2.load(os.path.join(PATH, "qtbase_cn.qm")):
        app.installTranslator(translator2)
    myWin = MyMainWindow()
    qdarktheme.setup_theme(
        theme="dark",
        custom_colors={
            "[dark]": {
                "background>base": "#1f2021",
            }
        },
    )
    myWin.show()
    clear_splash()
    return app.exec()


if __name__ == "__main__":
    main()

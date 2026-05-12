import sys
import os
import json
import time

from pathlib import Path
from datetime import datetime, timezone, timedelta

JST = timezone(timedelta(hours=9))

BOOT_START = time.perf_counter()

from PyQt6.QtWidgets import (
    QApplication,
    QWidget,
    QMenu,
    QHBoxLayout,
    QLabel,
)

from PyQt6.QtGui import (
    QColor,
    QAction,
    QPainter,
    QPen,
    QFont,
)

from PyQt6.QtCore import (
    Qt,
    QTimer,
)

# =========================
# 起動ログ
# =========================

def boot_log(label):

    now = time.perf_counter()

    print(
        f"[BOOT] {label}: "
        f"{now - BOOT_START:.3f} sec"
    )


# =========================
# 定数
# =========================

if getattr(sys, 'frozen', False):

    APP_DIR = Path(sys.executable).parent

else:

    APP_DIR = Path(__file__).parent

SETTINGS_FILE = APP_DIR / "settings.json"

DEFAULT_SETTINGS = {

    "font_family": "Arial",

    "font_color": "#FFFFFF",

    "background_color": "#000000",

    "window_color": "#111111",

    "settings_window_color": "#2B2B2B",

    "window_opacity": 0.95,

    "window_width": 700,

    "window_height": 300,

    "x": 100,

    "y": 100
}


# =========================
# 数字パネル
# =========================

class FlipDigit(QWidget):

    def __init__(self):

        super().__init__()

        self.value = "0"

        self.font_family = "Arial"

        self.font_size = 100

        self.text_color = QColor("#FFFFFF")

        self.bg_color = QColor("#000000")

        self.line_color = QColor("#111111")

        self.cached_font = None

    def set_value(self, value):

        if self.value != value:

            self.value = value

            self.update()

    def apply_style(self, settings, font_size):

        self.font_family = settings["font_family"]

        self.font_size = font_size

        self.text_color = QColor(
            settings["font_color"]
        )

        self.bg_color = QColor(
            settings["background_color"]
        )

        self.line_color = QColor(
            settings["window_color"]
        )

        font = QFont(
            self.font_family,
            self.font_size
        )

        font.setBold(True)

        self.cached_font = font

        self.update()

    def paintEvent(self, event):

        painter = QPainter(self)

        rect = self.rect()

        # 背景
        painter.setBrush(self.bg_color)

        painter.setPen(Qt.PenStyle.NoPen)

        painter.drawRoundedRect(
            rect,
            20,
            20
        )

        # 中央線
        pen = QPen(self.line_color)

        pen.setWidth(2)

        painter.setPen(pen)

        painter.drawLine(
            rect.left(),
            rect.center().y(),
            rect.right(),
            rect.center().y()
        )

        # フォント
        painter.setFont(self.cached_font)

        painter.setPen(self.text_color)

        painter.drawText(
            rect,
            Qt.AlignmentFlag.AlignCenter,
            self.value
        )


# =========================
# 設定ウィンドウ
# =========================

class SettingsDialog:

    pass


# =========================
# メイン時計
# =========================

class ClockWidget(QWidget):

    def __init__(self):

        super().__init__()

        boot_log("ClockWidget start")

        self.settings = self.load_settings()

        boot_log("settings loaded")

        self.setup_window()

        self.create_ui()

        # 遅延適用
        QTimer.singleShot(
            0,
            self.delayed_init
        )

        boot_log("ClockWidget end")

    # =====================
    # 遅延初期化
    # =====================

    def delayed_init(self):

        self.apply_settings()

        self.start_timer()

        boot_log("delayed init complete")

    # =====================
    # ウィンドウ
    # =====================

    def setup_window(self):

        self.setWindowFlags(
            Qt.WindowType.Window |
            Qt.WindowType.WindowStaysOnTopHint
        )

        self.setMinimumSize(
            300,
            150
        )

        self.resize(
            self.settings["window_width"],
            self.settings["window_height"]
        )

        self.move(
            self.settings["x"],
            self.settings["y"]
        )

    # =====================
    # UI
    # =====================

    def create_ui(self):

        layout = QHBoxLayout()

        layout.setSpacing(10)

        layout.setContentsMargins(
            15,
            15,
            15,
            15
        )

        self.setLayout(layout)

        self.digits = []

        # 時
        for _ in range(2):

            digit = FlipDigit()

            layout.addWidget(digit)

            self.digits.append(digit)

        # :
        self.colon = QLabel(":")

        self.colon.setAlignment(
            Qt.AlignmentFlag.AlignCenter
        )

        self.colon.setFixedWidth(20)

        layout.addWidget(self.colon)

        # 分
        for _ in range(2):

            digit = FlipDigit()

            layout.addWidget(digit)

            self.digits.append(digit)

    # =====================
    # 設定画面
    # =====================

    def open_settings(self):

        # 遅延 import
        from PyQt6.QtWidgets import (
            QColorDialog,
            QFontDialog,
            QVBoxLayout,
            QPushButton,
            QDialog,
            QSlider,
        )

        class SettingsDialog(QDialog):

            def __init__(dialog_self, parent, settings):

                super().__init__(parent)

                dialog_self.parent_widget = parent

                dialog_self.settings = (
                    settings.copy()
                )

                dialog_self.setWindowTitle(
                    "時計設定"
                )

                dialog_self.setFixedWidth(320)

                dialog_self.setWindowFlags(
                    dialog_self.windowFlags()
                    | Qt.WindowType.WindowStaysOnTopHint
                )

                layout = QVBoxLayout()

                layout.setSpacing(15)

                layout.setContentsMargins(
                    20,
                    20,
                    20,
                    20
                )

                dialog_self.setLayout(layout)

                # フォント
                font_btn = QPushButton(
                    "フォント変更"
                )

                layout.addWidget(font_btn)

                def change_font():

                    current_font = QFont(
                        dialog_self.settings[
                            "font_family"
                        ],
                        12
                    )

                    dialog = QFontDialog(
                        current_font,
                        dialog_self
                    )

                    dialog.currentFontChanged.connect(
                        preview_font
                    )

                    if dialog.exec():

                        font = dialog.selectedFont()

                        dialog_self.settings[
                            "font_family"
                        ] = font.family()

                        preview_update()

                def preview_font(font):

                    dialog_self.settings[
                        "font_family"
                    ] = font.family()

                    preview_update()

                font_btn.clicked.connect(
                    change_font
                )

                # 色
                text_btn = QPushButton(
                    "文字色変更"
                )

                layout.addWidget(text_btn)

                def change_color():

                    dialog = QColorDialog(
                        QColor(
                            dialog_self.settings[
                                "font_color"
                            ]
                        ),
                        dialog_self
                    )

                    dialog.currentColorChanged.connect(
                        preview_color
                    )

                    if dialog.exec():

                        dialog_self.settings[
                            "font_color"
                        ] = (
                            dialog.selectedColor()
                            .name()
                        )

                        preview_update()

                def preview_color(color):

                    dialog_self.settings[
                        "font_color"
                    ] = color.name()

                    preview_update()

                text_btn.clicked.connect(
                    change_color
                )

                # 透明度
                opacity_label = QLabel(
                    "透明度"
                )

                layout.addWidget(
                    opacity_label
                )

                slider = QSlider(
                    Qt.Orientation.Horizontal
                )

                slider.setMinimum(30)

                slider.setMaximum(100)

                slider.setValue(
                    int(
                        dialog_self.settings[
                            "window_opacity"
                        ] * 100
                    )
                )

                layout.addWidget(slider)

                def opacity_changed(value):

                    dialog_self.settings[
                        "window_opacity"
                    ] = value / 100

                    preview_update()

                slider.valueChanged.connect(
                    opacity_changed
                )

                ok_btn = QPushButton(
                    "OK"
                )

                ok_btn.clicked.connect(
                    dialog_self.accept
                )

                layout.addWidget(ok_btn)

                def preview_update():

                    dialog_self.parent_widget.settings = (
                        dialog_self.settings.copy()
                    )

                    dialog_self.parent_widget.apply_settings()

        dialog = SettingsDialog(
            self,
            self.settings
        )

        if dialog.exec():

            self.settings = dialog.settings

            self.apply_settings()

            self.save_settings()

    # =====================
    # 時計更新
    # =====================

    def start_timer(self):

        self.timer = QTimer()

        self.timer.timeout.connect(
            self.update_clock
        )

        self.timer.start(1000)

        self.update_clock()

    def update_clock(self):

        now = datetime.now(JST)

        time_str = now.strftime("%H%M")

        for i, digit in enumerate(
            self.digits
        ):

            digit.set_value(
                time_str[i]
            )

    # =====================
    # 設定適用
    # =====================

    def apply_settings(self):

        self.setStyleSheet(f"""
        background-color:
        {self.settings["window_color"]};

        border-radius:20px;
        """)

        self.setWindowOpacity(
            self.settings["window_opacity"]
        )

        width_size = int(
            self.width() * 0.13
        )

        height_size = int(
            self.height() * 0.45
        )

        dynamic_font_size = min(
            width_size,
            height_size
        )

        for digit in self.digits:

            digit.apply_style(
                self.settings,
                dynamic_font_size
            )

        colon_font_size = int(
            dynamic_font_size * 0.8
        )

        colon_font = QFont(
            self.settings["font_family"],
            colon_font_size
        )

        colon_font.setBold(True)

        self.colon.setFont(
            colon_font
        )

        self.colon.setStyleSheet(f"""
        color:
        {self.settings["font_color"]};

        background:
        transparent;
        """)

    # =====================
    # 設定保存
    # =====================

    def load_settings(self):

        settings = (
            DEFAULT_SETTINGS.copy()
        )

        if SETTINGS_FILE.exists():

            try:

                with open(
                    SETTINGS_FILE,
                    "r",
                    encoding="utf-8"
                ) as f:

                    loaded = json.load(f)

                    settings.update(loaded)

            except:

                pass

        return settings

    def save_settings(self):

        self.settings["x"] = self.x()

        self.settings["y"] = self.y()

        self.settings["window_width"] = (
            self.width()
        )

        self.settings["window_height"] = (
            self.height()
        )

        try:

            with open(
                SETTINGS_FILE,
                "w",
                encoding="utf-8"
            ) as f:

                json.dump(
                    self.settings,
                    f,
                    indent=4
                )

        except:

            pass

    # =====================
    # リサイズ
    # =====================

    def resizeEvent(self, event):

        self.apply_settings()

        super().resizeEvent(event)

    def moveEvent(self, event):

        QTimer.singleShot(
            300,
            self.save_settings
        )

        super().moveEvent(event)

    def closeEvent(self, event):

        self.save_settings()

        super().closeEvent(event)

    # =====================
    # 右クリックメニュー
    # =====================

    def contextMenuEvent(self, event):

        menu = QMenu(self)

        settings_action = QAction(
            "設定",
            self
        )

        startup_action = QAction(
            "スタートアップ登録",
            self
        )

        menu.addAction(
            settings_action
        )

        menu.addSeparator()

        menu.addAction(
            startup_action
        )

        settings_action.triggered.connect(
            self.open_settings
        )

        startup_action.triggered.connect(
            self.create_startup_shortcut
        )

        menu.exec(
            event.globalPos()
        )

    # =====================
    # 自動起動
    # =====================

    def create_startup_shortcut(self):

        from win32com.client import Dispatch

        startup = os.path.join(
            os.environ["APPDATA"],
            r"Microsoft\Windows\Start Menu\Programs\Startup"
        )

        bat_path = os.path.join(
            startup,
            "start_clock.bat"
        )

        exe_path = os.path.abspath(
            sys.executable
        )

        bat_content = f'''@echo off
timeout /t 5 /nobreak > nul
start "" "{exe_path}"
'''

        with open(
            bat_path,
            "w",
            encoding="utf-8"
        ) as f:

            f.write(bat_content)

    # =====================
    # 起動
    # =====================


if __name__ == "__main__":

    boot_log("before QApplication")

    app = QApplication(sys.argv)

    boot_log("after QApplication")

    window = ClockWidget()

    boot_log("after ClockWidget")

    window.show()

    boot_log("after show")

    sys.exit(app.exec())
import sys
import os
import json

from pathlib import Path
from datetime import datetime, timezone, timedelta

JST = timezone(timedelta(hours=9))

import winshell
from win32com.client import Dispatch

from PyQt6.QtWidgets import (
    QApplication,
    QWidget,
    QMenu,
    QColorDialog,
    QFontDialog,
    QHBoxLayout,
    QLabel,
    QVBoxLayout,
    QPushButton,
    QDialog,
    QSlider,
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
# 定数
# =========================

APP_DIR = Path(__file__).parent

APP_DATA_DIR = Path(os.environ.get("APPDATA", Path.home())) / "Clock"
SETTINGS_FILE = APP_DATA_DIR / "settings.json"
LEGACY_SETTINGS_FILE = APP_DIR / "settings.json"

DEFAULT_SETTINGS = {

    # フォント
    "font_family": "Arial",

    # 色
    "font_color": "#FFFFFF",
    "background_color": "#000000",
    "window_color": "#111111",

    # 設定ウィンドウ色
    "settings_window_color": "#2B2B2B",

    # 透明度
    "window_opacity": 0.95,

    # サイズ
    "window_width": 700,
    "window_height": 300,

    # 位置
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



    def set_value(self, value):

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

        self.update()

    def paintEvent(self, event):

        painter = QPainter(self)

        painter.setRenderHint(
            QPainter.RenderHint.Antialiasing
        )

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
        font = QFont(
            self.font_family,
            self.font_size
        )

        font.setBold(True)

        painter.setFont(font)

        painter.setPen(self.text_color)

        painter.drawText(
            rect,
            Qt.AlignmentFlag.AlignCenter,
            self.value
        )


# =========================
# 設定ウィンドウ
# =========================

class SettingsDialog(QDialog):

    def __init__(self, parent, settings):

        super().__init__(parent)

        self.parent_widget = parent

        self.settings = settings.copy()

        self.setWindowTitle("時計設定")

        self.setFixedWidth(320)

        self.setWindowFlags(
            self.windowFlags()
            | Qt.WindowType.WindowStaysOnTopHint
        )

        self.update_dialog_style()

        layout = QVBoxLayout()

        layout.setSpacing(15)

        layout.setContentsMargins(
            20,
            20,
            20,
            20
        )

        self.setLayout(layout)

        # =====================
        # ボタン
        # =====================

        font_btn = QPushButton("フォント変更")
        font_btn.clicked.connect(
            self.change_font
        )
        layout.addWidget(font_btn)

        text_btn = QPushButton("文字色変更")
        text_btn.clicked.connect(
            self.change_text_color
        )
        layout.addWidget(text_btn)

        bg_btn = QPushButton("数字背景色変更")
        bg_btn.clicked.connect(
            self.change_bg_color
        )
        layout.addWidget(bg_btn)

        window_btn = QPushButton("時計ウィンドウ色変更")
        window_btn.clicked.connect(
            self.change_window_color
        )
        layout.addWidget(window_btn)

        settings_btn = QPushButton("設定ウィンドウ色変更")
        settings_btn.clicked.connect(
            self.change_settings_window_color
        )
        layout.addWidget(settings_btn)

        # =====================
        # 透明度
        # =====================

        opacity_label = QLabel("透明度")
        opacity_label.setStyleSheet(
            "color:white; background:transparent;"
        )

        layout.addWidget(opacity_label)

        self.opacity_slider = QSlider(
            Qt.Orientation.Horizontal
        )

        self.opacity_slider.setMinimum(30)

        self.opacity_slider.setMaximum(100)

        self.opacity_slider.setValue(
            int(
                self.settings["window_opacity"]
                * 100
            )
        )

        self.opacity_slider.valueChanged.connect(
            self.change_opacity
        )

        layout.addWidget(
            self.opacity_slider
        )

        # =====================
        # リセット
        # =====================

        reset_btn = QPushButton(
            "デフォルトに戻す"
        )

        reset_btn.clicked.connect(
            self.reset_settings
        )

        layout.addWidget(reset_btn)

        # =====================
        # OK
        # =====================

        ok_btn = QPushButton("OK")

        ok_btn.clicked.connect(
            self.accept
        )

        layout.addWidget(ok_btn)

    # =====================
    # スタイル
    # =====================

    def update_dialog_style(self):

        bg = self.settings[
            "settings_window_color"
        ]

        self.setStyleSheet(f"""
        QDialog {{
            background-color: {bg};
            border-radius: 20px;
        }}

        QPushButton {{
            background-color: #3A3A3A;
            color: white;
            border: none;
            border-radius: 12px;
            padding: 12px;
            font-size: 14px;
        }}

        QPushButton:hover {{
            background-color: #505050;
        }}

        QLabel {{
            color: white;
            background: transparent;
        }}

        QSlider {{
            background: transparent;
        }}
        """)

    # =====================
    # プレビュー
    # =====================

    def preview_update(self):

        self.parent_widget.settings = (
            self.settings.copy()
        )

        self.parent_widget.apply_settings()

        self.update_dialog_style()

    # =====================
    # リセット
    # =====================

    def reset_settings(self):

        self.settings = (
            DEFAULT_SETTINGS.copy()
        )

        self.preview_update()

    # =====================
    # フォント
    # =====================

    def change_font(self):

        current_font = QFont(
            self.settings["font_family"],
            12
        )

        dialog = QFontDialog(
            current_font,
            self
        )

        # リアルタイム反映
        dialog.currentFontChanged.connect(
            self.preview_font_change
        )

        if dialog.exec():

            font = dialog.selectedFont()

            self.settings["font_family"] = (
                font.family()
            )

            self.preview_update()

    def preview_font_change(self, font):

        self.settings["font_family"] = (
            font.family()
        )

        self.preview_update()

    # =====================
    # 文字色
    # =====================

    def change_text_color(self):

        current_color = QColor(
            self.settings["font_color"]
        )

        dialog = QColorDialog(
            current_color,
            self
        )

        dialog.currentColorChanged.connect(
            self.preview_text_color
        )

        if dialog.exec():

            color = dialog.selectedColor()

            self.settings["font_color"] = (
                color.name()
            )

            self.preview_update()

    def preview_text_color(self, color):

        self.settings["font_color"] = (
            color.name()
        )

        self.preview_update()

    # =====================
    # 数字背景色
    # =====================

    def change_bg_color(self):

        current_color = QColor(
            self.settings["background_color"]
        )

        dialog = QColorDialog(
            current_color,
            self
        )

        dialog.currentColorChanged.connect(
            self.preview_bg_color
        )

        if dialog.exec():

            color = dialog.selectedColor()

            self.settings["background_color"] = (
                color.name()
            )

            self.preview_update()

    def preview_bg_color(self, color):

        self.settings["background_color"] = (
            color.name()
        )

        self.preview_update()

    # =====================
    # 時計背景色
    # =====================

    def change_window_color(self):

        current_color = QColor(
            self.settings["window_color"]
        )

        dialog = QColorDialog(
            current_color,
            self
        )

        dialog.currentColorChanged.connect(
            self.preview_window_color
        )

        if dialog.exec():

            color = dialog.selectedColor()

            self.settings["window_color"] = (
                color.name()
            )

            self.preview_update()

    def preview_window_color(self, color):

        self.settings["window_color"] = (
            color.name()
        )

        self.preview_update()

    # =====================
    # 設定背景色
    # =====================

    def change_settings_window_color(self):

        current_color = QColor(
            self.settings[
                "settings_window_color"
            ]
        )

        dialog = QColorDialog(
            current_color,
            self
        )

        dialog.currentColorChanged.connect(
            self.preview_settings_window_color
        )

        if dialog.exec():

            color = dialog.selectedColor()

            self.settings[
                "settings_window_color"
            ] = color.name()

            self.preview_update()

    def preview_settings_window_color(
        self,
        color
    ):

        self.settings[
            "settings_window_color"
        ] = color.name()

        self.preview_update()

    # =====================
    # 透明度
    # =====================

    def change_opacity(self, value):

        self.settings["window_opacity"] = (
            value / 100
        )

        self.preview_update()


# =========================
# メイン時計
# =========================

class ClockWidget(QWidget):

    def __init__(self):

        super().__init__()

        self.settings = self.load_settings()

        self.setup_window()

        self.create_ui()

        self.apply_settings()

        self.start_timer()

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

        # 全体余白
        layout.setContentsMargins(
            15,
            15,
            15,
            15
        )

        # 数字間隔
        DIGIT_SPACING = 10

        self.setLayout(layout)

        self.digits = []

        # =====================
        # 時
        # =====================

        for i in range(2):

            digit = FlipDigit()

            layout.addWidget(digit)

            self.digits.append(digit)

            if i == 0:
                layout.addSpacing(DIGIT_SPACING)

        # =====================
        # コロン
        # =====================

        # 左余白
        layout.addSpacing(18)

        self.colon = QLabel(":")

        self.colon.setAlignment(
            Qt.AlignmentFlag.AlignCenter
        )

        # コロン幅
        self.colon.setFixedWidth(35)

        layout.addWidget(self.colon)

        # 右余白
        layout.addSpacing(18)

        # =====================
        # 分
        # =====================

        for i in range(2):

            digit = FlipDigit()

            layout.addWidget(digit)

            self.digits.append(digit)

            if i == 0:
                layout.addSpacing(DIGIT_SPACING)
    # =====================
    # 設定画面
    # =====================

    def open_settings(self):

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

        # 時計ウィンドウ色
        self.setStyleSheet(f"""
        background-color:
        {self.settings["window_color"]};

        border-radius:20px;
        """)

        # 透明度
        self.setWindowOpacity(
            self.settings["window_opacity"]
        )

        # 自動文字サイズ
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

        # 数字
        for digit in self.digits:

            digit.apply_style(
                self.settings,
                dynamic_font_size
            )

        # :
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
        settings = DEFAULT_SETTINGS.copy()
        source = SETTINGS_FILE if SETTINGS_FILE.exists() else LEGACY_SETTINGS_FILE
        if source.exists():
            try:
                with source.open("r", encoding="utf-8") as f:
                    loaded = json.load(f)
                if isinstance(loaded, dict):
                    settings.update(loaded)
            except (OSError, UnicodeError, json.JSONDecodeError):
                pass

        # Only accept values with the types and ranges used by the UI.
        for key in ("font_family",):
            if not isinstance(settings.get(key), str) or not settings[key].strip():
                settings[key] = DEFAULT_SETTINGS[key]
        for key in ("font_color", "background_color", "window_color", "settings_window_color"):
            color = QColor(str(settings.get(key, "")))
            settings[key] = color.name() if color.isValid() else DEFAULT_SETTINGS[key]
        try:
            opacity = float(settings["window_opacity"])
            settings["window_opacity"] = min(1.0, max(0.3, opacity))
        except (TypeError, ValueError):
            settings["window_opacity"] = DEFAULT_SETTINGS["window_opacity"]
        for key, minimum, maximum in (("window_width", 300, 4000), ("window_height", 150, 3000), ("x", -10000, 10000), ("y", -10000, 10000)):
            try:
                settings[key] = min(maximum, max(minimum, int(settings[key])))
            except (TypeError, ValueError):
                settings[key] = DEFAULT_SETTINGS[key]
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

        APP_DATA_DIR.mkdir(parents=True, exist_ok=True)
        temporary_file = SETTINGS_FILE.with_suffix(".json.tmp")
        with temporary_file.open("w", encoding="utf-8") as f:
            json.dump(self.settings, f, indent=4, ensure_ascii=False)
            f.flush()
            os.fsync(f.fileno())
        temporary_file.replace(SETTINGS_FILE)

    # =====================
    # リサイズ
    # =====================

    def resizeEvent(self, event):

        self.apply_settings()

        self.save_settings()

        super().resizeEvent(event)

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

        startup = winshell.startup()

        shortcut_path = os.path.join(
            startup,
            "DesktopClock.lnk"
        )

        target = sys.executable

        shell = Dispatch(
            "WScript.Shell"
        )

        shortcut = shell.CreateShortCut(
            shortcut_path
        )

        shortcut.Targetpath = target

        if not getattr(sys, "frozen", False):
            shortcut.Arguments = f'"{Path(__file__).resolve()}"'

        shortcut.WorkingDirectory = str(
            APP_DIR
        )

        shortcut.save()

        print("Startup登録完了")


    # =====================
    # 終了
    # =====================

    def closeEvent(self, event):

        self.save_settings()

        event.accept()


# =========================
# 起動
# =========================

if __name__ == "__main__":

    app = QApplication(sys.argv)

    window = ClockWidget()

    window.show()

    sys.exit(app.exec())

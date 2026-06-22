"""
analyze_dialogue_android.py

Android GUI (Kivy) version of the dialogue analysis app.
Uses raw HTTP (requests) instead of the anthropic SDK to avoid
pydantic-core's compiled Rust extension, which cannot be cross-compiled
by python-for-android.

Build APK:
    pip install buildozer
    buildozer android debug
"""

import json
import os
import threading
from datetime import datetime

from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.scrollview import ScrollView
from kivy.uix.label import Label
from kivy.uix.textinput import TextInput
from kivy.uix.button import Button
from kivy.uix.popup import Popup
from kivy.clock import Clock
from kivy.lang import Builder
from kivy.metrics import dp
import requests
from openpyxl import load_workbook

from dialogue_core import read_conversation, build_prompt, apply_formatting

# Android-specific imports (available only inside a packaged APK)
try:
    from android.permissions import request_permissions, Permission
    from android.storage import primary_external_storage_path
    ANDROID = True
except ImportError:
    ANDROID = False

try:
    from plyer import filechooser
    PLYER_AVAILABLE = True
except ImportError:
    PLYER_AVAILABLE = False


# ---------------------------------------------------------------------------
# Claude API call via raw HTTP (no pydantic dependency)
# ---------------------------------------------------------------------------

def call_claude_http(api_key: str, conversation_rows: list[dict]) -> list[dict]:
    prompt = build_prompt(conversation_rows)

    response = requests.post(
        "https://api.anthropic.com/v1/messages",
        headers={
            "x-api-key": api_key,
            "anthropic-version": "2023-06-01",
            "anthropic-beta": "interleaved-thinking-2025-05-14",
            "content-type": "application/json",
        },
        json={
            "model": "claude-opus-4-8",
            "max_tokens": 4096,
            "thinking": {"type": "adaptive"},
            "messages": [{"role": "user", "content": prompt}],
        },
        timeout=120,
    )

    if response.status_code != 200:
        raise ValueError(f"API エラー {response.status_code}: {response.text[:300]}")

    data = response.json()
    text_content = ""
    for block in data.get("content", []):
        if block.get("type") == "text":
            text_content = block.get("text", "")
            break

    start = text_content.find("{")
    end = text_content.rfind("}") + 1
    if start == -1 or end == 0:
        raise ValueError(f"Claude の応答に JSON が含まれていませんでした:\n{text_content}")

    result = json.loads(text_content[start:end])
    return result.get("selections", [])


# ---------------------------------------------------------------------------
# KV layout
# ---------------------------------------------------------------------------

KV = """
<MainLayout>:
    orientation: 'vertical'
    padding: dp(12)
    spacing: dp(8)

    Label:
        text: '対話分析 - 垂直的な深まり検出'
        size_hint_y: None
        height: dp(44)
        bold: True
        font_size: dp(17)

    BoxLayout:
        size_hint_y: None
        height: dp(48)
        spacing: dp(8)
        Label:
            text: 'API キー:'
            size_hint_x: None
            width: dp(90)
            halign: 'right'
            valign: 'middle'
            text_size: self.size
        TextInput:
            id: api_key_input
            password: True
            multiline: False
            hint_text: 'sk-ant-...'
            size_hint_y: None
            height: dp(44)

    BoxLayout:
        size_hint_y: None
        height: dp(48)
        spacing: dp(8)
        Button:
            text: 'ファイルを選択'
            size_hint_x: None
            width: dp(150)
            on_press: app.open_file()
        Label:
            id: file_label
            text: '（未選択）'
            text_size: self.size
            halign: 'left'
            valign: 'middle'

    BoxLayout:
        size_hint_y: None
        height: dp(52)
        spacing: dp(10)
        Button:
            id: run_btn
            text: '分析を実行する'
            disabled: True
            on_press: app.run_analysis()
        Button:
            id: save_btn
            text: '結果を保存'
            disabled: True
            on_press: app.save_file()

    Label:
        text: 'ログ:'
        size_hint_y: None
        height: dp(26)
        halign: 'left'
        text_size: self.size
        valign: 'middle'

    ScrollView:
        id: scroll_view
        Label:
            id: log_label
            text: 'API キーを入力してファイルを選択してください。'
            text_size: self.width, None
            size_hint_y: None
            height: self.texture_size[1]
            halign: 'left'
            valign: 'top'
            padding: dp(4), dp(4)
"""


class MainLayout(BoxLayout):
    pass


# ---------------------------------------------------------------------------
# Application
# ---------------------------------------------------------------------------

class DialogueAnalysisApp(App):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self._input_path = ""
        self._workbook = None

    def build(self):
        Builder.load_string(KV)
        self.layout = MainLayout()

        if ANDROID:
            request_permissions([
                Permission.READ_EXTERNAL_STORAGE,
                Permission.WRITE_EXTERNAL_STORAGE,
                Permission.INTERNET,
            ])

        return self.layout

    # ------------------------------------------------------------------
    # File selection
    # ------------------------------------------------------------------

    def open_file(self):
        if PLYER_AVAILABLE:
            filechooser.open_file(
                on_selection=self._on_file_selected,
                filters=["*.xlsx", "*.xlsm", "*.xls"],
                title="Excel ファイルを選択",
            )
        else:
            self._show_popup("エラー",
                             "ファイル選択が利用できません。\n"
                             "plyer パッケージが必要です。")

    def _on_file_selected(self, selection):
        if not selection:
            return
        path = selection[0]
        self._input_path = path
        self._workbook = None
        self.layout.ids.save_btn.disabled = True
        self.layout.ids.file_label.text = os.path.basename(path)
        self.layout.ids.run_btn.disabled = False
        self._log_clear()
        self._log(f"入力ファイル: {path}\n")

    # ------------------------------------------------------------------
    # Analysis
    # ------------------------------------------------------------------

    def run_analysis(self):
        api_key = self.layout.ids.api_key_input.text.strip()
        if not api_key:
            self._show_popup("APIキー未入力",
                             "Anthropic API キーを入力してください。\n"
                             "https://console.anthropic.com/ で取得できます。")
            return
        if not self._input_path:
            self._show_popup("警告", "先にファイルを選択してください。")
            return

        self.layout.ids.run_btn.disabled = True
        self.layout.ids.save_btn.disabled = True
        threading.Thread(target=self._analysis_worker,
                         args=(api_key,), daemon=True).start()

    def _analysis_worker(self, api_key: str):
        try:
            self._log("Excel ファイルを読み込んでいます...\n")
            wb = load_workbook(self._input_path)
            ws = wb.worksheets[0]

            self._log("対話データを解析しています...\n")
            rows = read_conversation(ws)
            if not rows:
                Clock.schedule_once(lambda dt: self._on_error(
                    "対話データが見つかりませんでした（A列・B列が空です）。"))
                return
            self._log(f"  {len(rows)} 行の対話を読み込みました\n\n")

            self._log("Claude に分析させています...\n（数十秒かかる場合があります）\n")
            selections = call_claude_http(api_key, rows)
            self._log(f"  {len(selections)} か所の垂直的な深まりを特定しました\n\n")

            self._log("セルに書式を適用しています...\n")
            for line in apply_formatting(ws, selections):
                self._log(line + "\n")

            self._log("\n--- 分析結果 ---\n")
            for i, sel in enumerate(selections, 1):
                self._log(
                    f"\n[{i}] 行{sel.get('row')} 列{sel.get('column')}\n"
                    f"    {sel.get('comment', '')}\n"
                )

            self._workbook = wb
            Clock.schedule_once(lambda dt: self._on_success())

        except Exception as exc:
            msg = str(exc)
            Clock.schedule_once(lambda dt: self._on_error(msg))

    # ------------------------------------------------------------------
    # Save
    # ------------------------------------------------------------------

    def save_file(self):
        if self._workbook is None:
            self._show_popup("警告", "先に分析を実行してください。")
            return

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"result_{timestamp}.xlsx"

        if ANDROID:
            try:
                downloads = os.path.join(primary_external_storage_path(), "Download")
            except Exception:
                downloads = "/sdcard/Download"
        else:
            downloads = os.path.expanduser("~/Downloads")

        os.makedirs(downloads, exist_ok=True)
        save_path = os.path.join(downloads, filename)

        try:
            self._workbook.save(save_path)
            self._log(f"\n保存しました: {save_path}\n")
            self._show_popup("保存完了", f"保存しました:\n{save_path}")
        except Exception as exc:
            self._show_popup("保存エラー", str(exc))

    # ------------------------------------------------------------------
    # Callbacks
    # ------------------------------------------------------------------

    def _on_success(self):
        self.layout.ids.run_btn.disabled = False
        self.layout.ids.save_btn.disabled = False

    def _on_error(self, message: str):
        self.layout.ids.run_btn.disabled = False
        self._log(f"\n[エラー] {message}\n")
        self._show_popup("エラー", message)

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------

    def _log(self, text: str):
        def _do(dt):
            label = self.layout.ids.log_label
            label.text += text
            self.layout.ids.scroll_view.scroll_y = 0
        Clock.schedule_once(_do)

    def _log_clear(self):
        def _do(dt):
            self.layout.ids.log_label.text = ""
        Clock.schedule_once(_do)

    def _show_popup(self, title: str, message: str):
        def _do(dt):
            content = BoxLayout(orientation="vertical", padding=dp(12), spacing=dp(10))
            content.add_widget(Label(
                text=message,
                text_size=(dp(320), None),
                size_hint_y=None,
                height=dp(90),
                halign="center",
            ))
            btn = Button(text="OK", size_hint_y=None, height=dp(48))
            content.add_widget(btn)
            popup = Popup(title=title, content=content,
                          size_hint=(0.88, None), height=dp(220))
            btn.bind(on_press=popup.dismiss)
            popup.open()
        Clock.schedule_once(_do)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    DialogueAnalysisApp().run()


if __name__ == "__main__":
    main()

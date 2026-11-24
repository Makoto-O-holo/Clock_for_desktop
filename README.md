# Clock for desktop

Windows デスクトップに日本時間（JST）の時計を表示するアプリです。時計を右クリックすると、表示設定とスタートアップ登録を選べます。

## 機能

- 常に手前に表示する時計ウィンドウ
- フォント、文字色、背景色、透明度の変更
- ウィンドウ位置とサイズの保存
- ユーザーが選んだ場合のみ Windows スタートアップへ登録

## 動作環境

- Windows 10 / 11
- Python 3.10 以降（ソースから実行する場合）

## ソースから実行

```powershell
git clone https://github.com/mkt6006/Clock.git
cd Clock
py -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python main.py
```

右クリックメニューの「スタートアップ登録」を選ぶと、現在の実行形式に合わせたショートカットを Windows の Startup フォルダーに作ります。解除するには Startup フォルダーから `DesktopClock.lnk` を削除してください。

## 設定とプライバシー

設定は `%APPDATA%\Clock\settings.json` に保存します。画面位置などの設定値はこのPC内だけで使い、ネットワーク送信しません。初回起動時に、旧版の隣接 `settings.json` があれば読み込みます。公開リポジトリには個人の画面位置を含めていません。

このアプリはネットワーク通信、更新確認、テレメトリーを行いません。スタートアップ登録はメニューから明示的に選んだときだけ行います。

## 開発・ビルド

依存パッケージは [`requirements.txt`](requirements.txt) を参照してください。Windows 上で実行ファイルを作る場合は Nuitka を追加し、次を実行します。

```powershell
python -m pip install nuitka
python -m nuitka --standalone --enable-plugin=pyqt6 --windows-console-mode=disable main.py
```

ビルド出力とユーザー設定は Git 管理対象外です。

## ライセンス

このリポジトリのアプリケーションコードには、現時点で再配布・改変を許諾するライセンスを付与していません。公開は閲覧目的で、権利者の許可なく複製・改変・再配布できることを意味しません。依存ライブラリはそれぞれのライセンスに従います。

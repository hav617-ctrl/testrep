@echo off
REM Windows 上でこのバッチを実行すると xlsm_to_xlsx.exe が dist\ に生成されます
REM 事前に Python 3.x (64bit) をインストールし PATH を通しておいてください

echo [1/3] 依存ライブラリをインストール中...
pip install -r requirements.txt

echo [2/3] PyInstaller でビルド中...
pyinstaller xlsm_to_xlsx.spec --clean

echo [3/3] 完了
echo 実行ファイル: dist\xlsm_to_xlsx.exe
pause

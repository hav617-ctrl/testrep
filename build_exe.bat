@echo off
echo [1/3] Installing dependencies...
pip install -r requirements.txt

echo [2/3] Building exe...
pyinstaller xlsm_to_xlsx.spec --clean

echo [3/3] Done!
echo Output: dist\xlsm_to_xlsx.exe
pause

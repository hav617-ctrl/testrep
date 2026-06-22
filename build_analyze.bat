@echo off
echo [1/3] Installing dependencies...
pip install -r requirements.txt

echo [2/3] Building exe...
python -m PyInstaller analyze_dialogue.spec --clean

echo [3/3] Done!
echo Output: dist\analyze_dialogue.exe
pause

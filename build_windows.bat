@echo off
echo ============================================
echo  Build LuanChuyenHangHoa.exe
echo ============================================

:: Install dependencies
pip install -r requirements.txt

:: Build single-file exe
pyinstaller app.spec --clean --noconfirm

echo.
echo Done! File exe nam tai: dist\LuanChuyenHangHoa.exe
pause

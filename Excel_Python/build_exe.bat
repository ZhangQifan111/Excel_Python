@echo off
cd /d "%~dp0"
echo Current directory: %cd%
echo.
echo Installing PyInstaller...
pip install pyinstaller -i https://pypi.tuna.tsinghua.edu.cn/simple
echo.
echo Building EXE...
pyinstaller --onefile --windowed excel_data_visualizer.py
echo.
echo Done! Check dist folder for EXE.
pause
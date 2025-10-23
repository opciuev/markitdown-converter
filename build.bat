@echo off
echo Installing dependencies...
pip install PySide6 PyInstaller markitdown[all] openpyxl

echo Cleaning old builds...
if exist "dist" rmdir /s /q "dist"
if exist "build" rmdir /s /q "build"

echo Building executable with spec file...
pyinstaller markitdown_ui.spec

echo Build complete!
echo Executable location: dist\MarkItDownConverter.exe
echo.
echo Note: This version will show a console window for debugging.
echo If everything works, you can edit markitdown_ui.spec and change console=True to console=False
pause

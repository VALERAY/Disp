@echo off
echo Installing dependencies...
.venv\Scripts\python.exe -m pip install -r requirements.txt
echo.
echo Done! Press any key to exit...
pause >nul


@echo off
REM ================================================
REM Install Python dependencies (no virtual env)
REM ================================================

echo Checking Python installation...
python --version || (echo Python not found! & exit /b 1)

echo Upgrading pip...
python -m pip install --upgrade pip

echo Installing dependencies from requirements.txt...
pip install -r requirements.txt

echo.
echo ================================================
echo All dependencies installed successfully.
echo ================================================
pause
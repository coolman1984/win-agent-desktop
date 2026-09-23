@echo off
REM Excel demo: an agent builds and verifies a Sales Command Center, hands-free.
REM Opens its OWN Excel process - your open workbooks are not touched.
REM Keep hands off the mouse and keyboard for about 2 minutes while it runs.
cd /d "%~dp0"
python demo_excel.py
echo.
pause

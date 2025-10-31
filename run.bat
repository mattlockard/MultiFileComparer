@echo off
REM Create virtual environment if it doesn't exist
if not exist venv (
    python -m venv venv
)

REM Activate the virtual environment
call venv\Scripts\activate

REM Run the FileComparer script with any arguments
python FileComparer.py %*

pause
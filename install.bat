@echo off

:: Read PYTHON variable from environment.env
for /F "tokens=1* delims==" %%A in (environment.env) do (
    if /I "%%A"=="PYTHON" (
        set PYTHON_VERSION=%%B
    )
)

echo Creating a virtual environment with Python version %PYTHON_VERSION%
%PYTHON_VERSION% -m venv env

echo Activating the virtual environment
call env\Scripts\activate

echo Installing dependencies from req.txt
pip install -r req.txt

echo Done.
pause

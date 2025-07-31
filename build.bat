@echo off
chcp 65001 > nul
echo Generating Windows executable with Nuitka...

REM Activate virtual environment if needed (adjust path)
call .\env\Scripts\activate.bat

REM Run Nuitka command
python -m nuitka ^
    --output-dir=build ^
    --onefile ^
    --standalone ^
    --windows-console-mode=force ^
    --output-filename=AccessDevKit.exe ^
    --include-package=typer ^
    --include-package=rich ^
    --include-package=pyodbc ^
    --include-package=openpyxl ^
    --include-package=adodbapi ^
    --include-package=fire ^
    --include-package=termcolor ^
    --include-data-dir=src\templates=src\templates ^
    src/main.py

IF %ERRORLEVEL% NEQ 0 (
    echo.
    echo An error occurred.
    echo.
    pause
    exit /b %ERRORLEVEL%
)

echo.
echo Executable generation completed.
echo.
pause
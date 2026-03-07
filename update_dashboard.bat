@echo off
echo Updating Dashboard Data from Excel...
python convert_data.py
if %ERRORLEVEL% EQU 0 (
    echo.
    echo Update Successful! Refresh your dashboard.
) else (
    echo.
    echo Update Failed. Please check the error messages above.
)
pause

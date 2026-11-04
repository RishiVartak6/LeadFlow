@echo off
cd /d "%~dp0"
title Push LeadFlow to GitHub
echo ==============================================
echo   Pushing LeadFlow to GitHub
echo   Repository: https://github.com/RishiVartak6/LeadFlow.git
echo ==============================================
echo.
git push -u origin main
echo.
if %ERRORLEVEL% equ 0 (
    echo [SUCCESS] Successfully pushed to GitHub!
) else (
    echo [ERROR] Push failed. If prompted, please complete GitHub sign-in above.
)
echo.
pause

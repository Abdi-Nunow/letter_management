@echo off
title Deploying Letter Management System to Live Firebase
cd /d "C:\Users\abdin\.gemini\antigravity\scratch\letter-management-f1f19"
echo ========================================================
echo   Deploying Somali Revenue Bureau DDS - LMS to Live
echo   Target: https://letter-management-f1f19.web.app/
echo ========================================================
echo.
echo Checking Firebase account...
echo Please ensure you are logged in with the Google Account that owns 'letter-management-f1f19'.
echo.
call firebase login:add
echo.
echo Deploying hosting files...
call firebase deploy --only hosting --project letter-management-f1f19
echo.
echo ========================================================
echo   Deployment Process Completed!
echo ========================================================
pause

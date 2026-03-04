@echo off
cd /d "%~dp0"

echo Adding all changes...
git add -A

echo.
set /p msg="Commit message (or press Enter for 'update'): "
if "%msg%"=="" set msg=update

git commit -m "%msg%"

echo.
echo Pushing to GitHub...
git push origin master

echo.
echo Done!
pause

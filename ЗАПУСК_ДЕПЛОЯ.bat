@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ========================================
echo  Shape Builder — заливка в GlobalMacros
echo  Папка: %~dp0
echo ========================================
echo.
echo 1) Запустите CorelDRAW, откройте документ.
echo 2) Один раз Alt+F11, закройте окно VBA.
echo 3) Нажмите любую клавишу — пойдет deploy_direct.ps1
echo.
pause >nul

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0deploy_direct.ps1"
set ERR=%ERRORLEVEL%
echo.
if %ERR% neq 0 (
  echo ОШИБКА деплоя, код %ERR%. Скопируйте текст выше.
) else (
  echo Готово. Откройте Alt+F11 - Файл - Сохранить GlobalMacros.
  echo В заголовке формы должно быть: v0.1.3 [2026-contour-bar]
)
echo.
pause

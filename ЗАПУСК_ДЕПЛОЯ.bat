@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ========================================
echo  Shape Builder — заливка в GlobalMacros (CorelDRAW 2026)
echo  Папка: %~dp0
echo ========================================
echo.
echo 1) Запустите CorelDRAW, откройте документ.
echo 2) Один раз Alt+F11, закройте окно VBA.
echo 3) Нажмите любую клавишу — пойдет deploy_all.ps1 (Corel+deploy, лог deploy_last_log.txt)
echo.
pause >nul

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0deploy_all.ps1"
set ERR=%ERRORLEVEL%
echo.
if %ERR% neq 0 (
  echo ОШИБКА деплоя, код %ERR%. Скопируйте текст выше.
) else (
  echo Готово. Откройте Alt+F11 - Файл - Сохранить GlobalMacros.
  echo В заголовке: v0.1.4 [2026-verify], под заголовком фасада - строка про 0.1.4 контур
  echo Запуск макроса: Сервис - Макросы - GlobalMacros - Module1.Фасады
)
echo.
pause

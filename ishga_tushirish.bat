@echo off
chcp 65001 >nul
cd /d "%~dp0"

set PY=
where py >nul 2>nul && set PY=py
if "%PY%"=="" ( where python >nul 2>nul && set PY=python )
if "%PY%"=="" (
    echo [XATO] Python topilmadi. https://www.python.org/downloads/ dan o'rnating
    echo        (o'rnatishda "Add python.exe to PATH" belgisini qo'ying^).
    pause
    exit /b 1
)

if not exist ".deps_ok" (
    echo Kerakli kutubxonalar o'rnatilmoqda (faqat birinchi marta^)...
    %PY% -m pip install --disable-pip-version-check -q -r requirements.txt
    if errorlevel 1 (
        echo [XATO] Kutubxonalarni o'rnatib bo'lmadi. Internet ulanishini tekshiring.
        pause
        exit /b 1
    )
    echo ok> .deps_ok
)

%PY% hisob_tahlil.py %*
if errorlevel 1 pause

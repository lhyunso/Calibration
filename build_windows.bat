@echo off
:: ============================================================
:: MCAL Windows Build Script
:: Python 3.11 및 pip 필요 (https://www.python.org/downloads/)
:: 실행: build_windows.bat
:: 결과: dist\MCAL\MCAL.exe
:: ============================================================

:: Python 3.11 버전 확인
python --version 2>&1 | findstr /R "3\.11\." >nul
if errorlevel 1 (
    echo 오류: Python 3.11이 필요합니다.
    echo https://www.python.org/downloads/ 에서 3.11.x 설치 후 재실행하세요.
    pause
    exit /b 1
)

echo [1/4] 의존성 설치 중...
pip install -r requirements.txt
if errorlevel 1 (
    echo 오류: pip install 실패
    pause
    exit /b 1
)

echo [2/4] 기존 빌드 정리 중...
if exist build rmdir /s /q build
if exist dist  rmdir /s /q dist

echo [3/4] PyInstaller 빌드 중...
pyinstaller MCAL.spec --clean --noconfirm
if errorlevel 1 (
    echo 오류: PyInstaller 빌드 실패
    pause
    exit /b 1
)

echo [4/4] 완료!
echo 실행 파일 위치: dist\MCAL\MCAL.exe
echo.
echo 배포 시 dist\MCAL\ 폴더 전체를 전달하세요.
pause

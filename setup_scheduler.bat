@echo off
chcp 65001 >nul
echo ===================================
echo  원부재료 자동동기화 스케줄러 등록
echo ===================================
echo.

:: 매일 오전 9시 실행
schtasks /create /tn "TheBreadBlue_IngredientSync" /tr "py \"%~dp0auto_sync_ingredients.py\"" /sc daily /st 09:00 /f

if %errorlevel% equ 0 (
    echo.
    echo [완료] 매일 오전 9시 자동 실행 등록됨
    echo  - 작업명: TheBreadBlue_IngredientSync
    echo  - 수동 실행: py auto_sync_ingredients.py
    echo  - 스케줄러 확인: taskschd.msc
) else (
    echo.
    echo [실패] 관리자 권한으로 다시 실행하세요
)

echo.
pause

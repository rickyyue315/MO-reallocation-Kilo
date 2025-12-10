@echo off
chcp 65001 > nul
title 庫存調貨建議系統 v1.0

echo.
echo ========================================
echo   庫存調貨建議系統 v1.0
echo ========================================
echo.

REM 檢查Python是否安裝
python --version > nul 2>&1
if %errorlevel% neq 0 (
    echo ❌ 錯誤: 未找到Python，請先安裝Python 3.8或更高版本
    echo.
    echo 下載地址: https://www.python.org/downloads/
    echo.
    pause
    exit /b 1
)

echo ✅ Python環境檢測通過
echo.

REM 檢查依賴是否安裝
echo 🔍 檢查系統依賴...
python -c "import pandas, streamlit, numpy, openpyxl, xlsxwriter, matplotlib, seaborn" > nul 2>&1
if %errorlevel% neq 0 (
    echo ⚠️  警告: 檢測到缺少必要的依賴包
    echo.
    echo 請選擇:
    echo 1. 自動安裝依賴 (推薦)
    echo 2. 手動安裝依賴
    echo 3. 繼續運行 (可能失敗)
    echo.
    set /p choice=請輸入選項 (1-3): 
    
    if "%choice%"=="1" (
        echo.
        echo 🚀 開始自動安裝依賴...
        python install_dependencies.py
        if %errorlevel% neq 0 (
            echo ❌ 依賴安裝失敗
            pause
            exit /b 1
        )
        echo ✅ 依賴安裝完成
    ) else if "%choice%"=="2" (
        echo.
        echo 📝 請手動運行以下命令安裝依賴:
        echo pip install pandas openpyxl streamlit numpy xlsxwriter matplotlib seaborn
        echo.
        pause
        exit /b 0
    ) else if "%choice%"=="3" (
        echo ⚠️  繼續運行可能會出現錯誤
    ) else (
        echo ❌ 無效選項
        pause
        exit /b 1
    )
) else (
    echo ✅ 依賴檢查通過
)

echo.
echo 🚀 啟動庫存調貨建議系統...
echo.

REM 啟動Streamlit應用
streamlit run app.py

REM 檢查運行結果
if %errorlevel% neq 0 (
    echo.
    echo ❌ 系統運行出現錯誤
    echo.
    echo 可能的解決方案:
    echo 1. 檢查Python版本是否為3.8或更高
    echo 2. 運行 'python install_dependencies.py' 安裝依賴
    echo 3. 檢查防火牆設置
    echo 4. 確認端口8501未被佔用
    echo.
    pause
    exit /b 1
)

echo.
echo 🎉 系統已正常關閉
pause
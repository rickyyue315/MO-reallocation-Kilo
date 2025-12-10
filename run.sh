#!/bin/bash

# 庫存調貨建議系統 v1.0 - Linux/macOS運行腳本

# 設置顏色
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

# 打印帶顏色的消息
print_message() {
    local color=$1
    local message=$2
    echo -e "${color}${message}${NC}"
}

print_header() {
    echo "========================================"
    echo "  庫存調貨建議系統 v1.0"
    echo "========================================"
    echo
}

# 檢查命令是否存在
command_exists() {
    command -v "$1" >/dev/null 2>&1
}

# 檢查Python版本
check_python() {
    if command_exists python3; then
        PYTHON_CMD="python3"
    elif command_exists python; then
        PYTHON_CMD="python"
    else
        print_message $RED "❌ 錯誤: 未找到Python，請先安裝Python 3.8或更高版本"
        echo
        echo "Ubuntu/Debian安裝命令:"
        echo "  sudo apt update"
        echo "  sudo apt install python3 python3-pip"
        echo
        echo "CentOS/RHEL安裝命令:"
        echo "  sudo yum install python3 python3-pip"
        echo
        echo "macOS安裝命令:"
        echo "  brew install python3"
        echo
        echo "下載地址: https://www.python.org/downloads/"
        exit 1
    fi
    
    # 檢查Python版本
    PYTHON_VERSION=$($PYTHON_CMD -c "import sys; print(f'{sys.version_info.major}.{sys.version_info.minor}')")
    REQUIRED_VERSION="3.8"
    
    if [ "$(printf '%s\n' "$REQUIRED_VERSION" "$PYTHON_VERSION" | sort -V | head -n1)" != "$REQUIRED_VERSION" ]; then
        print_message $RED "❌ 錯誤: Python版本過低，當前版本: $PYTHON_VERSION，需要: $REQUIRED_VERSION或更高"
        exit 1
    fi
    
    print_message $GREEN "✅ Python環境檢測通過 (版本: $PYTHON_VERSION)"
}

# 檢查依賴
check_dependencies() {
    echo
    print_message $BLUE "🔍 檢查系統依賴..."
    
    # 檢查pip
    if ! command_exists pip3 && ! command_exists pip; then
        print_message $YELLOW "⚠️  警告: 未找到pip，嘗試安裝..."
        if command_exists apt; then
            sudo apt install python3-pip -y
        elif command_exists yum; then
            sudo yum install python3-pip -y
        elif command_exists brew; then
            brew install python3
        else
            print_message $RED "❌ 無法自動安裝pip，請手動安裝"
            exit 1
        fi
    fi
    
    # 確定pip命令
    if command_exists pip3; then
        PIP_CMD="pip3"
    else
        PIP_CMD="pip"
    fi
    
    # 檢查Python依賴
    DEPS_CHECK=$($PYTHON_CMD -c "
import sys
try:
    import pandas, streamlit, numpy, openpyxl, xlsxwriter, matplotlib, seaborn
    print('OK')
except ImportError as e:
    print(f'MISSING: {e}')
    sys.exit(1)
" 2>&1)
    
    if [[ "$DEPS_CHECK" != "OK" ]]; then
        print_message $YELLOW "⚠️  警告: 檢測到缺少必要的依賴包"
        echo
        echo "請選擇:"
        echo "1. 自動安裝依賴 (推薦)"
        echo "2. 手動安裝依賴"
        echo "3. 繼續運行 (可能失敗)"
        echo
        read -p "請輸入選項 (1-3): " choice
        
        case $choice in
            1)
                echo
                print_message $BLUE "🚀 開始自動安裝依賴..."
                $PYTHON_CMD install_dependencies.py
                if [ $? -ne 0 ]; then
                    print_message $RED "❌ 依賴安裝失敗"
                    exit 1
                fi
                print_message $GREEN "✅ 依賴安裝完成"
                ;;
            2)
                echo
                print_message $BLUE "📝 請手動運行以下命令安裝依賴:"
                echo "$PIP_CMD install pandas openpyxl streamlit numpy xlsxwriter matplotlib seaborn"
                echo
                exit 0
                ;;
            3)
                print_message $YELLOW "⚠️  繼續運行可能會出現錯誤"
                ;;
            *)
                print_message $RED "❌ 無效選項"
                exit 1
                ;;
        esac
    else
        print_message $GREEN "✅ 依賴檢查通過"
    fi
}

# 啟動應用
start_application() {
    echo
    print_message $BLUE "🚀 啟動庫存調貨建議系統..."
    echo
    
    # 啟動Streamlit應用
    streamlit run app.py
    
    # 檢查運行結果
    if [ $? -ne 0 ]; then
        echo
        print_message $RED "❌ 系統運行出現錯誤"
        echo
        echo "可能的解決方案:"
        echo "1. 檢查Python版本是否為3.8或更高"
        echo "2. 運行 '$PYTHON_CMD install_dependencies.py' 安裝依賴"
        echo "3. 檢查防火牆設置"
        echo "4. 確認端口8501未被佔用"
        echo
        exit 1
    fi
    
    echo
    print_message $GREEN "🎉 系統已正常關閉"
}

# 主函數
main() {
    print_header
    check_python
    check_dependencies
    start_application
}

# 錯誤處理
trap 'echo; print_message $RED "❌ 腳本運行被中斷"; exit 1' INT TERM

# 運行主函數
main "$@"
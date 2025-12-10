"""
依賴安裝腳本
用於自動安裝庫存調貨建議系統v1.0所需的所有依賴包
"""

import subprocess
import sys
import os
from typing import List, Tuple

class DependencyInstaller:
    """依賴安裝器"""
    
    # 核心依賴包列表
    CORE_DEPENDENCIES = [
        "pandas>=1.5.0",
        "openpyxl>=3.0.10",
        "streamlit>=1.25.0",
        "numpy>=1.21.0",
        "xlsxwriter>=3.0.0",
        "matplotlib>=3.5.0",
        "seaborn>=0.11.0"
    ]
    
    # 可選依賴包列表
    OPTIONAL_DEPENDENCIES = [
        "plotly>=5.0.0",  # 用於高級圖表
        "altair>=4.0.0",  # 用於交互式圖表
        "scipy>=1.7.0",   # 用於高級數據分析
        "scikit-learn>=1.0.0"  # 用於機器學習功能
    ]
    
    def __init__(self):
        self.python_executable = sys.executable
        self.pip_command = [self.python_executable, "-m", "pip"]
    
    def run_command(self, command: List[str], description: str) -> Tuple[bool, str]:
        """
        運行命令並返回結果
        
        Args:
            command: 要運行的命令列表
            description: 命令描述
            
        Returns:
            (是否成功, 輸出信息)
        """
        print(f"🔄 {description}...")
        
        try:
            result = subprocess.run(
                command,
                capture_output=True,
                text=True,
                check=True
            )
            
            print(f"✅ {description}成功")
            return True, result.stdout
            
        except subprocess.CalledProcessError as e:
            print(f"❌ {description}失敗")
            print(f"錯誤信息: {e.stderr}")
            return False, e.stderr
    
    def upgrade_pip(self) -> bool:
        """
        升級pip到最新版本
        
        Returns:
            是否升級成功
        """
        success, _ = self.run_command(
            self.pip_command + ["install", "--upgrade", "pip"],
            "升級pip"
        )
        return success
    
    def install_dependencies(self, dependencies: List[str], description: str) -> bool:
        """
        安裝依賴包列表
        
        Args:
            dependencies: 依賴包列表
            description: 安裝描述
            
        Returns:
            是否安裝成功
        """
        success, _ = self.run_command(
            self.pip_command + ["install"] + dependencies,
            description
        )
        return success
    
    def install_with_mirror(self, dependencies: List[str], mirror_url: str) -> bool:
        """
        使用鏡像源安裝依賴包
        
        Args:
            dependencies: 依賴包列表
            mirror_url: 鏡像源URL
            
        Returns:
            是否安裝成功
        """
        success, _ = self.run_command(
            self.pip_command + ["install", "-i", mirror_url] + dependencies,
            f"使用鏡像源安裝依賴包"
        )
        return success
    
    def check_python_version(self) -> bool:
        """
        檢查Python版本是否滿足要求
        
        Returns:
            是否滿足要求
        """
        version = sys.version_info
        print(f"🐍 當前Python版本: {version.major}.{version.minor}.{version.micro}")
        
        if version.major < 3 or (version.major == 3 and version.minor < 8):
            print("❌ Python版本過低，需要Python 3.8或更高版本")
            return False
        
        print("✅ Python版本滿足要求")
        return True
    
    def check_installed_packages(self) -> List[str]:
        """
        檢查已安裝的包
        
        Returns:
            已安裝的包列表
        """
        try:
            import pkg_resources
            installed_packages = [d.project_name for d in pkg_resources.working_set]
            return installed_packages
        except ImportError:
            return []
    
    def get_package_version(self, package_name: str) -> str:
        """
        獲取已安裝包的版本
        
        Args:
            package_name: 包名
            
        Returns:
            包版本
        """
        try:
            import pkg_resources
            version = pkg_resources.get_distribution(package_name).version
            return version
        except:
            return "未安裝"
    
    def display_installation_status(self):
        """顯示安裝狀態"""
        print("\n📊 依賴包安裝狀態:")
        print("-" * 50)
        
        all_dependencies = self.CORE_DEPENDENCIES + self.OPTIONAL_DEPENDENCIES
        
        for dep in all_dependencies:
            package_name = dep.split(">=")[0].split("==")[0]
            version = self.get_package_version(package_name)
            status = "✅ 已安裝" if version != "未安裝" else "❌ 未安裝"
            core_marker = " [核心]" if dep in self.CORE_DEPENDENCIES else " [可選]"
            print(f"{package_name:<15} {version:<10} {status}{core_marker}")
        
        print("-" * 50)
    
    def install_core_dependencies(self, use_mirror: bool = False, mirror_url: str = None) -> bool:
        """
        安裝核心依賴包
        
        Args:
            use_mirror: 是否使用鏡像源
            mirror_url: 鏡像源URL
            
        Returns:
            是否安裝成功
        """
        print("\n🔧 開始安裝核心依賴包...")
        
        if use_mirror and mirror_url:
            success = self.install_with_mirror(self.CORE_DEPENDENCIES, mirror_url)
        else:
            success = self.install_dependencies(self.CORE_DEPENDENCIES, "安裝核心依賴包")
        
        return success
    
    def install_optional_dependencies(self, use_mirror: bool = False, mirror_url: str = None) -> bool:
        """
        安裝可選依賴包
        
        Args:
            use_mirror: 是否使用鏡像源
            mirror_url: 鏡像源URL
            
        Returns:
            是否安裝成功
        """
        print("\n🔧 開始安裝可選依賴包...")
        
        if use_mirror and mirror_url:
            success = self.install_with_mirror(self.OPTIONAL_DEPENDENCIES, mirror_url)
        else:
            success = self.install_dependencies(self.OPTIONAL_DEPENDENCIES, "安裝可選依賴包")
        
        return success
    
    def verify_installation(self) -> bool:
        """
        驗證安裝是否成功
        
        Returns:
            是否驗證成功
        """
        print("\n🔍 驗證安裝...")
        
        try:
            # 嘗試導入核心依賴
            import pandas
            import openpyxl
            import streamlit
            import numpy
            import xlsxwriter
            import matplotlib
            import seaborn
            
            print("✅ 所有核心依賴包驗證成功")
            return True
            
        except ImportError as e:
            print(f"❌ 依賴包驗證失敗: {e}")
            return False
    
    def create_requirements_txt(self):
        """創建requirements.txt文件"""
        print("\n📝 創建requirements.txt文件...")
        
        try:
            with open("requirements.txt", "w", encoding="utf-8") as f:
                f.write("# 庫存調貨建議系統v1.0依賴包\n")
                f.write("# 核心依賴\n")
                for dep in self.CORE_DEPENDENCIES:
                    f.write(f"{dep}\n")
                
                f.write("\n# 可選依賴\n")
                for dep in self.OPTIONAL_DEPENDENCIES:
                    f.write(f"#{dep}\n")
            
            print("✅ requirements.txt文件創建成功")
            return True
            
        except Exception as e:
            print(f"❌ 創建requirements.txt文件失敗: {e}")
            return False


def main():
    """主函數"""
    print("🚀 庫存調貨建議系統v1.0 - 依賴安裝腳本")
    print("=" * 50)
    
    # 創建安裝器實例
    installer = DependencyInstaller()
    
    # 檢查Python版本
    if not installer.check_python_version():
        print("\n❌ 安裝失敗：Python版本不滿足要求")
        return
    
    # 顯示當前安裝狀態
    installer.display_installation_status()
    
    # 詢問是否使用鏡像源
    use_mirror = input("\n是否使用國內鏡像源安裝？(y/n): ").lower().strip() == 'y'
    mirror_url = None
    
    if use_mirror:
        print("\n可用的鏡像源:")
        print("1. 清華大學鏡像: https://pypi.tuna.tsinghua.edu.cn/simple/")
        print("2. 阿里雲鏡像: https://mirrors.aliyun.com/pypi/simple/")
        print("3. 豆瓣鏡像: https://pypi.douban.com/simple/")
        print("4. 自定義鏡像")
        
        choice = input("請選擇鏡像源 (1-4): ").strip()
        
        if choice == "1":
            mirror_url = "https://pypi.tuna.tsinghua.edu.cn/simple/"
        elif choice == "2":
            mirror_url = "https://mirrors.aliyun.com/pypi/simple/"
        elif choice == "3":
            mirror_url = "https://pypi.douban.com/simple/"
        elif choice == "4":
            mirror_url = input("請輸入鏡像源URL: ").strip()
        else:
            print("使用默認源安裝")
            use_mirror = False
    
    # 升級pip
    print("\n🔄 升級pip...")
    if not installer.upgrade_pip():
        print("⚠️ pip升級失敗，但繼續安裝依賴包")
    
    # 安裝核心依賴
    if not installer.install_core_dependencies(use_mirror, mirror_url):
        print("\n❌ 核心依賴安裝失敗")
        return
    
    # 詢問是否安裝可選依賴
    install_optional = input("\n是否安裝可選依賴包？(y/n): ").lower().strip() == 'y'
    
    if install_optional:
        if not installer.install_optional_dependencies(use_mirror, mirror_url):
            print("\n⚠️ 可選依賴安裝失敗，但核心功能仍可正常使用")
    
    # 驗證安裝
    if not installer.verify_installation():
        print("\n❌ 安裝驗證失敗")
        return
    
    # 創建requirements.txt文件
    installer.create_requirements_txt()
    
    # 顯示最終狀態
    print("\n📊 最終安裝狀態:")
    installer.display_installation_status()
    
    print("\n🎉 安裝完成！")
    print("\n📋 使用說明:")
    print("1. 運行系統: streamlit run app.py")
    print("2. 或使用運行腳本: run.bat (Windows) 或 run.sh (Linux/macOS)")
    print("3. 在瀏覽器中打開顯示的URL")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n⚠️ 用戶中斷安裝過程")
    except Exception as e:
        print(f"\n❌ 安裝過程中發生錯誤: {e}")
"""
系統測試腳本
庫存調貨建議系統(澳門優先版)
用於測試庫存調貨建議系統(澳門優先版)v1.0的各個模塊功能
"""

import os
import sys
import traceback
from datetime import datetime
import pandas as pd

# 導入系統模塊
try:
    from data_processor import DataProcessor
    from business_logic import BusinessLogic
    from excel_generator import ExcelGenerator
    IMPORTS_AVAILABLE = True
except ImportError as e:
    print(f"❌ 導入模塊失敗: {e}")
    IMPORTS_AVAILABLE = False

class SystemTester:
    """系統測試器"""
    
    def __init__(self):
        self.test_results = []
        self.data_processor = None
        self.business_logic = None
        self.excel_generator = None
        
        if IMPORTS_AVAILABLE:
            self.data_processor = DataProcessor()
            self.business_logic = BusinessLogic()
            self.excel_generator = ExcelGenerator()
    
    def log_test_result(self, test_name: str, success: bool, message: str = ""):
        """
        記錄測試結果
        
        Args:
            test_name: 測試名稱
            success: 是否成功
            message: 消息
        """
        status = "PASS" if success else "FAIL"
        self.test_results.append({
            'test_name': test_name,
            'success': success,
            'message': message,
            'status': status
        })
        print(f"{status} | {test_name}: {message}")
    
    def test_data_processor(self):
        """測試數據處理模塊"""
        print("\nTesting data processor module...")
        
        if not self.data_processor:
            self.log_test_result("數據處理模塊初始化", False, "模塊導入失敗")
            return
        
        # 測試數據驗證器
        try:
            validator = self.data_processor.validator
            # 創建一個簡單的測試DataFrame
            test_df = pd.DataFrame({
                'Article': ['123456789012', '234567890123'],
                'Article Description': ['Test Product 1', 'Test Product 2'],
                'OM': ['OM001', 'OM002'],
                'RP Type': ['ND', 'RF'],
                'Site': ['SITE001', 'SITE002'],
                'MOQ': [10, 20],
                'SaSa Net Stock': [100, 200],
                'Pending Received': [5, 10],
                'Safety Stock': [50, 80],
                'Last Month Sold Qty': [30, 40],
                'MTD Sold Qty': [15, 20]
            })
            is_valid, missing_cols = validator.validate_columns(test_df)
            self.log_test_result(
                "數據驗證器",
                is_valid,
                f"缺少欄位: {missing_cols}" if not is_valid else "驗證通過"
            )
        except Exception as e:
            self.log_test_result("數據驗證器", False, f"錯誤: {str(e)}")
        
        # 測試數據清洗器
        try:
            cleaner = self.data_processor.cleaner
            # 創建一個包含數值問題的測試DataFrame
            test_df = pd.DataFrame({
                'MOQ': ['10', '20', 'invalid'],
                'SaSa Net Stock': ['100', '200', '300'],
                'Pending Received': ['5', '10', '15'],
                'Safety Stock': ['50', '80', '120']
            })
            numeric_columns = ['MOQ', 'SaSa Net Stock', 'Pending Received', 'Safety Stock']
            cleaned_data = cleaner.clean_numeric_columns(test_df, numeric_columns)
            success = len(cleaned_data) > 0
            self.log_test_result(
                "數據清洗器",
                success,
                f"清洗後 {len(cleaned_data)} 條記錄" if success else "清洗失敗"
            )
        except Exception as e:
            self.log_test_result("數據清洗器", False, f"錯誤: {str(e)}")
    
    def test_business_logic(self):
        """測試業務邏輯模塊"""
        print("\nTesting business logic module...")
        
        if not self.business_logic:
            self.log_test_result("業務邏輯模塊初始化", False, "模塊導入失敗")
            return
        
        # 創建測試數據
        try:
            test_data = pd.DataFrame({
                'Article': ['123456789012', '123456789012', '234567890123', '234567890123', '345678901234', '345678901234'],
                'Article Description': ['Test Product 1', 'Test Product 1', 'Test Product 2', 'Test Product 2', 'Test Product 3', 'Test Product 3'],
                'OM': ['OM001', 'OM002', 'OM001', 'OM002', 'OM001', 'OM002'],
                'RP Type': ['ND', 'RF', 'ND', 'RF', 'RF', 'RF'],
                'Site': ['SITE001', 'SITE002', 'SITE003', 'SITE004', 'SITE005', 'SITE006'],
                'MOQ': [10, 20, 10, 20, 15, 15],
                'SaSa Net Stock': [100, 200, 150, 50, 80, 30],
                'Pending Received': [5, 10, 5, 5, 10, 5],
                'Safety Stock': [50, 80, 60, 40, 50, 30],
                'Last Month Sold Qty': [30, 40, 35, 20, 25, 15],
                'MTD Sold Qty': [15, 20, 18, 10, 12, 8]
            })
            # 計算有效銷量
            test_data['Effective Sold Qty'] = test_data['Last Month Sold Qty'] + test_data['MTD Sold Qty']
        except Exception as e:
            self.log_test_result("業務邏輯測試", False, f"創建測試數據失敗: {str(e)}")
            return
        
        # 測試A模式
        try:
            success, recommendations_a, stats_a = self.business_logic.generate_transfer_recommendations(
                test_data, "A"
            )
            self.log_test_result(
                "A模式調貨建議", 
                success, 
                f"生成 {len(recommendations_a) if recommendations_a else 0} 條建議" if success else "生成失敗"
            )
        except Exception as e:
            self.log_test_result("A mode transfer recommendations", False, f"Error: {str(e)}")
        
        # 測試B模式
        try:
            success, recommendations_b, stats_b = self.business_logic.generate_transfer_recommendations(
                test_data, "B"
            )
            self.log_test_result(
                "B模式調貨建議", 
                success, 
                f"生成 {len(recommendations_b) if recommendations_b else 0} 條建議" if success else "生成失敗"
            )
        except Exception as e:
            self.log_test_result("B mode transfer recommendations", False, f"Error: {str(e)}")
        
        
        # 測試質量檢查
        try:
            if 'recommendations_a' in locals() and recommendations_a:
                is_valid, errors = self.business_logic.quality_checker.check_recommendations(recommendations_a)
                self.log_test_result(
                    "質量檢查",
                    is_valid,
                    f"發現 {len(errors)} 個錯誤" if not is_valid else "檢查通過"
                )
        except Exception as e:
            self.log_test_result("質量檢查", False, f"錯誤: {str(e)}")
    
    def test_excel_generator(self):
        """測試Excel生成模塊"""
        print("\nTesting Excel generator module...")
        
        if not self.excel_generator:
            self.log_test_result("Excel生成模塊初始化", False, "模塊導入失敗")
            return
        
        # 創建測試數據
        try:
            test_data = pd.DataFrame({
                'Article': ['123456789012', '123456789012', '234567890123', '234567890123'],
                'Article Description': ['Test Product 1', 'Test Product 1', 'Test Product 2', 'Test Product 2'],
                'OM': ['OM001', 'OM002', 'OM001', 'OM002'],
                'RP Type': ['ND', 'RF', 'ND', 'RF'],
                'Site': ['SITE001', 'SITE002', 'SITE003', 'SITE004'],
                'MOQ': [10, 20, 10, 20],
                'SaSa Net Stock': [100, 200, 150, 50],
                'Pending Received': [5, 10, 5, 5],
                'Safety Stock': [50, 80, 60, 40],
                'Last Month Sold Qty': [30, 40, 35, 20],
                'MTD Sold Qty': [15, 20, 18, 10]
            })
            # 計算有效銷量
            test_data['Effective Sold Qty'] = test_data['Last Month Sold Qty'] + test_data['MTD Sold Qty']
            
            success, recommendations, stats = self.business_logic.generate_transfer_recommendations(
                test_data, "A"
            )
            
            if not success or not recommendations:
                self.log_test_result("Excel生成測試", False, "無法生成調貨建議")
                return
        except Exception as e:
            self.log_test_result("Excel生成測試", False, f"創建測試數據失敗: {str(e)}")
            return
        
        # 測試Excel文件生成
        try:
            test_filename = "test_output.xlsx"
            success, message, file_path = self.excel_generator.generate_excel_file(
                recommendations, stats, filename=test_filename
            )
            
            if success and os.path.exists(file_path):
                self.log_test_result(
                    "Excel文件生成",
                    True,
                    f"文件已生成: {file_path}"
                )
                
                # 清理測試文件
                try:
                    os.remove(file_path)
                except:
                    pass
            else:
                self.log_test_result("Excel文件生成", False, message)
        except Exception as e:
            self.log_test_result("Excel文件生成", False, f"錯誤: {str(e)}")
        
        # 測試調貨建議格式化
        try:
            formatter = self.excel_generator.transfer_formatter
            df = formatter.format_recommendations_to_dataframe(recommendations)
            success = len(df) > 0
            self.log_test_result(
                "調貨建議格式化",
                success,
                f"格式化 {len(df)} 條記錄" if success else "格式化失敗"
            )
        except Exception as e:
            self.log_test_result("調貨建議格式化", False, f"錯誤: {str(e)}")
        
        # 測試統計摘要格式化
        try:
            formatter = self.excel_generator.summary_formatter
            df = formatter.create_summary_dataframe(stats)
            success = len(df) > 0
            self.log_test_result(
                "統計摘要格式化",
                success,
                f"格式化 {len(df)} 條記錄" if success else "格式化失敗"
            )
        except Exception as e:
            self.log_test_result("統計摘要格式化", False, f"錯誤: {str(e)}")
    
    def test_integration(self):
        """測試集成功能"""
        print("\nTesting integration functionality...")
        
        try:
            # 創建測試數據
            test_data = pd.DataFrame({
                'Article': ['123456789012', '123456789012', '234567890123', '234567890123', '345678901234', '345678901234', '456789012345', '456789012345'],
                'Article Description': ['Test Product 1', 'Test Product 1', 'Test Product 2', 'Test Product 2', 'Test Product 3', 'Test Product 3', 'Test Product 4', 'Test Product 4'],
                'OM': ['OM001', 'OM002', 'OM001', 'OM002', 'OM001', 'OM002', 'OM001', 'OM002'],
                'RP Type': ['ND', 'RF', 'ND', 'RF', 'RF', 'RF', 'ND', 'RF'],
                'Site': ['SITE001', 'SITE002', 'SITE003', 'SITE004', 'SITE005', 'SITE006', 'SITE007', 'SITE008'],
                'MOQ': [10, 20, 10, 20, 15, 15, 10, 20],
                'SaSa Net Stock': [100, 200, 150, 50, 80, 30, 120, 40],
                'Pending Received': [5, 10, 5, 5, 10, 5, 8, 12],
                'Safety Stock': [50, 80, 60, 40, 50, 30, 70, 45],
                'Last Month Sold Qty': [30, 40, 35, 20, 25, 15, 45, 18],
                'MTD Sold Qty': [15, 20, 18, 10, 12, 8, 22, 9]
            })
            # 計算有效銷量
            test_data['Effective Sold Qty'] = test_data['Last Month Sold Qty'] + test_data['MTD Sold Qty']
            
            # 處理數據
            processed_data = self.data_processor.transformer.calculate_effective_sold_qty(test_data)
            
            # 生成A模式調貨建議
            success_a, recommendations_a, stats_a = self.business_logic.generate_transfer_recommendations(
                processed_data, "A"
            )
            
            # 生成B模式調貨建議
            success_b, recommendations_b, stats_b = self.business_logic.generate_transfer_recommendations(
                processed_data, "B"
            )
            
            # 生成Excel文件
            if success_a and recommendations_a:
                excel_success_a, _, _ = self.excel_generator.generate_excel_file(
                    recommendations_a, stats_a, filename="test_integration_a.xlsx"
                )
            else:
                excel_success_a = False
            
            if success_b and recommendations_b:
                excel_success_b, _, _ = self.excel_generator.generate_excel_file(
                    recommendations_b, stats_b, filename="test_integration_b.xlsx"
                )
            else:
                excel_success_b = False
            
            integration_success = (
                len(test_data) > 0 and
                len(processed_data) > 0 and
                success_a and success_b and
                excel_success_a and excel_success_b
            )
            
            self.log_test_result(
                "集成測試",
                integration_success,
                "完整流程測試通過" if integration_success else "集成測試失敗"
            )
            
            # 清理測試文件
            for filename in ["test_integration_a.xlsx", "test_integration_b.xlsx"]:
                if os.path.exists(filename):
                    try:
                        os.remove(filename)
                    except:
                        pass
                        
        except Exception as e:
            self.log_test_result("集成測試", False, f"錯誤: {str(e)}")
    
    def test_edge_cases(self):
        """測試邊界情況"""
        print("\nTesting edge cases...")
        
        # 測試空數據
        try:
            empty_df = pd.DataFrame()
            success, recommendations, stats = self.business_logic.generate_transfer_recommendations(
                empty_df, "A"
            )
            # 空數據應該返回成功但無建議
            expected_success = True
            expected_empty = len(recommendations) == 0
            edge_case_success = success == expected_success and expected_empty
            
            self.log_test_result(
                "空數據處理", 
                edge_case_success, 
                "正確處理空數據" if edge_case_success else "空數據處理異常"
            )
        except Exception as e:
            self.log_test_result("Empty data processing", False, f"Error: {str(e)}")
        
        # 測試單一商品
        try:
            single_article_data = pd.DataFrame({
                'Article': ['123456789012', '123456789012'],
                'Article Description': ['Test Product 1', 'Test Product 1'],
                'OM': ['OM001', 'OM002'],
                'RP Type': ['ND', 'RF'],
                'Site': ['SITE001', 'SITE002'],
                'MOQ': [10, 20],
                'SaSa Net Stock': [100, 200],
                'Pending Received': [5, 10],
                'Safety Stock': [50, 80],
                'Last Month Sold Qty': [30, 40],
                'MTD Sold Qty': [15, 20]
            })
            # 計算有效銷量
            single_article_data['Effective Sold Qty'] = single_article_data['Last Month Sold Qty'] + single_article_data['MTD Sold Qty']
            
            success, recommendations, stats = self.business_logic.generate_transfer_recommendations(
                single_article_data, "A"
            )
            
            self.log_test_result(
                "單一商品處理",
                success,
                f"處理 {len(single_article_data)} 條記錄" if success else "處理失敗"
            )
        except Exception as e:
            self.log_test_result("單一商品處理", False, f"錯誤: {str(e)}")
        
        # 測試無效模式
        try:
            test_data = pd.DataFrame({
                'Article': ['123456789012', '123456789012', '234567890123', '234567890123'],
                'Article Description': ['Test Product 1', 'Test Product 1', 'Test Product 2', 'Test Product 2'],
                'OM': ['OM001', 'OM002', 'OM001', 'OM002'],
                'RP Type': ['ND', 'RF', 'ND', 'RF'],
                'Site': ['SITE001', 'SITE002', 'SITE003', 'SITE004'],
                'MOQ': [10, 20, 10, 20],
                'SaSa Net Stock': [100, 200, 150, 50],
                'Pending Received': [5, 10, 5, 5],
                'Safety Stock': [50, 80, 60, 40],
                'Last Month Sold Qty': [30, 40, 35, 20],
                'MTD Sold Qty': [15, 20, 18, 10]
            })
            # 計算有效銷量
            test_data['Effective Sold Qty'] = test_data['Last Month Sold Qty'] + test_data['MTD Sold Qty']
            
            success, recommendations, stats = self.business_logic.generate_transfer_recommendations(
                test_data, "INVALID"
            )
            
            # 無效模式應該返回失敗
            self.log_test_result(
                "無效模式處理",
                not success,
                "正確拒絕無效模式" if not success else "無效模式處理異常"
            )
        except Exception as e:
            self.log_test_result("無效模式處理", False, f"錯誤: {str(e)}")
    
    def generate_test_report(self):
        """生成測試報告"""
        print("\nTest report")
        print("=" * 60)
        
        total_tests = len(self.test_results)
        passed_tests = sum(1 for result in self.test_results if result['success'])
        failed_tests = total_tests - passed_tests
        success_rate = (passed_tests / total_tests * 100) if total_tests > 0 else 0
        
        print(f"Total tests: {total_tests}")
        print(f"Passed tests: {passed_tests}")
        print(f"Failed tests: {failed_tests}")
        print(f"Success rate: {success_rate:.1f}%")
        print("=" * 60)
        
        print("\nDetailed results:")
        for result in self.test_results:
            print(f"{result['status']} | {result['test_name']}: {result['message']}")
        
        # 保存測試報告到文件
        try:
            report_data = {
                'test_time': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                'total_tests': total_tests,
                'passed_tests': passed_tests,
                'failed_tests': failed_tests,
                'success_rate': success_rate,
                'test_results': self.test_results
            }
            
            report_df = pd.DataFrame(self.test_results)
            report_filename = f"test_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            
            with pd.ExcelWriter(report_filename, engine='openpyxl') as writer:
                report_df.to_excel(writer, sheet_name='Test Results', index=False)
                
                # 添加摘要信息
                summary_data = {
                    'Metric': ['Test Time', 'Total Tests', 'Passed Tests', 'Failed Tests', 'Success Rate (%)'],
                    'Value': [report_data['test_time'], total_tests, passed_tests, failed_tests, success_rate]
                }
                summary_df = pd.DataFrame(summary_data)
                summary_df.to_excel(writer, sheet_name='Summary', index=False)
            
            print(f"\nTest report saved: {report_filename}")
            
        except Exception as e:
            print(f"\nWarning: Failed to save test report: {str(e)}")
    
    def run_all_tests(self):
        """運行所有測試"""
        print("Starting system tests")
        print("=" * 60)
        print(f"Test time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print("=" * 60)
        
        if not IMPORTS_AVAILABLE:
            print("Cannot run tests: Missing required modules")
            return
        
        # 運行各項測試
        self.test_data_processor()
        self.test_business_logic()
        self.test_excel_generator()
        self.test_integration()
        self.test_edge_cases()
        
        # 生成測試報告
        self.generate_test_report()


def main():
    """主函數"""
    try:
        # 創建測試器並運行測試
        tester = SystemTester()
        tester.run_all_tests()
        
    except KeyboardInterrupt:
        print("\n\nUser interrupted the testing process")
    except Exception as e:
        print(f"\nError during testing: {str(e)}")
        traceback.print_exc()


if __name__ == "__main__":
    main()
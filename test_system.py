"""
系統測試腳本
用於測試庫存調貨建議系統v1.0的各個模塊功能
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
        
        # 測試模擬數據生成
        try:
            mock_data = self.data_processor.generate_mock_data(num_articles=5, seed=42)
            success = len(mock_data) > 0
            self.log_test_result(
                "模擬數據生成", 
                success, 
                f"生成 {len(mock_data)} 條記錄" if success else "生成失敗"
            )
        except Exception as e:
            self.log_test_result("Mock data generation", False, f"Error: {str(e)}")
        
        # 測試數據驗證
        try:
            if 'mock_data' in locals():
                validator = self.data_processor.validator
                is_valid, missing_cols = validator.validate_columns(mock_data)
                self.log_test_result(
                    "數據驗證", 
                    is_valid, 
                    f"缺少欄位: {missing_cols}" if not is_valid else "驗證通過"
                )
        except Exception as e:
            self.log_test_result("Data validation", False, f"Error: {str(e)}")
        
        # 測試數據清洗
        try:
            if 'mock_data' in locals():
                cleaner = self.data_processor.cleaner
                numeric_columns = ['MOQ', 'SaSa Net Stock', 'Pending Received', 'Safety Stock']
                cleaned_data = cleaner.clean_numeric_columns(mock_data, numeric_columns)
                success = len(cleaned_data) == len(mock_data)
                self.log_test_result(
                    "數據清洗", 
                    success, 
                    f"清洗後 {len(cleaned_data)} 條記錄" if success else "清洗失敗"
                )
        except Exception as e:
            self.log_test_result("Data cleaning", False, f"Error: {str(e)}")
    
    def test_business_logic(self):
        """測試業務邏輯模塊"""
        print("\nTesting business logic module...")
        
        if not self.business_logic:
            self.log_test_result("業務邏輯模塊初始化", False, "模塊導入失敗")
            return
        
        # 生成測試數據
        try:
            test_data = self.data_processor.generate_mock_data(num_articles=5, seed=42)
        except:
            self.log_test_result("Business logic test", False, "Cannot generate test data")
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
        
        # 測試C模式
        try:
            success, recommendations_c, stats_c = self.business_logic.generate_transfer_recommendations(
                test_data, "C"
            )
            self.log_test_result(
                "C模式調貨建議",
                success,
                f"生成 {len(recommendations_c) if recommendations_c else 0} 條建議" if success else "生成失敗"
            )
        except Exception as e:
            self.log_test_result("C mode transfer recommendations", False, f"Error: {str(e)}")
        
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
            self.log_test_result("Quality check", False, f"Error: {str(e)}")
    
    def test_excel_generator(self):
        """測試Excel生成模塊"""
        print("\nTesting Excel generator module...")
        
        if not self.excel_generator:
            self.log_test_result("Excel生成模塊初始化", False, "模塊導入失敗")
            return
        
        # 生成測試數據
        try:
            test_data = self.data_processor.generate_mock_data(num_articles=3, seed=42)
            success, recommendations, stats = self.business_logic.generate_transfer_recommendations(
                test_data, "A"
            )
            
            if not success or not recommendations:
                self.log_test_result("Excel generation test", False, "Cannot generate transfer recommendations")
                return
        except:
            self.log_test_result("Excel generation test", False, "Cannot generate test data")
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
            self.log_test_result("Excel file generation", False, f"Error: {str(e)}")
        
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
            self.log_test_result("Transfer recommendations formatting", False, f"Error: {str(e)}")
        
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
            self.log_test_result("Summary statistics formatting", False, f"Error: {str(e)}")
    
    def test_integration(self):
        """測試集成功能"""
        print("\nTesting integration functionality...")
        
        try:
            # 生成測試數據
            test_data = self.data_processor.generate_mock_data(num_articles=10, seed=123)
            
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
            
            # 生成C模式調貨建議
            success_c, recommendations_c, stats_c = self.business_logic.generate_transfer_recommendations(
                processed_data, "C"
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
                
            if success_c and recommendations_c:
                excel_success_c, _, _ = self.excel_generator.generate_excel_file(
                    recommendations_c, stats_c, filename="test_integration_c.xlsx"
                )
            else:
                excel_success_c = False
            
            integration_success = (
                len(test_data) > 0 and
                len(processed_data) > 0 and
                success_a and success_b and success_c and
                excel_success_a and excel_success_b and excel_success_c
            )
            
            self.log_test_result(
                "集成測試", 
                integration_success, 
                "完整流程測試通過" if integration_success else "集成測試失敗"
            )
            
            # 清理測試文件
            for filename in ["test_integration_a.xlsx", "test_integration_b.xlsx", "test_integration_c.xlsx"]:
                if os.path.exists(filename):
                    try:
                        os.remove(filename)
                    except:
                        pass
                        
        except Exception as e:
            self.log_test_result("Integration test", False, f"Error: {str(e)}")
    
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
            single_article_data = self.data_processor.generate_mock_data(num_articles=1, seed=456)
            success, recommendations, stats = self.business_logic.generate_transfer_recommendations(
                single_article_data, "A"
            )
            
            self.log_test_result(
                "單一商品處理", 
                success, 
                f"處理 {len(single_article_data)} 條記錄" if success else "處理失敗"
            )
        except Exception as e:
            self.log_test_result("Single article processing", False, f"Error: {str(e)}")
        
        # 測試無效模式
        try:
            test_data = self.data_processor.generate_mock_data(num_articles=3, seed=789)
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
            self.log_test_result("Invalid mode processing", False, f"Error: {str(e)}")
    
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
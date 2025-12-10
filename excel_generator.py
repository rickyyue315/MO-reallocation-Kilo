"""
Excel輸出模塊 v1.8
庫存調貨建議系統(澳門優先版)
負責生成包含調貨建議和統計摘要的Excel文件
"""

import pandas as pd
import numpy as np
from typing import Dict, List, Tuple, Optional
import logging
import xlsxwriter
from datetime import datetime
import os

# 設置日誌
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class TransferFormatter:
    """調貨建議格式化器"""
    
    @staticmethod
    def format_recommendations_to_dataframe(recommendations: List[Dict]) -> pd.DataFrame:
        """
        將調貨建議列表轉換為數據框
        
        Args:
            recommendations: 調貨建議列表
            
        Returns:
            格式化的數據框
        """
        if not recommendations:
            return pd.DataFrame()
        
        # 定義欄位順序
        column_order = [
            'Article', 'Product Desc', 'Transfer OM', 'Transfer Site', 'Receive OM', 'Receive Site',
            'Transfer Qty', 'Transfer Site Original Stock', 'Transfer Site After Transfer Stock',
            'Transfer Site Safety Stock', 'Transfer Site MOQ', 'Remark',
            'Transfer Site Last Month Sold Qty', 'Transfer Site MTD Sold Qty',
            'Receive Site Last Month Sold Qty', 'Receive Site MTD Sold Qty',
            'Receive Original Stock', 'Notes'
        ]
        
        # 創建數據框
        df = pd.DataFrame(recommendations)
        
        # 確保所有欄位都存在
        for col in column_order:
            if col not in df.columns:
                df[col] = ""
        
        # 重新排序列
        df = df[column_order]
        
        return df
    
    @staticmethod
    def get_column_widths() -> Dict[str, int]:
        """
        獲取Excel欄位寬度設置
        
        Returns:
            欄位寬度字典
        """
        return {
            'Article': 12,
            'Product Desc': 25,
            'Transfer OM': 12,
            'Transfer Site': 12,
            'Receive OM': 12,
            'Receive Site': 12,
            'Transfer Qty': 10,
            'Transfer Site Original Stock': 18,
            'Transfer Site After Transfer Stock': 20,
            'Transfer Site Safety Stock': 18,
            'Transfer Site MOQ': 12,
            'Remark': 25,
            'Transfer Site Last Month Sold Qty': 18,
            'Transfer Site MTD Sold Qty': 15,
            'Receive Site Last Month Sold Qty': 18,
            'Receive Site MTD Sold Qty': 15,
            'Receive Original Stock': 15,
            'Notes': 200
        }

class SummaryFormatter:
    """統計摘要格式化器"""
    
    @staticmethod
    def create_summary_dataframe(stats: Dict) -> pd.DataFrame:
        """
        創建統計摘要數據框
        
        Args:
            stats: 統計信息字典
            
        Returns:
            統計摘要數據框
        """
        summary_data = []
        
        # 基本統計
        summary_data.extend([
            ['基本統計', '', ''],
            ['總建議數量', stats.get('total_recommendations', 0), '條'],
            ['總轉移數量', stats.get('total_transfer_quantity', 0), '件'],
            ['涉及商品數量', stats.get('unique_articles', 0), '個'],
            ['轉出店鋪數量', stats.get('unique_transfer_sites', 0), '個'],
            ['接收店鋪數量', stats.get('unique_receive_sites', 0), '個'],
            ['', '', ''],
            ['轉出類型統計', '', '']
        ])
        
        # 轉出類型統計
        transfer_type_stats = stats.get('transfer_type_stats', {})
        for transfer_type, type_stats in transfer_type_stats.items():
            summary_data.append([
                f'  {transfer_type}',
                type_stats.get('count', 0),
                f'條 ({type_stats.get("quantity", 0)}件)'
            ])
        
        summary_data.extend([
            ['', '', ''],
            ['接收優先級統計', '', '']
        ])
        
        # 接收優先級統計
        receive_priority_stats = stats.get('receive_priority_stats', {})
        for priority, priority_stats in receive_priority_stats.items():
            summary_data.append([
                f'  {priority}',
                priority_stats.get('count', 0),
                f'條 ({priority_stats.get("quantity", 0)}件)'
            ])
        
        # 創建數據框
        df = pd.DataFrame(summary_data, columns=['項目', '數值', '單位'])
        
        return df
    
    @staticmethod
    def create_transfer_type_chart_data(stats: Dict) -> pd.DataFrame:
        """
        創建轉出類型圖表數據
        
        Args:
            stats: 統計信息字典
            
        Returns:
            轉出類型圖表數據框
        """
        transfer_type_stats = stats.get('transfer_type_stats', {})
        
        chart_data = []
        for transfer_type, type_stats in transfer_type_stats.items():
            chart_data.append({
                '轉出類型': transfer_type,
                '建議數量': type_stats.get('count', 0),
                '轉移數量': type_stats.get('quantity', 0)
            })
        
        return pd.DataFrame(chart_data)
    
    @staticmethod
    def create_receive_priority_chart_data(stats: Dict) -> pd.DataFrame:
        """
        創建接收優先級圖表數據
        
        Args:
            stats: 統計信息字典
            
        Returns:
            接收優先級圖表數據框
        """
        receive_priority_stats = stats.get('receive_priority_stats', {})
        
        chart_data = []
        for priority, priority_stats in receive_priority_stats.items():
            chart_data.append({
                '接收優先級': priority,
                '建議數量': priority_stats.get('count', 0),
                '轉移數量': priority_stats.get('quantity', 0)
            })
        
        return pd.DataFrame(chart_data)

class ExcelWriter:
    """Excel文件寫入器"""
    
    def __init__(self):
        self.transfer_formatter = TransferFormatter()
        self.summary_formatter = SummaryFormatter()
    
    def write_excel_file(self, recommendations: List[Dict], stats: Dict, output_path: str) -> Tuple[bool, str]:
        """
        寫入Excel文件
        
        Args:
            recommendations: 調貨建議列表
            stats: 統計信息字典
            output_path: 輸出文件路徑
            
        Returns:
            (是否成功, 消息)
        """
        try:
            # 創建Excel工作簿
            with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
                # 寫入調貨建議工作表
                self._write_transfer_recommendations_sheet(writer, recommendations)
                
                # 寫入統計摘要工作表
                self._write_summary_dashboard_sheet(writer, stats)
                
                # 寫入圖表數據工作表
                self._write_chart_data_sheets(writer, stats)
            
            return True, f"Excel文件已成功生成: {output_path}"
            
        except Exception as e:
            logger.error(f"寫入Excel文件時發生錯誤: {str(e)}")
            return False, f"寫入Excel文件時發生錯誤: {str(e)}"
    
    def _write_transfer_recommendations_sheet(self, writer: pd.ExcelWriter, recommendations: List[Dict]):
        """
        寫入調貨建議工作表
        
        Args:
            writer: Excel寫入器
            recommendations: 調貨建議列表
        """
        # 格式化調貨建議
        df = self.transfer_formatter.format_recommendations_to_dataframe(recommendations)
        
        if df.empty:
            # 如果沒有數據，創建空表
            df = pd.DataFrame(columns=[
                'Article', 'Product Desc', 'Transfer OM', 'Transfer Site', 'Receive OM', 'Receive Site',
                'Transfer Qty', 'Transfer Site Original Stock', 'Transfer Site After Transfer Stock',
                'Transfer Site Safety Stock', 'Transfer Site MOQ', 'Remark',
                'Transfer Site Last Month Sold Qty', 'Transfer Site MTD Sold Qty',
                'Receive Site Last Month Sold Qty', 'Receive Site MTD Sold Qty',
                'Receive Original Stock', 'Notes'
            ])
        
        # 寫入數據
        df.to_excel(writer, sheet_name='Transfer Recommendations', index=False)
        
        # 獲取工作簿和工作表對象
        workbook = writer.book
        worksheet = writer.sheets['Transfer Recommendations']
        
        # 設置欄位寬度
        column_widths = self.transfer_formatter.get_column_widths()
        for col_num, column in enumerate(df.columns):
            if column in column_widths:
                worksheet.set_column(col_num, col_num, column_widths[column])
        
        # 設置標題格式
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#D7E4BD',
            'border': 1
        })
        
        # 應用標題格式
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
        
        # 設置數據格式
        data_format = workbook.add_format({
            'border': 1,
            'text_wrap': True,
            'valign': 'top'
        })
        
        # 應用數據格式
        for row_num in range(1, len(df) + 1):
            for col_num in range(len(df.columns)):
                worksheet.write(row_num, col_num, df.iloc[row_num-1, col_num], data_format)
    
    def _write_summary_dashboard_sheet(self, writer: pd.ExcelWriter, stats: Dict):
        """
        寫入統計摘要工作表
        
        Args:
            writer: Excel寫入器
            stats: 統計信息字典
        """
        # 創建統計摘要
        summary_df = self.summary_formatter.create_summary_dataframe(stats)
        
        # 寫入數據
        summary_df.to_excel(writer, sheet_name='Summary Dashboard', index=False)
        
        # 獲取工作簿和工作表對象
        workbook = writer.book
        worksheet = writer.sheets['Summary Dashboard']
        
        # 設置欄位寬度
        worksheet.set_column('A:A', 25)
        worksheet.set_column('B:B', 15)
        worksheet.set_column('C:C', 25)
        
        # 設置標題格式
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#D7E4BD',
            'border': 1
        })
        
        # 應用標題格式
        for col_num, value in enumerate(summary_df.columns.values):
            worksheet.write(0, col_num, value, header_format)
        
        # 設置數據格式
        data_format = workbook.add_format({
            'border': 1,
            'text_wrap': True,
            'valign': 'top'
        })
        
        # 應用數據格式
        for row_num in range(1, len(summary_df) + 1):
            for col_num in range(len(summary_df.columns)):
                worksheet.write(row_num, col_num, summary_df.iloc[row_num-1, col_num], data_format)
        
        # 添加生成時間
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        worksheet.write(len(summary_df) + 3, 0, f"生成時間: {current_time}")
    
    def _write_chart_data_sheets(self, writer: pd.ExcelWriter, stats: Dict):
        """
        寫入圖表數據工作表
        
        Args:
            writer: Excel寫入器
            stats: 統計信息字典
        """
        # 轉出類型圖表數據
        transfer_chart_df = self.summary_formatter.create_transfer_type_chart_data(stats)
        transfer_chart_df.to_excel(writer, sheet_name='Chart_Transfer_Type', index=False)
        
        # 接收優先級圖表數據
        receive_chart_df = self.summary_formatter.create_receive_priority_chart_data(stats)
        receive_chart_df.to_excel(writer, sheet_name='Chart_Receive_Priority', index=False)

class ExcelGenerator:
    """Excel生成器主類"""
    
    def __init__(self):
        self.excel_writer = ExcelWriter()
    
    def generate_excel_file(self, recommendations: List[Dict], stats: Dict, 
                           output_dir: str = ".", filename: str = None) -> Tuple[bool, str, str]:
        """
        生成Excel文件
        
        Args:
            recommendations: 調貨建議列表
            stats: 統計信息字典
            output_dir: 輸出目錄
            filename: 文件名（可選）
            
        Returns:
            (是否成功, 消息, 文件路徑)
        """
        try:
            # 生成文件名
            if not filename:
                current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
                filename = f"transfer_recommendations_{current_time}.xlsx"
            
            # 確保文件擴展名正確
            if not filename.endswith('.xlsx'):
                filename += '.xlsx'
            
            # 創建輸出目錄（如果不存在）
            os.makedirs(output_dir, exist_ok=True)
            
            # 構建完整文件路徑
            output_path = os.path.join(output_dir, filename)
            
            # 寫入Excel文件
            success, message = self.excel_writer.write_excel_file(
                recommendations, stats, output_path
            )
            
            if success:
                return True, message, output_path
            else:
                return False, message, ""
                
        except Exception as e:
            logger.error(f"生成Excel文件時發生錯誤: {str(e)}")
            return False, f"生成Excel文件時發生錯誤: {str(e)}", ""


# 測試代碼
if __name__ == "__main__":
    # 測試Excel生成器
    from data_processor import DataProcessor
    from business_logic import BusinessLogic
    
    # 創建測試數據
    processor = DataProcessor()
    test_data = processor.generate_mock_data(num_articles=5, seed=42)
    
    # 創建業務邏輯實例
    business_logic = BusinessLogic()
    
    # 生成調貨建議
    success, recommendations, stats = business_logic.generate_transfer_recommendations(test_data, "A")
    
    if success and recommendations:
        # 創建Excel生成器
        excel_generator = ExcelGenerator()
        
        # 生成Excel文件
        success, message, file_path = excel_generator.generate_excel_file(
            recommendations, stats, filename="test_recommendations.xlsx"
        )
        
        print(f"Excel生成測試結果: {success}")
        print(f"消息: {message}")
        if file_path:
            print(f"文件路徑: {file_path}")
    else:
        print("沒有生成調貨建議，無法測試Excel生成")
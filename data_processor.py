"""
數據預處理模塊 v1.8.1
負責Excel文件讀取、數據驗證、清洗和轉換，以及模擬數據生成
"""

import pandas as pd
import numpy as np
import re
from typing import Dict, List, Tuple, Optional, Union
import logging

# 設置日誌
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class DataValidator:
    """數據驗證器，驗證數據完整性和格式"""
    
    REQUIRED_COLUMNS = [
        'Article', 'Article Description', 'OM', 'RP Type', 'Site',
        'MOQ', 'SaSa Net Stock', 'Pending Received', 'Safety Stock',
        'Last Month Sold Qty', 'MTD Sold Qty'
    ]
    
    ALTERNATIVE_DESCRIPTION_COLUMNS = ['Article Long Text (60 Chars)']
    
    @staticmethod
    def validate_columns(df: pd.DataFrame) -> Tuple[bool, List[str]]:
        """
        驗證必需欄位是否存在
        
        Args:
            df: 輸入數據框
            
        Returns:
            (是否有效, 缺失欄位列表)
        """
        missing_columns = []
        
        # 檢查必需欄位
        for col in DataValidator.REQUIRED_COLUMNS:
            if col not in df.columns:
                # 檢查商品描述的替代欄位
                if col == 'Article Description':
                    has_alternative = any(alt_col in df.columns 
                                        for alt_col in DataValidator.ALTERNATIVE_DESCRIPTION_COLUMNS)
                    if not has_alternative:
                        missing_columns.append(col)
                else:
                    missing_columns.append(col)
        
        is_valid = len(missing_columns) == 0
        return is_valid, missing_columns
    
    @staticmethod
    def validate_article_format(articles: pd.Series) -> Tuple[bool, List[str]]:
        """
        驗證Article格式（12位文本）
        
        Args:
            articles: Article系列
            
        Returns:
            (是否有效, 無效Article列表)
        """
        invalid_articles = []
        
        for idx, article in enumerate(articles):
            if pd.isna(article) or not isinstance(article, str):
                invalid_articles.append(f"行{idx+1}: {article}")
            elif len(article.strip()) != 12 or not article.strip().isdigit():
                invalid_articles.append(f"行{idx+1}: {article}")
        
        is_valid = len(invalid_articles) == 0
        return is_valid, invalid_articles
    
    @staticmethod
    def validate_rp_types(rp_types: pd.Series) -> Tuple[bool, List[str]]:
        """
        驗證RP Type欄位（只能是ND或RF）
        
        Args:
            rp_types: RP Type系列
            
        Returns:
            (是否有效, 無效RP Type列表)
        """
        valid_types = {'ND', 'RF'}
        invalid_types = []
        
        for idx, rp_type in enumerate(rp_types):
            if pd.isna(rp_type) or rp_type.upper() not in valid_types:
                invalid_types.append(f"行{idx+1}: {rp_type}")
        
        is_valid = len(invalid_types) == 0
        return is_valid, invalid_types


class DataCleaner:
    """數據清洗器，處理缺失值和異常值"""
    
    @staticmethod
    def clean_numeric_columns(df: pd.DataFrame, numeric_columns: List[str]) -> pd.DataFrame:
        """
        清洗數值欄位，處理負值和異常值
        
        Args:
            df: 輸入數據框
            numeric_columns: 數值欄位列表
            
        Returns:
            清洗後的數據框
        """
        df_cleaned = df.copy()
        
        for col in numeric_columns:
            if col in df_cleaned.columns:
                # 轉換為數值類型
                df_cleaned[col] = pd.to_numeric(df_cleaned[col], errors='coerce')
                
                # 處理負值（除了銷量欄位）
                if 'Sold' not in col:
                    df_cleaned[col] = df_cleaned[col].clip(lower=0)
                
                # 填充缺失值為0
                df_cleaned[col] = df_cleaned[col].fillna(0)
        
        return df_cleaned
    
    @staticmethod
    def clean_text_columns(df: pd.DataFrame, text_columns: List[str]) -> pd.DataFrame:
        """
        清洗文本欄位，去除多餘空格
        
        Args:
            df: 輸入數據框
            text_columns: 文本欄位列表
            
        Returns:
            清洗後的數據框
        """
        df_cleaned = df.copy()
        
        for col in text_columns:
            if col in df_cleaned.columns:
                df_cleaned[col] = df_cleaned[col].astype(str).str.strip()
        
        return df_cleaned
    
    @staticmethod
    def remove_duplicates(df: pd.DataFrame) -> pd.DataFrame:
        """
        移除重複行
        
        Args:
            df: 輸入數據框
            
        Returns:
            去重後的數據框
        """
        return df.drop_duplicates()


class DataTransformer:
    """數據轉換器，數據類型轉換和格式化"""
    
    @staticmethod
    def standardize_article_column(df: pd.DataFrame) -> pd.DataFrame:
        """
        標準化Article欄位為12位文本格式
        
        Args:
            df: 輸入數據框
            
        Returns:
            標準化後的數據框
        """
        df_transformed = df.copy()
        
        if 'Article' in df_transformed.columns:
            # 確保Article為字符串類型
            df_transformed['Article'] = df_transformed['Article'].astype(str)
            
            # 填充前導零至12位
            df_transformed['Article'] = df_transformed['Article'].apply(
                lambda x: x.zfill(12)[:12] if x.replace('.', '').isdigit() else x
            )
        
        return df_transformed
    
    @staticmethod
    def standardize_rp_type(df: pd.DataFrame) -> pd.DataFrame:
        """
        標準化RP Type欄位為大寫
        
        Args:
            df: 輸入數據框
            
        Returns:
            標準化後的數據框
        """
        df_transformed = df.copy()
        
        if 'RP Type' in df_transformed.columns:
            df_transformed['RP Type'] = df_transformed['RP Type'].astype(str).str.upper()
        
        return df_transformed
    
    @staticmethod
    def calculate_effective_sold_qty(df: pd.DataFrame) -> pd.DataFrame:
        """
        計算有效銷量
        
        Args:
            df: 輸入數據框
            
        Returns:
            添加有效銷量欄位的數據框
        """
        df_transformed = df.copy()
        
        if 'Last Month Sold Qty' in df_transformed.columns and 'MTD Sold Qty' in df_transformed.columns:
            # 有效銷量 = 上月銷量 + 本月銷量
            df_transformed['Effective Sold Qty'] = (
                df_transformed['Last Month Sold Qty'] + df_transformed['MTD Sold Qty']
            )
        
        return df_transformed


class MockDataGenerator:
    """模擬數據生成器，生成測試用模擬數據"""
    
    @staticmethod
    def generate_mock_articles(num_articles: int = 20) -> List[str]:
        """
        生成模擬商品編號
        
        Args:
            num_articles: 商品數量
            
        Returns:
            商品編號列表
        """
        return [f"{i:012d}" for i in range(100000000000, 100000000000 + num_articles)]
    
    @staticmethod
    def generate_mock_descriptions(num_articles: int = 20) -> List[str]:
        """
        生成模擬商品描述
        
        Args:
            num_articles: 商品數量
            
        Returns:
            商品描述列表
        """
        descriptions = []
        for i in range(num_articles):
            desc = f"商品{i+1:03d}測試描述"
            descriptions.append(desc)
        return descriptions
    
    @staticmethod
    def generate_mock_sites() -> List[Dict]:
        """
        生成模擬店鋪信息
        
        Returns:
            店鋪信息列表
        """
        sites = []
        
        # ND類型店鋪
        for i in range(1, 6):
            sites.append({
                'OM': f'OM{i:02d}',
                'RP Type': 'ND',
                'Site': f'ND{i:03d}'
            })
        
        # RF類型店鋪 - 包括HA, H, HC, HD類型
        rf_sites = [
            {'prefix': 'HA', 'count': 3},
            {'prefix': 'H', 'count': 3},
            {'prefix': 'HC', 'count': 3},
            {'prefix': 'HD', 'count': 3}
        ]
        
        for site_type in rf_sites:
            for i in range(1, site_type['count'] + 1):
                sites.append({
                    'OM': f'OM{len(sites)+1:02d}',
                    'RP Type': 'RF',
                    'Site': f"{site_type['prefix']}{i:03d}"
                })
        
        return sites
    
    @staticmethod
    def generate_mock_inventory_data(num_articles: int = 20, seed: int = 42) -> pd.DataFrame:
        """
        生成完整的模擬庫存數據
        
        Args:
            num_articles: 商品數量
            seed: 隨機種子
            
        Returns:
            模擬庫存數據框
        """
        np.random.seed(seed)
        
        # 生成基本數據
        articles = MockDataGenerator.generate_mock_articles(num_articles)
        descriptions = MockDataGenerator.generate_mock_descriptions(num_articles)
        sites = MockDataGenerator.generate_mock_sites()
        
        # 創建數據行
        data_rows = []
        
        for article_idx, (article, description) in enumerate(zip(articles, descriptions)):
            for site in sites:
                # 生成隨機數據
                moq = np.random.randint(1, 6)
                safety_stock = np.random.randint(5, 20)
                net_stock = np.random.randint(0, 50)
                pending_received = np.random.randint(0, 20)
                last_month_sold = np.random.randint(0, 30)
                mtd_sold = np.random.randint(0, 20)
                
                # 為某些店鋪創建缺貨情況
                if site['RP Type'] == 'RF' and np.random.random() < 0.3:
                    net_stock = 0
                    pending_received = 0
                
                # 為某些店鋪創建過剩庫存
                if site['RP Type'] == 'RF' and np.random.random() < 0.2:
                    net_stock = safety_stock + np.random.randint(10, 30)
                
                data_rows.append({
                    'Article': article,
                    'Article Description': description,
                    'OM': site['OM'],
                    'RP Type': site['RP Type'],
                    'Site': site['Site'],
                    'MOQ': moq,
                    'SaSa Net Stock': net_stock,
                    'Pending Received': pending_received,
                    'Safety Stock': safety_stock,
                    'Last Month Sold Qty': last_month_sold,
                    'MTD Sold Qty': mtd_sold
                })
        
        return pd.DataFrame(data_rows)


class DataProcessor:
    """數據處理主類，協調各個處理步驟"""
    
    def __init__(self):
        self.validator = DataValidator()
        self.cleaner = DataCleaner()
        self.transformer = DataTransformer()
        self.mock_generator = MockDataGenerator()
    
    def process_uploaded_file(self, file_path: str) -> Tuple[bool, Union[pd.DataFrame, str], Dict]:
        """
        處理上傳的Excel文件
        
        Args:
            file_path: 文件路徑
            
        Returns:
            (是否成功, 處理後的數據框或錯誤信息, 處理統計)
        """
        try:
            # 讀取Excel文件
            df = pd.read_excel(file_path)
            
            # 驗證欄位
            is_valid, missing_columns = self.validator.validate_columns(df)
            if not is_valid:
                error_msg = f"缺少必需欄位: {', '.join(missing_columns)}"
                return False, error_msg, {}
            
            # 標準化商品描述欄位名稱
            if 'Article Description' not in df.columns:
                for alt_col in DataValidator.ALTERNATIVE_DESCRIPTION_COLUMNS:
                    if alt_col in df.columns:
                        df = df.rename(columns={alt_col: 'Article Description'})
                        break
            
            # 數據清洗
            numeric_columns = [
                'MOQ', 'SaSa Net Stock', 'Pending Received', 'Safety Stock',
                'Last Month Sold Qty', 'MTD Sold Qty'
            ]
            text_columns = ['Article', 'Article Description', 'OM', 'RP Type', 'Site']
            
            df = self.cleaner.clean_numeric_columns(df, numeric_columns)
            df = self.cleaner.clean_text_columns(df, text_columns)
            df = self.cleaner.remove_duplicates(df)
            
            # 數據轉換
            df = self.transformer.standardize_article_column(df)
            df = self.transformer.standardize_rp_type(df)
            df = self.transformer.calculate_effective_sold_qty(df)
            
            # 驗證處理後的數據
            is_valid, invalid_articles = self.validator.validate_article_format(df['Article'])
            if not is_valid:
                error_msg = f"Article格式錯誤: {', '.join(invalid_articles[:5])}"
                if len(invalid_articles) > 5:
                    error_msg += f" 等{len(invalid_articles)}個錯誤"
                return False, error_msg, {}
            
            is_valid, invalid_rp_types = self.validator.validate_rp_types(df['RP Type'])
            if not is_valid:
                error_msg = f"RP Type錯誤: {', '.join(invalid_rp_types[:5])}"
                if len(invalid_rp_types) > 5:
                    error_msg += f" 等{len(invalid_rp_types)}個錯誤"
                return False, error_msg, {}
            
            # 生成處理統計
            stats = {
                'total_rows': len(df),
                'unique_articles': df['Article'].nunique(),
                'unique_sites': df['Site'].nunique(),
                'nd_sites': len(df[df['RP Type'] == 'ND']['Site'].unique()),
                'rf_sites': len(df[df['RP Type'] == 'RF']['Site'].unique()),
                'total_stock': df['SaSa Net Stock'].sum(),
                'total_safety_stock': df['Safety Stock'].sum()
            }
            
            return True, df, stats
            
        except Exception as e:
            logger.error(f"處理文件時發生錯誤: {str(e)}")
            return False, f"處理文件時發生錯誤: {str(e)}", {}
    
    def generate_mock_data(self, num_articles: int = 20, seed: int = 42) -> pd.DataFrame:
        """
        生成模擬數據
        
        Args:
            num_articles: 商品數量
            seed: 隨機種子
            
        Returns:
            模擬數據框
        """
        df = self.mock_generator.generate_mock_inventory_data(num_articles, seed)
        
        # 應用相同的處理流程
        df = self.transformer.calculate_effective_sold_qty(df)
        
        return df


# 測試代碼
if __name__ == "__main__":
    processor = DataProcessor()
    
    # 測試模擬數據生成
    mock_data = processor.generate_mock_data(num_articles=5)
    print("模擬數據生成測試:")
    print(mock_data.head())
    print(f"數據形狀: {mock_data.shape}")
    print(f"商品數量: {mock_data['Article'].nunique()}")
    print(f"店鋪數量: {mock_data['Site'].nunique()}")
"""
業務邏輯模塊 v1.9
庫存調貨建議系統(澳門優先版)
實現調貨算法和業務規則，包括A模式(保守轉貨)、B模式(加強轉貨)和C模式(全量轉貨)
"""

import pandas as pd
import numpy as np
from typing import Dict, List, Tuple, Optional, Union
import logging
from enum import Enum

# 設置日誌
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class TransferType(Enum):
    """轉貨類型枚舉"""
    ND_TRANSFER = "ND轉出"
    RF_SURPLUS_TRANSFER = "RF過剩轉出"
    RF_ENHANCED_TRANSFER = "RF加強轉出"
    C_COMPLETE_TRANSFER = "C模式全量轉出"

class ReceivePriority(Enum):
    """接收優先級枚舉"""
    URGENT_SHORTAGE = "緊急缺貨"
    POTENTIAL_SHORTAGE = "潛在缺貨"

class TransferModeA:
    """A模式(保守轉貨)邏輯"""
    
    @staticmethod
    def identify_nd_transfer_candidates(df: pd.DataFrame) -> pd.DataFrame:
        """
        識別ND類型轉出候選
        
        Args:
            df: 庫存數據框
            
        Returns:
            ND轉出候選數據框
        """
        nd_candidates = df[df['RP Type'] == 'ND'].copy()
        
        if len(nd_candidates) == 0:
            return pd.DataFrame()
        
        # ND類型可轉出全部淨庫存
        nd_candidates['transfer_type'] = TransferType.ND_TRANSFER.value
        nd_candidates['available_quantity'] = nd_candidates['SaSa Net Stock']
        nd_candidates['original_stock'] = nd_candidates['SaSa Net Stock']
        
        # 只保留有庫存的記錄
        nd_candidates = nd_candidates[nd_candidates['available_quantity'] > 0]
        
        return nd_candidates
    
    @staticmethod
    def identify_rf_surplus_transfer_candidates(df: pd.DataFrame) -> pd.DataFrame:
        """
        識別RF類型過剩轉出候選
        
        Args:
            df: 庫存數據框
            
        Returns:
            RF過剩轉出候選數據框
        """
        rf_candidates = df[df['RP Type'] == 'RF'].copy()
        
        if len(rf_candidates) == 0:
            return pd.DataFrame()
        
        # 計算每個商品在各店鋪的有效銷量排名
        rf_candidates['max_sold_per_article'] = rf_candidates.groupby('Article')['Effective Sold Qty'].transform('max')
        
        # 條件篩選
        condition1 = rf_candidates['SaSa Net Stock'] + rf_candidates['Pending Received'] > rf_candidates['Safety Stock']
        condition2 = rf_candidates['Effective Sold Qty'] < rf_candidates['max_sold_per_article']
        
        rf_candidates = rf_candidates[condition1 & condition2].copy()
        
        if len(rf_candidates) == 0:
            return pd.DataFrame()
        
        # 計算可轉出數量
        total_stock = rf_candidates['SaSa Net Stock'] + rf_candidates['Pending Received']
        safety_stock = rf_candidates['Safety Stock']
        
        # 基礎可轉出 = (庫存+在途) - 安全庫存
        basic_transferable = total_stock - safety_stock
        
        # 上限控制 = (庫存+在途) × 40%，但最少出貨2件
        upper_limit = np.maximum(total_stock * 0.4, 2)
        
        # 實際轉出 = min(基礎可轉出, max(上限控制, 2))
        calculated_quantity = np.floor(np.minimum(basic_transferable, upper_limit)).astype(int)
        
        # 確保可轉出數量不超過SaSa Net Stock
        rf_candidates['available_quantity'] = np.minimum(calculated_quantity, rf_candidates['SaSa Net Stock'])
        
        # 確保轉出後剩餘庫存不低於安全庫存
        remaining_stock = total_stock - rf_candidates['available_quantity']
        valid_transfers = remaining_stock >= safety_stock
        
        rf_candidates = rf_candidates[valid_transfers].copy()
        rf_candidates['transfer_type'] = TransferType.RF_SURPLUS_TRANSFER.value
        rf_candidates['original_stock'] = rf_candidates['SaSa Net Stock']
        
        # 只保留有可轉出數量的記錄
        rf_candidates = rf_candidates[rf_candidates['available_quantity'] > 0]
        
        return rf_candidates

class TransferModeB:
    """B模式(加強轉貨)邏輯"""
    
    @staticmethod
    def identify_nd_transfer_candidates(df: pd.DataFrame) -> pd.DataFrame:
        """
        識別ND類型轉出候選（與A模式相同）
        
        Args:
            df: 庫存數據框
            
        Returns:
            ND轉出候選數據框
        """
        return TransferModeA.identify_nd_transfer_candidates(df)
    
    @staticmethod
    def identify_rf_transfer_candidates(df: pd.DataFrame) -> pd.DataFrame:
        """
        識別RF類型轉出候選（加強模式）
        
        Args:
            df: 庫存數據框
            
        Returns:
            RF轉出候選數據框
        """
        rf_candidates = df[df['RP Type'] == 'RF'].copy()
        
        if len(rf_candidates) == 0:
            return pd.DataFrame()
        
        # 計算每個商品在各店鋪的有效銷量排名
        rf_candidates['max_sold_per_article'] = rf_candidates.groupby('Article')['Effective Sold Qty'].transform('max')
        
        # 條件篩選
        condition1 = rf_candidates['SaSa Net Stock'] + rf_candidates['Pending Received'] > rf_candidates['MOQ']
        condition2 = rf_candidates['Effective Sold Qty'] < rf_candidates['max_sold_per_article']
        
        rf_candidates = rf_candidates[condition1 & condition2].copy()
        
        if len(rf_candidates) == 0:
            return pd.DataFrame()
        
        # 計算可轉出數量
        total_stock = rf_candidates['SaSa Net Stock'] + rf_candidates['Pending Received']
        moq = rf_candidates['MOQ']
        
        # 基礎可轉出 = (庫存+在途) – (MOQ數量)
        basic_transferable = total_stock - moq
        
        # 上限控制 = (庫存+在途) × 80%，但最少出貨2件
        upper_limit = np.maximum(total_stock * 0.8, 2)
        
        # 實際轉出 = min(基礎可轉出, max(上限控制, 2))
        calculated_quantity = np.floor(np.minimum(basic_transferable, upper_limit)).astype(int)
        
        # 確保可轉出數量不超過SaSa Net Stock
        rf_candidates['available_quantity'] = np.minimum(calculated_quantity, rf_candidates['SaSa Net Stock'])
        
        # 判斷轉出類型
        remaining_stock = total_stock - rf_candidates['available_quantity']
        safety_stock = rf_candidates['Safety Stock']
        
        # 如果轉出後剩餘庫存 ≥ 安全庫存：RF過剩轉出
        # 如果轉出後剩餘庫存 < 安全庫存：RF加強轉出
        rf_candidates['transfer_type'] = np.where(
            remaining_stock >= safety_stock,
            TransferType.RF_SURPLUS_TRANSFER.value,
            TransferType.RF_ENHANCED_TRANSFER.value
        )
        
        rf_candidates['original_stock'] = rf_candidates['SaSa Net Stock']
        
        # 只保留有可轉出數量的記錄
        rf_candidates = rf_candidates[rf_candidates['available_quantity'] > 0]
        
        return rf_candidates

class TransferModeC:
    """C模式(全量轉貨)邏輯"""
    
    @staticmethod
    def identify_nd_transfer_candidates(df: pd.DataFrame) -> pd.DataFrame:
        """
        識別ND類型轉出候選（與A、B模式相同）
        
        Args:
            df: 庫存數據框
            
        Returns:
            ND轉出候選數據框
        """
        nd_candidates = df[df['RP Type'] == 'ND'].copy()
        
        if len(nd_candidates) == 0:
            return pd.DataFrame()
        
        # 應用C模式的OM和Article限制
        # 按OM和Article分組，檢查銷售量
        nd_candidates['min_sold_in_om_article'] = nd_candidates.groupby(['OM', 'Article'])['Effective Sold Qty'].transform('min')
        
        # 只保留銷售量為同OM同Article中最小的記錄
        nd_candidates = nd_candidates[
            nd_candidates['Effective Sold Qty'] == nd_candidates['min_sold_in_om_article']
        ].copy()
        
        if len(nd_candidates) == 0:
            return pd.DataFrame()
        
        # ND類型可轉出全部淨庫存
        # 如果銷售量為0，可以轉出全部庫存
        nd_candidates['transfer_type'] = TransferType.ND_TRANSFER.value
        nd_candidates['available_quantity'] = nd_candidates['SaSa Net Stock']
        nd_candidates['original_stock'] = nd_candidates['SaSa Net Stock']
        
        # 只保留有庫存的記錄
        nd_candidates = nd_candidates[nd_candidates['available_quantity'] > 0]
        
        return nd_candidates
    
    @staticmethod
    def identify_rf_transfer_candidates(df: pd.DataFrame) -> pd.DataFrame:
        """
        識別RF類型轉出候選（C模式特殊規則）
        
        Args:
            df: 庫存數據框
            
        Returns:
            RF轉出候選數據框
        """
        rf_candidates = df[df['RP Type'] == 'RF'].copy()
        
        if len(rf_candidates) == 0:
            return pd.DataFrame()
        
        # 應用C模式的OM和Article限制
        # 按OM和Article分組，檢查銷售量
        rf_candidates['min_sold_in_om_article'] = rf_candidates.groupby(['OM', 'Article'])['Effective Sold Qty'].transform('min')
        
        # 只保留銷售量為同OM同Article中最小的記錄
        rf_candidates = rf_candidates[
            rf_candidates['Effective Sold Qty'] == rf_candidates['min_sold_in_om_article']
        ].copy()
        
        if len(rf_candidates) == 0:
            return pd.DataFrame()
        
        # C模式特殊規則：如果銷售量為0，可以轉出全部庫存
        # 如果銷售量>0，則不轉出（保留庫存給有銷售的店鋪）
        
        # 計算可轉出數量
        total_stock = rf_candidates['SaSa Net Stock'] + rf_candidates['Pending Received']
        
        # 對於銷售量為0的店鋪，可以轉出全部庫存
        zero_sales_mask = rf_candidates['Effective Sold Qty'] == 0
        rf_candidates.loc[zero_sales_mask, 'available_quantity'] = rf_candidates.loc[zero_sales_mask, 'SaSa Net Stock']
        rf_candidates.loc[zero_sales_mask, 'transfer_type'] = TransferType.C_COMPLETE_TRANSFER.value
        
        # 對於銷售量>0的店鋪，不轉出
        positive_sales_mask = rf_candidates['Effective Sold Qty'] > 0
        rf_candidates = rf_candidates[~positive_sales_mask].copy()
        
        if len(rf_candidates) == 0:
            return rf_candidates[zero_sales_mask].copy() if len(rf_candidates[zero_sales_mask]) > 0 else pd.DataFrame()
        
        rf_candidates['original_stock'] = rf_candidates['SaSa Net Stock']
        
        # 只保留有庫存的記錄
        rf_candidates = rf_candidates[rf_candidates['available_quantity'] > 0]
        
        return rf_candidates


class ReceiveRules:
    """接收規則（兩種模式通用）"""
    
    @staticmethod
    def identify_urgent_shortage_candidates(df: pd.DataFrame) -> pd.DataFrame:
        """
        識別緊急缺貨補貨候選
        
        Args:
            df: 庫存數據框
            
        Returns:
            緊急缺貨補貨候選數據框
        """
        rf_candidates = df[df['RP Type'] == 'RF'].copy()
        
        if len(rf_candidates) == 0:
            return pd.DataFrame()
        
        # 條件篩選
        condition1 = rf_candidates['SaSa Net Stock'] + rf_candidates['Pending Received'] == 0
        condition2 = rf_candidates['Effective Sold Qty'] > 0
        
        urgent_candidates = rf_candidates[condition1 & condition2].copy()
        
        if len(urgent_candidates) == 0:
            return pd.DataFrame()
        
        # 需求數量：Safety Stock
        urgent_candidates['demand_quantity'] = urgent_candidates['Safety Stock']
        urgent_candidates['receive_priority'] = ReceivePriority.URGENT_SHORTAGE.value
        urgent_candidates['original_stock'] = urgent_candidates['SaSa Net Stock']
        
        return urgent_candidates
    
    @staticmethod
    def identify_potential_shortage_candidates(df: pd.DataFrame) -> pd.DataFrame:
        """
        識別潛在缺貨補貨候選
        
        Args:
            df: 庫存數據框
            
        Returns:
            潛在缺貨補貨候選數據框
        """
        rf_candidates = df[df['RP Type'] == 'RF'].copy()
        
        if len(rf_candidates) == 0:
            return pd.DataFrame()
        
        # 計算每個商品在各店鋪的有效銷量排名
        rf_candidates['max_sold_per_article'] = rf_candidates.groupby('Article')['Effective Sold Qty'].transform('max')
        
        # 條件篩選
        condition1 = rf_candidates['SaSa Net Stock'] + rf_candidates['Pending Received'] < rf_candidates['Safety Stock']
        condition2 = rf_candidates['Effective Sold Qty'] == rf_candidates['max_sold_per_article']
        
        potential_candidates = rf_candidates[condition1 & condition2].copy()
        
        if len(potential_candidates) == 0:
            return pd.DataFrame()
        
        # 需求數量：Safety Stock - (SaSa Net Stock + Pending Received)
        total_stock = potential_candidates['SaSa Net Stock'] + potential_candidates['Pending Received']
        potential_candidates['demand_quantity'] = potential_candidates['Safety Stock'] - total_stock
        potential_candidates['receive_priority'] = ReceivePriority.POTENTIAL_SHORTAGE.value
        potential_candidates['original_stock'] = potential_candidates['SaSa Net Stock']
        
        return potential_candidates

class MatchingAlgorithm:
    """匹配算法"""
    
    # 匹配優先級順序
    MATCHING_PRIORITY = [
        (TransferType.ND_TRANSFER, ReceivePriority.URGENT_SHORTAGE),
        (TransferType.ND_TRANSFER, ReceivePriority.POTENTIAL_SHORTAGE),
        (TransferType.RF_SURPLUS_TRANSFER, ReceivePriority.URGENT_SHORTAGE),
        (TransferType.RF_SURPLUS_TRANSFER, ReceivePriority.POTENTIAL_SHORTAGE),
        (TransferType.RF_ENHANCED_TRANSFER, ReceivePriority.URGENT_SHORTAGE),
        (TransferType.RF_ENHANCED_TRANSFER, ReceivePriority.POTENTIAL_SHORTAGE),
        (TransferType.C_COMPLETE_TRANSFER, ReceivePriority.URGENT_SHORTAGE),
        (TransferType.C_COMPLETE_TRANSFER, ReceivePriority.POTENTIAL_SHORTAGE)
    ]
    
    @staticmethod
    def can_transfer_between_sites(transfer_site: str, receive_site: str) -> bool:
        """
        檢查店鋪之間是否可以轉貨
        
        Args:
            transfer_site: 轉出店鋪
            receive_site: 接收店鋪
        
        Returns:
            是否可以轉貨
        """
        # HA店舖,HB店舖,HC店舖可以出貨去HD店舖
        # HD店舖絕對不能出貨去HA店舖,HB店舖,HC店舖
        # HA店舖,HB店舖,HC店舖轉去HA店舖,HB店舖,HC店舖需限制同組OM, 但轉去HD店舖不受OM限制
        
        transfer_prefix = transfer_site[:2] if len(transfer_site) >= 2 else ""
        receive_prefix = receive_site[:2] if len(receive_site) >= 2 else ""
        
        # HD店舖絕對不能出貨去HA店舖,HB店舖,HC店舖
        if transfer_prefix == "HD" and receive_prefix in ["HA", "HB", "HC"]:
            return False
        
        # 其他情況都允許
        return True
    
    @staticmethod
    def check_om_restriction(transfer_om: str, transfer_site: str, receive_site: str,
                           transfer_candidates: pd.DataFrame, receive_candidates: pd.DataFrame) -> bool:
        """
        檢查OM限制
        
        Args:
            transfer_om: 轉出店鋪的OM
            transfer_site: 轉出店鋪
            receive_site: 接收店鋪
            transfer_candidates: 轉出候選數據框
            receive_candidates: 接收候選數據框
        
        Returns:
            是否符合OM限制
        """
        transfer_prefix = transfer_site[:2] if len(transfer_site) >= 2 else ""
        receive_prefix = receive_site[:2] if len(receive_site) >= 2 else ""
        
        # HA店舖,HB店舖,HC店舖轉去HA店舖,HB店舖,HC店舖需限制同組OM
        if (transfer_prefix in ["HA", "HB", "HC"] and
            receive_prefix in ["HA", "HB", "HC"]):
            # 需要檢查接收店鋪的OM是否與轉出店鋪相同
            receive_sites = receive_candidates[receive_candidates['Site'] == receive_site]
            if len(receive_sites) > 0:
                receive_om = receive_sites['OM'].iloc[0]
                if transfer_om != receive_om:
                    return False
        
        # 轉去HD店舖不受OM限制
        return True
    
    @staticmethod
    def execute_matching(transfer_candidates: pd.DataFrame, receive_candidates: pd.DataFrame) -> List[Dict]:
        """
        執行匹配算法
        
        Args:
            transfer_candidates: 轉出候選數據框
            receive_candidates: 接收候選數據框
            
        Returns:
            調貨建議列表
        """
        if len(transfer_candidates) == 0 or len(receive_candidates) == 0:
            return []
        
        # 按商品分組處理
        articles = transfer_candidates['Article'].unique()
        transfer_recommendations = []
        
        for article in articles:
            # 獲取該商品的轉出和接收候選
            article_transfer = transfer_candidates[transfer_candidates['Article'] == article].copy()
            article_receive = receive_candidates[receive_candidates['Article'] == article].copy()
            
            if len(article_transfer) == 0 or len(article_receive) == 0:
                continue
            
            # 計算總供給和總需求
            total_available = article_transfer['available_quantity'].sum()
            total_demand = article_receive['demand_quantity'].sum()
            
            # 如果總供給大於總需求，則優先從銷量最差的店鋪轉出
            if total_available > total_demand:
                # 按銷量從低到高排序（銷量最差的排前面）
                article_transfer_sorted = article_transfer.sort_values(by='Effective Sold Qty', ascending=True)
            else:
                # 否則按原優先級處理
                article_transfer_sorted = article_transfer.copy()

            # 按優先級順序進行匹配
            # 注意：當供給大於需求時，優先級由銷量排序決定
            for _, transfer_row in article_transfer_sorted.iterrows():
                # 檢查是否還有需求
                remaining_demand = article_receive['demand_quantity'].sum()
                if remaining_demand <= 0:
                    break
                
                # 按優先級順序尋找接收候選
                for transfer_type, receive_priority in MatchingAlgorithm.MATCHING_PRIORITY:
                    # 檢查當前轉出候選是否符合當前優先級
                    if transfer_row['transfer_type'] != transfer_type.value:
                        continue
                    
                    # 獲取當前優先級的接收候選
                    current_receive = article_receive[article_receive['receive_priority'] == receive_priority.value].copy()
                    
                    if len(current_receive) == 0:
                        continue
                    
                    # 按需求量從大到小排序接收候選，優先滿足需求大的店鋪
                    current_receive_sorted = current_receive.sort_values(by='demand_quantity', ascending=False)
                    
                    # 執行匹配
                    for _, receive_row in current_receive_sorted.iterrows():
                        # 檢查店鋪之間是否可以轉貨
                        if not MatchingAlgorithm.can_transfer_between_sites(
                            transfer_row['Site'], receive_row['Site']
                        ):
                            continue
                        
                        # 檢查OM限制
                        if not MatchingAlgorithm.check_om_restriction(
                            transfer_row['OM'], transfer_row['Site'], receive_row['Site'],
                            article_transfer, article_receive
                        ):
                            continue
                        
                        # 計算轉移數量
                        transfer_qty = int(min(
                            transfer_row['available_quantity'],
                            receive_row['demand_quantity']
                        ))
                        
                        # 確保轉出數量不超過轉出店鋪的SaSa Net Stock
                        transfer_qty = min(transfer_qty, transfer_row['SaSa Net Stock'])
                        
                        if transfer_qty <= 0:
                            continue
                        
                        # 調貨數量優化：如果只有1件，嘗試調高到2件
                        if transfer_qty == 1:
                            # 檢查是否可以增加到2件
                            # 必須同時滿足：available_quantity >= 2, SaSa Net Stock >= 2, 且 demand_quantity >= 2
                            if (transfer_row['available_quantity'] >= 2 and
                                transfer_row['SaSa Net Stock'] >= 2 and
                                receive_row['demand_quantity'] >= 2):
                                transfer_qty = 2
                        
                        # 創建調貨建議
                        recommendation = MatchingAlgorithm.create_recommendation(
                            transfer_row, receive_row, transfer_qty
                        )
                        
                        transfer_recommendations.append(recommendation)
                        
                        # 更新轉出和接收候選的剩餘數量
                        transfer_candidates.at[transfer_row.name, 'available_quantity'] -= transfer_qty
                        receive_candidates.at[receive_row.name, 'demand_quantity'] -= transfer_qty
                        article_transfer.at[transfer_row.name, 'available_quantity'] -= transfer_qty
                        article_receive.at[receive_row.name, 'demand_quantity'] -= transfer_qty
                        
                        # 直接更新當前轉出行的可用數量，確保下次循環使用最新值
                        transfer_row['available_quantity'] -= transfer_qty
                        
                        # 如果轉出店鋪已無可轉數量，則跳出內層循環
                        if transfer_row['available_quantity'] <= 0:
                            break
        
        return transfer_recommendations
    
    @staticmethod
    def create_recommendation(transfer_row: pd.Series, receive_row: pd.Series, transfer_qty: int) -> Dict:
        """
        創建調貨建議
        
        Args:
            transfer_row: 轉出候選行
            receive_row: 接收候選行
            transfer_qty: 轉移數量
            
        Returns:
            調貨建議字典
        """
        # 計算轉出後庫存
        transfer_after_stock = transfer_row['original_stock'] - transfer_qty
        
        # 獲取店鋪類別
        transfer_site_type = transfer_row['Site'][:2] if len(transfer_row['Site']) >= 2 else ""
        receive_site_type = receive_row['Site'][:2] if len(receive_row['Site']) >= 2 else ""
        
        # 創建備註，包含店鋪類別信息
        if transfer_row['transfer_type'] == TransferType.ND_TRANSFER.value:
            transfer_type_desc = "ND轉出"
            if transfer_site_type == "HD":
                remark = f"ND轉出 → HD店鋪：澳門優先，HD店鋪為主要目標"
            else:
                remark = f"ND轉出 → {receive_site_type}店鋪：標準轉出"
        elif transfer_row['transfer_type'] == TransferType.RF_SURPLUS_TRANSFER.value:
            transfer_type_desc = "RF過剩轉出"
            if receive_site_type == "HD":
                remark = f"RF過剩轉出 → HD店鋪：澳門優先，庫存充足轉出"
            else:
                remark = f"RF過剩轉出 → {receive_site_type}店鋪：標準轉出"
        elif transfer_row['transfer_type'] == TransferType.RF_ENHANCED_TRANSFER.value:
            transfer_type_desc = "RF加強轉出"
            if receive_site_type == "HD":
                remark = f"RF加強轉出 → HD店鋪：澳門優先，加強轉出支援"
            else:
                remark = f"RF加強轉出 → {receive_site_type}店鋪：標準加強轉出"
        elif transfer_row['transfer_type'] == TransferType.C_COMPLETE_TRANSFER.value:
            transfer_type_desc = "C模式全量轉出"
            if receive_site_type == "HD":
                remark = f"C模式全量轉出 → HD店鋪：澳門優先，全量轉出支援"
            else:
                remark = f"C模式全量轉出 → {receive_site_type}店鋪：標準全量轉出"
        else:
            transfer_type_desc = "未知轉出類型"
            remark = f"未知轉出類型 → {receive_site_type}店鋪"
        
        # 創建詳細說明，包含調貨邏輯和計算方式
        if transfer_row['transfer_type'] == TransferType.ND_TRANSFER.value:
            transfer_logic = "ND店鋪轉出：ND類型店鋪可轉出全部淨庫存，適合關閉或調整店鋪"
            transfer_calc = "轉出數量 = 全部淨庫存"
        elif transfer_row['transfer_type'] == TransferType.RF_SURPLUS_TRANSFER.value:
            transfer_logic = "RF過剩轉出：庫存充足的RF店鋪，轉出後剩餘庫存不低於安全庫存"
            total_stock = transfer_row['SaSa Net Stock'] + transfer_row['Pending Received']
            basic_transferable = total_stock - transfer_row['Safety Stock']
            upper_limit = max(total_stock * 0.4, 2)
            transfer_calc = f"轉出數量 = min(基礎可轉出{basic_transferable}, 上限控制{upper_limit})"
        elif transfer_row['transfer_type'] == TransferType.RF_ENHANCED_TRANSFER.value:
            transfer_logic = "RF加強轉出：庫存超過MOQ的RF店鋪，轉出後可能低於安全庫存"
            total_stock = transfer_row['SaSa Net Stock'] + transfer_row['Pending Received']
            basic_transferable = total_stock - transfer_row['MOQ']
            upper_limit = max(total_stock * 0.8, 2)
            transfer_calc = f"轉出數量 = min(基礎可轉出{basic_transferable}, 上限控制{upper_limit})"
        elif transfer_row['transfer_type'] == TransferType.C_COMPLETE_TRANSFER.value:
            transfer_logic = "C模式全量轉出：同OM同Article中銷量最少的店鋪可轉出全部庫存"
            transfer_calc = "轉出數量 = 全部淨庫存（銷量為0的店鋪）"
        else:
            transfer_logic = "未知轉出類型"
            transfer_calc = "未知計算方式"
        
        if receive_row['receive_priority'] == ReceivePriority.URGENT_SHORTAGE.value:
            receive_logic = "緊急缺貨補貨：完全無庫存+在途且曾有銷售記錄的RF店鋪"
            if receive_site_type == "HD":
                receive_logic += "，澳門優先：HD店鋪為重點補貨對象"
            receive_calc = f"需求數量 = 安全庫存 {transfer_row['Safety Stock']}"
        elif receive_row['receive_priority'] == ReceivePriority.POTENTIAL_SHORTAGE.value:
            total_stock = receive_row['SaSa Net Stock'] + receive_row['Pending Received']
            demand_qty = transfer_row['Safety Stock'] - total_stock
            receive_logic = "潛在缺貨補貨：庫存不足且有效銷量為最高值的RF店鋪"
            if receive_site_type == "HD":
                receive_logic += "，澳門優先：HD店鋪為重點補貨對象"
            receive_calc = f"需求數量 = 安全庫存 {transfer_row['Safety Stock']} - 總庫存 {total_stock} = {demand_qty}"
        else:
            receive_logic = "未知接收優先級"
            receive_calc = "未知需求計算"
        
        # 創建澳門優先說明
        if receive_site_type == "HD" or transfer_site_type == "HD":
            macau_priority = "澳門優先：HD店鋪為澳門地區重點支援對象，優先考慮其庫存需求和銷售特點"
        else:
            macau_priority = "標準調貨：按常規優先級處理"
        
        # 創建詳細說明
        notes = f"{transfer_type_desc} → {receive_row['receive_priority']} | {transfer_logic} | {receive_calc} | {transfer_calc} | {macau_priority}"
        
        recommendation = {
            'Article': transfer_row['Article'],
            'Product Desc': transfer_row['Article Description'],
            'Transfer OM': transfer_row['OM'],
            'Transfer Site': transfer_row['Site'],
            'Receive OM': receive_row['OM'],
            'Receive Site': receive_row['Site'],
            'Transfer Qty': int(transfer_qty),
            'Transfer Site Original Stock': transfer_row['original_stock'],
            'Transfer Site After Transfer Stock': transfer_after_stock,
            'Transfer Site Safety Stock': transfer_row['Safety Stock'],
            'Transfer Site MOQ': transfer_row['MOQ'],
            'Transfer Site Last Month Sold Qty': transfer_row['Last Month Sold Qty'],
            'Transfer Site MTD Sold Qty': transfer_row['MTD Sold Qty'],
            'Receive Site Last Month Sold Qty': receive_row['Last Month Sold Qty'],
            'Receive Site MTD Sold Qty': receive_row['MTD Sold Qty'],
            'Receive Original Stock': receive_row['original_stock'],
            'Remark': remark,
            'Notes': notes
        }
        
        return recommendation

class QualityChecker:
    """質量檢查器，確保調貨建議的準確性和合理性"""
    
    @staticmethod
    def check_recommendations(recommendations: List[Dict]) -> Tuple[bool, List[str]]:
        """
        檢查調貨建議質量
        
        Args:
            recommendations: 調貨建議列表
            
        Returns:
            (是否通過檢查, 錯誤信息列表)
        """
        errors = []
        
        for i, rec in enumerate(recommendations):
            # 檢查轉出與接收的Article必須完全一致
            # (這個檢查在創建建議時已經確保)
            
            # 檢查Transfer Qty必須為正整數
            transfer_qty = rec['Transfer Qty']
            # 確保是整數類型
            if not isinstance(transfer_qty, int):
                try:
                    transfer_qty = int(float(transfer_qty))
                    # 更新建議中的值
                    rec['Transfer Qty'] = transfer_qty
                except (ValueError, TypeError):
                    errors.append(f"建議{i+1}: Transfer Qty必須為正整數")
                    continue
            
            if transfer_qty <= 0:
                errors.append(f"建議{i+1}: Transfer Qty必須為正整數")
            
            # 檢查Transfer Qty不得超過轉出店鋪的原始SaSa Net Stock
            if rec['Transfer Qty'] > rec['Transfer Site Original Stock']:
                errors.append(f"建議{i+1}: Transfer Qty超過轉出店鋪原始庫存")
            
            # 檢查Transfer Site和Receive Site不能相同
            if rec['Transfer Site'] == rec['Receive Site']:
                errors.append(f"建議{i+1}: 轉出店鋪和接收店鋪不能相同")
            
            # 檢查最終輸出的Article欄位必須是12位文本格式
            if not isinstance(rec['Article'], str) or len(rec['Article']) != 12:
                errors.append(f"建議{i+1}: Article必須是12位文本格式")
        
        is_valid = len(errors) == 0
        return is_valid, errors

class BusinessLogic:
    """業務邏輯主類，協調各個業務邏輯組件"""
    
    def __init__(self):
        self.transfer_mode_a = TransferModeA()
        self.transfer_mode_b = TransferModeB()
        self.transfer_mode_c = TransferModeC()
        self.receive_rules = ReceiveRules()
        self.matching_algorithm = MatchingAlgorithm()
        self.quality_checker = QualityChecker()
    
    def generate_transfer_recommendations(self, df: pd.DataFrame, mode: str) -> Tuple[bool, Union[List[Dict], str], Dict]:
        """
        生成調貨建議
        
        Args:
            df: 庫存數據框
            mode: 轉貨模式 ("A", "B" 或 "C")
            
        Returns:
            (是否成功, 調貨建議列表或錯誤信息, 統計信息)
        """
        try:
            if mode not in ["A", "B", "C"]:
                return False, "不支持的轉貨模式，請選擇A、B或C", {}
            
            # 識別轉出候選
            if mode == "A":
                # A模式
                nd_transfer_candidates = self.transfer_mode_a.identify_nd_transfer_candidates(df)
                rf_transfer_candidates = self.transfer_mode_a.identify_rf_surplus_transfer_candidates(df)
            elif mode == "B":
                # B模式
                nd_transfer_candidates = self.transfer_mode_b.identify_nd_transfer_candidates(df)
                rf_transfer_candidates = self.transfer_mode_b.identify_rf_transfer_candidates(df)
            else:
                # C模式
                nd_transfer_candidates = self.transfer_mode_c.identify_nd_transfer_candidates(df)
                rf_transfer_candidates = self.transfer_mode_c.identify_rf_transfer_candidates(df)
            
            # 合併轉出候選
            transfer_candidates = pd.concat([nd_transfer_candidates, rf_transfer_candidates], ignore_index=True)
            
            # 識別接收候選
            urgent_receive_candidates = self.receive_rules.identify_urgent_shortage_candidates(df)
            potential_receive_candidates = self.receive_rules.identify_potential_shortage_candidates(df)
            
            # 合併接收候選
            receive_candidates = pd.concat([urgent_receive_candidates, potential_receive_candidates], ignore_index=True)
            
            # 執行匹配算法
            recommendations = self.matching_algorithm.execute_matching(transfer_candidates, receive_candidates)
            
            # 質量檢查
            is_valid, errors = self.quality_checker.check_recommendations(recommendations)
            if not is_valid:
                error_msg = f"質量檢查失敗: {'; '.join(errors)}"
                return False, error_msg, {}
            
            # 生成統計信息
            stats = self.generate_statistics(recommendations, df)
            
            return True, recommendations, stats
            
        except Exception as e:
            logger.error(f"生成調貨建議時發生錯誤: {str(e)}")
            return False, f"生成調貨建議時發生錯誤: {str(e)}", {}
    
    def generate_statistics(self, recommendations: List[Dict], original_data: pd.DataFrame) -> Dict:
        """
        生成統計信息
        
        Args:
            recommendations: 調貨建議列表
            original_data: 原始數據框
            
        Returns:
            統計信息字典
        """
        if not recommendations:
            return {
                'total_recommendations': 0,
                'total_transfer_quantity': 0,
                'unique_articles': 0,
                'unique_transfer_sites': 0,
                'unique_receive_sites': 0,
                'transfer_type_stats': {},
                'receive_priority_stats': {}
            }
        
        # 基本統計
        total_recommendations = len(recommendations)
        total_transfer_quantity = sum(rec['Transfer Qty'] for rec in recommendations)
        unique_articles = len(set(rec['Article'] for rec in recommendations))
        unique_transfer_sites = len(set(rec['Transfer Site'] for rec in recommendations))
        unique_receive_sites = len(set(rec['Receive Site'] for rec in recommendations))
        
        # 轉出類型統計
        transfer_type_stats = {}
        for rec in recommendations:
            transfer_type = rec['Remark']
            if transfer_type not in transfer_type_stats:
                transfer_type_stats[transfer_type] = {'count': 0, 'quantity': 0}
            transfer_type_stats[transfer_type]['count'] += 1
            transfer_type_stats[transfer_type]['quantity'] += rec['Transfer Qty']
        
        # 接收優先級統計
        receive_priority_stats = {}
        for rec in recommendations:
            receive_priority = rec['Notes'].split(' → ')[1] if ' → ' in rec['Notes'] else '未知'
            if receive_priority not in receive_priority_stats:
                receive_priority_stats[receive_priority] = {'count': 0, 'quantity': 0}
            receive_priority_stats[receive_priority]['count'] += 1
            receive_priority_stats[receive_priority]['quantity'] += rec['Transfer Qty']
        
        return {
            'total_recommendations': total_recommendations,
            'total_transfer_quantity': total_transfer_quantity,
            'unique_articles': unique_articles,
            'unique_transfer_sites': unique_transfer_sites,
            'unique_receive_sites': unique_receive_sites,
            'transfer_type_stats': transfer_type_stats,
            'receive_priority_stats': receive_priority_stats
        }


# 測試代碼
if __name__ == "__main__":
    # 測試業務邏輯
    from data_processor import DataProcessor
    
    # 創建測試數據
    processor = DataProcessor()
    test_data = processor.generate_mock_data(num_articles=5, seed=42)
    
    # 創建業務邏輯實例
    business_logic = BusinessLogic()
    
    # 測試A模式
    success, result_a, stats_a = business_logic.generate_transfer_recommendations(test_data, "A")
    print(f"A模式測試結果: {success}")
    if success:
        print(f"建議數量: {len(result_a)}")
        print(f"統計信息: {stats_a}")
    else:
        print(f"錯誤: {result_a}")
    
    # 測試B模式
    success, result_b, stats_b = business_logic.generate_transfer_recommendations(test_data, "B")
    print(f"B模式測試結果: {success}")
    if success:
        print(f"建議數量: {len(result_b)}")
        print(f"統計信息: {stats_b}")
    else:
        print(f"錯誤: {result_b}")

"""
Streamlit主應用程序 v1.9
庫存調貨建議系統v1.0的用戶界面和應用程序流程控制
"""

import streamlit as st
import pandas as pd
import numpy as np
import io
import base64
from datetime import datetime
import os
import sys

# 導入自定義模塊
from data_processor import DataProcessor
from business_logic import BusinessLogic
from excel_generator import ExcelGenerator

# 設置頁面配置
st.set_page_config(
    page_title="庫存調貨建議系統 v1.0",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 自定義CSS
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 1rem;
    }
    .sub-header {
        font-size: 1.5rem;
        color: #2ca02c;
        margin-bottom: 1rem;
    }
    .success-message {
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        border-radius: 5px;
        padding: 10px;
        margin-bottom: 10px;
    }
    .error-message {
        background-color: #f8d7da;
        border: 1px solid #f5c6cb;
        border-radius: 5px;
        padding: 10px;
        margin-bottom: 10px;
    }
    .info-message {
        background-color: #d1ecf1;
        border: 1px solid #bee5eb;
        border-radius: 5px;
        padding: 10px;
        margin-bottom: 10px;
    }
    .warning-message {
        background-color: #fff3cd;
        border: 1px solid #ffeaa7;
        border-radius: 5px;
        padding: 10px;
        margin-bottom: 10px;
    }
</style>
""", unsafe_allow_html=True)

class InventoryTransferApp:
    """庫存調貨建議系統主應用類"""
    
    def __init__(self):
        self.data_processor = DataProcessor()
        self.business_logic = BusinessLogic()
        self.excel_generator = ExcelGenerator()
        
        # 初始化會話狀態
        if 'processed_data' not in st.session_state:
            st.session_state.processed_data = None
        if 'recommendations' not in st.session_state:
            st.session_state.recommendations = None
        if 'stats' not in st.session_state:
            st.session_state.stats = None
        if 'mode' not in st.session_state:
            st.session_state.mode = "A"
    
    def render_header(self):
        """渲染頁面標題"""
        st.markdown('<h1 class="main-header">📦 庫存調貨建議系統 v1.0</h1>', unsafe_allow_html=True)
        st.markdown("---")
    
    def render_sidebar(self):
        """渲染側邊欄"""
        st.sidebar.markdown('<h2 class="sub-header">系統設置</h2>', unsafe_allow_html=True)
        
        # 模式選擇
        st.sidebar.subheader("🔧 轉貨模式選擇")
        mode_options = {
            "A": "A模式 (保守轉貨)",
            "B": "B模式 (加強轉貨)",
            "C": "C模式 (全量轉貨)"
        }
        
        selected_mode = st.sidebar.selectbox(
            "選擇轉貨模式",
            options=list(mode_options.keys()),
            format_func=lambda x: mode_options[x],
            index=list(mode_options.keys()).index(st.session_state.mode) if st.session_state.mode in mode_options else 0
        )
        
        st.session_state.mode = selected_mode
        
        # 模式說明
        if selected_mode == "A":
            st.sidebar.info("""
            **A模式 (保守轉貨)**
            - 轉出後剩餘庫存不低於安全庫存
            - 轉出類型為RF過剩轉出
            - 適合保守的庫存管理策略
            """)
        elif selected_mode == "B":
            st.sidebar.info("""
            **B模式 (加強轉貨)**
            - 轉出後剩餘庫存可能低於安全庫存
            - 轉出類型包括RF過剩轉出和RF加強轉出
            - 適合積極的庫存優化策略
            """)
        else:
            st.sidebar.info("""
            **C模式 (全量轉貨)**
            - 忽視A模式及B模式的限制
            - ND Shop可以轉去ND Shop
            - 需要限制同一個OM組別及同一個Article
            - 轉出店舖的銷售量必須為同組最少
            - 接收店舖的銷售量必須為同組最多
            - 轉出店舖的銷售量如果為0件，轉出數量可全數轉出
            """)
        
        st.sidebar.markdown("---")
        
        # 系統信息
        st.sidebar.subheader("ℹ️ 系統信息")
        st.sidebar.text(f"當前時間: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        st.sidebar.text("版本: v1.0")
        
        # 重置按鈕
        if st.sidebar.button("🔄 重置系統", type="secondary"):
            self.reset_system()
    
    def render_data_upload_section(self):
        """渲染數據上傳區域"""
        st.markdown('<h2 class="sub-header">📂 數據上傳</h2>', unsafe_allow_html=True)
        
        # 創建兩列布局
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("上傳Excel文件")
            uploaded_file = st.file_uploader(
                "選擇Excel文件",
                type=['xlsx'],
                help="請上傳包含庫存數據的Excel文件(.xlsx格式)"
            )
            
            if uploaded_file is not None:
                # 顯示文件信息
                st.success(f"文件已上傳: {uploaded_file.name}")
                
                # 處理上傳的文件
                with st.spinner("正在處理文件..."):
                    # 保存上傳的文件到臨時位置
                    temp_dir = "temp"
                    os.makedirs(temp_dir, exist_ok=True)
                    temp_path = os.path.join(temp_dir, uploaded_file.name)
                    
                    with open(temp_path, "wb") as f:
                        f.write(uploaded_file.getbuffer())
                    
                    # 處理文件
                    success, result, stats = self.data_processor.process_uploaded_file(temp_path)
                    
                    if success:
                        st.session_state.processed_data = result
                        st.success("✅ 文件處理成功！")
                        
                        # 顯示處理統計
                        self.display_processing_stats(stats)
                    else:
                        st.error(f"❌ 文件處理失敗: {result}")
        
        with col2:
            st.subheader("生成模擬數據")
            st.info("如果沒有真實數據，可以使用模擬數據進行測試")
            
            # 模擬數據參數
            num_articles = st.slider("商品數量", min_value=5, max_value=50, value=10)
            seed = st.number_input("隨機種子", value=42, step=1)
            
            if st.button("🎲 生成模擬數據", type="primary"):
                with st.spinner("正在生成模擬數據..."):
                    mock_data = self.data_processor.generate_mock_data(num_articles, seed)
                    st.session_state.processed_data = mock_data
                    st.success("✅ 模擬數據生成成功！")
                    
                    # 顯示數據統計
                    stats = {
                        'total_rows': len(mock_data),
                        'unique_articles': mock_data['Article'].nunique(),
                        'unique_sites': mock_data['Site'].nunique(),
                        'nd_sites': len(mock_data[mock_data['RP Type'] == 'ND']['Site'].unique()),
                        'rf_sites': len(mock_data[mock_data['RP Type'] == 'RF']['Site'].unique()),
                        'total_stock': mock_data['SaSa Net Stock'].sum(),
                        'total_safety_stock': mock_data['Safety Stock'].sum()
                    }
                    self.display_processing_stats(stats)
    
    def display_processing_stats(self, stats):
        """顯示數據處理統計"""
        st.markdown("#### 📊 數據統計")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.metric("總記錄數", stats.get('total_rows', 0))
            st.metric("商品數量", stats.get('unique_articles', 0))
        
        with col2:
            st.metric("店鋪數量", stats.get('unique_sites', 0))
            st.metric("ND店鋪", stats.get('nd_sites', 0))
        
        with col3:
            st.metric("RF店鋪", stats.get('rf_sites', 0))
            st.metric("總庫存", stats.get('total_stock', 0))
    
    def render_data_preview_section(self):
        """渲染數據預覽區域"""
        if st.session_state.processed_data is not None:
            st.markdown('<h2 class="sub-header">👀 數據預覽</h2>', unsafe_allow_html=True)
            
            # 顯示數據概覽
            st.write(f"數據形狀: {st.session_state.processed_data.shape}")
            
            # 顯示前幾行數據
            st.dataframe(st.session_state.processed_data.head(10))
            
            # 顯示數據統計
            if st.checkbox("顯示詳細統計"):
                st.write("#### 數值欄位統計")
                st.dataframe(st.session_state.processed_data.describe())
                
                st.write("#### RP Type分布")
                rp_type_counts = st.session_state.processed_data['RP Type'].value_counts()
                st.bar_chart(rp_type_counts)
    
    def render_analysis_section(self):
        """渲染分析區域"""
        if st.session_state.processed_data is not None:
            st.markdown('<h2 class="sub-header">🔍 調貨分析</h2>', unsafe_allow_html=True)
            
            # 生成調貨建議按鈕
            if st.button("🚀 生成調貨建議", type="primary", use_container_width=True):
                with st.spinner("正在生成調貨建議..."):
                    success, result, stats = self.business_logic.generate_transfer_recommendations(
                        st.session_state.processed_data, st.session_state.mode
                    )
                    
                    if success:
                        st.session_state.recommendations = result
                        st.session_state.stats = stats
                        st.success("✅ 調貨建議生成成功！")
                    else:
                        st.error(f"❌ 生成調貨建議失敗: {result}")
                        st.session_state.recommendations = None
                        st.session_state.stats = None
    
    def render_results_section(self):
        """渲染結果展示區域"""
        if st.session_state.recommendations is not None:
            st.markdown('<h2 class="sub-header">📋 調貨建議結果</h2>', unsafe_allow_html=True)
            
            # 顯示統計摘要
            self.display_recommendation_stats()
            
            # 顯示調貨建議詳情
            self.display_recommendation_details()
            
            # 顯示圖表
            self.display_charts()
            
            # Excel下載
            self.render_download_section()
    
    def display_recommendation_stats(self):
        """顯示調貨建議統計"""
        stats = st.session_state.stats
        
        st.markdown("#### 📈 統計摘要")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.metric("總建議數", stats.get('total_recommendations', 0))
            st.metric("涉及商品", stats.get('unique_articles', 0))
        
        with col2:
            st.metric("總轉移數量", stats.get('total_transfer_quantity', 0))
            st.metric("轉出店鋪", stats.get('unique_transfer_sites', 0))
        
        with col3:
            st.metric("接收店鋪", stats.get('unique_receive_sites', 0))
        
        # 轉出類型統計
        if 'transfer_type_stats' in stats and stats['transfer_type_stats']:
            st.markdown("##### 轉出類型分布")
            transfer_df = pd.DataFrame(stats['transfer_type_stats']).T
            st.dataframe(transfer_df)
        
        # 接收優先級統計
        if 'receive_priority_stats' in stats and stats['receive_priority_stats']:
            st.markdown("##### 接收優先級分布")
            receive_df = pd.DataFrame(stats['receive_priority_stats']).T
            st.dataframe(receive_df)
    
    def display_recommendation_details(self):
        """顯示調貨建議詳情"""
        st.markdown("#### 📝 調貨建議詳情")
        
        recommendations = st.session_state.recommendations
        
        if recommendations:
            # 轉換為數據框
            df = pd.DataFrame(recommendations)
            
            # 顯示數據表格
            st.dataframe(df, use_container_width=True)
            
            # 提供搜索和篩選功能
            if st.checkbox("啟用搜索和篩選"):
                col1, col2 = st.columns(2)
                
                with col1:
                    # 按商品搜索
                    search_article = st.text_input("搜索商品編號")
                    
                with col2:
                    # 按店鋪篩選
                    transfer_sites = df['Transfer Site'].unique().tolist()
                    selected_transfer_site = st.selectbox("篩選轉出店鋪", ["全部"] + transfer_sites)
                
                # 應用篩選
                filtered_df = df.copy()
                
                if search_article:
                    filtered_df = filtered_df[filtered_df['Article'].str.contains(search_article, case=False)]
                
                if selected_transfer_site != "全部":
                    filtered_df = filtered_df[filtered_df['Transfer Site'] == selected_transfer_site]
                
                st.dataframe(filtered_df, use_container_width=True)
        else:
            st.info("沒有生成調貨建議")
    
    def display_charts(self):
        """顯示圖表"""
        stats = st.session_state.stats
        
        st.markdown("#### 📊 可視化分析")
        
        col1, col2 = st.columns(2)
        
        with col1:
            # 轉出類型圖表
            if 'transfer_type_stats' in stats and stats['transfer_type_stats']:
                transfer_data = []
                for transfer_type, type_stats in stats['transfer_type_stats'].items():
                    transfer_data.append({
                        '轉出類型': transfer_type,
                        '建議數量': type_stats.get('count', 0),
                        '轉移數量': type_stats.get('quantity', 0)
                    })
                
                transfer_df = pd.DataFrame(transfer_data)
                st.bar_chart(transfer_df.set_index('轉出類型')['建議數量'])
        
        with col2:
            # 接收優先級圖表
            if 'receive_priority_stats' in stats and stats['receive_priority_stats']:
                receive_data = []
                for priority, priority_stats in stats['receive_priority_stats'].items():
                    receive_data.append({
                        '接收優先級': priority,
                        '建議數量': priority_stats.get('count', 0),
                        '轉移數量': priority_stats.get('quantity', 0)
                    })
                
                receive_df = pd.DataFrame(receive_data)
                st.bar_chart(receive_df.set_index('接收優先級')['建議數量'])
    
    def render_download_section(self):
        """渲染下載區域"""
        st.markdown("#### 💾 下載結果")
        
        col1, col2 = st.columns(2)
        
        with col1:
            # 生成Excel文件
            if st.button("📊 生成Excel文件", type="primary"):
                with st.spinner("正在生成Excel文件..."):
                    success, message, file_path = self.excel_generator.generate_excel_file(
                        st.session_state.recommendations,
                        st.session_state.stats
                    )
                    
                    if success:
                        st.success(message)
                        st.session_state.excel_file_path = file_path
                    else:
                        st.error(message)
        
        with col2:
            # 下載Excel文件
            if 'excel_file_path' in st.session_state and st.session_state.excel_file_path:
                if os.path.exists(st.session_state.excel_file_path):
                    with open(st.session_state.excel_file_path, "rb") as file:
                        btn = st.download_button(
                            label="📥 下載Excel文件",
                            data=file.read(),
                            file_name=os.path.basename(st.session_state.excel_file_path),
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
    
    def reset_system(self):
        """重置系統狀態"""
        keys_to_clear = ['processed_data', 'recommendations', 'stats', 'excel_file_path']
        for key in keys_to_clear:
            if key in st.session_state:
                del st.session_state[key]
        
        # 清理臨時文件
        temp_dir = "temp"
        if os.path.exists(temp_dir):
            import shutil
            shutil.rmtree(temp_dir)
        
        st.success("系統已重置")
        st.experimental_rerun()
    
    def render_footer(self):
        """渲染頁腳"""
        st.markdown("---")
        st.markdown("""
        <div style='text-align: center; color: #666; font-size: 0.8em;'>
        庫存調貨建議系統 v1.0 | 基於Streamlit構建 | © 2025
        </div>
        """, unsafe_allow_html=True)
    
    def run(self):
        """運行應用程序"""
        # 渲染頁面標題
        self.render_header()
        
        # 渲染側邊欄
        self.render_sidebar()
        
        # 渲染主要內容區域
        # 數據上傳區域
        self.render_data_upload_section()
        
        # 數據預覽區域
        self.render_data_preview_section()
        
        # 分析區域
        self.render_analysis_section()
        
        # 結果展示區域
        self.render_results_section()
        
        # 渲染頁腳
        self.render_footer()

# 主程序入口
def main():
    """主函數"""
    try:
        # 創建並運行應用
        app = InventoryTransferApp()
        app.run()
    except Exception as e:
        st.error(f"應用程序運行錯誤: {str(e)}")
        st.error("請檢查系統日誌或聯繫系統管理員")

# 運行主程序
if __name__ == "__main__":
    main()
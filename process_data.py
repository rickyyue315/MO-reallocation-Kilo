#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Data processing script - Directly process uploaded Excel file and generate reports
"""

import pandas as pd
import os
from datetime import datetime
import sys

def load_and_process_data(file_path):
    """Load and process the data file"""
    print("="*50)
    print("Inventory Transfer Recommendation System v1.0 - Data Processing")
    print("="*50)
    
    # Check if data file exists
    if not os.path.exists(file_path):
        print(f"Error: Data file not found {file_path}")
        return None
    
    print(f"Processing data file: {file_path}")
    
    # 1. Load data
    print("\nStep 1: Loading data...")
    try:
        df = pd.read_excel(file_path)
        print(f"Data loaded successfully, {len(df)} rows of records")
        print(f"Data columns: {list(df.columns)}")
        
        # Select relevant columns
        relevant_columns = [
            'Article', 'Article Description', 'Article Long Text (60 Chars)',
            'OM', 'RP Type', 'Site', 'MOQ', 'SaSa Net Stock', 
            'Pending Received', 'Safety Stock', 'Last Month Sold Qty', 'MTD Sold Qty'
        ]
        
        # Check if all relevant columns exist
        missing_cols = [col for col in relevant_columns if col not in df.columns]
        if missing_cols:
            print(f"Warning: Missing columns: {missing_cols}")
            # Try to find alternative column names
            column_mapping = {
                'Article Description': 'Article Long Text (60 Chars)' if 'Article Description' not in df.columns else 'Article Description',
                'Article Long Text (60 Chars)': 'Article Description' if 'Article Long Text (60 Chars)' not in df.columns else 'Article Long Text (60 Chars)'
            }
            
            for col in missing_cols:
                if col in column_mapping and column_mapping[col] in df.columns:
                    df[col] = df[column_mapping[col]]
                    print(f"Using {column_mapping[col]} as {col}")
        
        # Filter to only relevant columns
        available_cols = [col for col in relevant_columns if col in df.columns]
        df = df[available_cols]
        
        # Handle missing values
        df = df.fillna(0)
        
        # Convert Article to string with leading zeros if needed
        if 'Article' in df.columns:
            df['Article'] = df['Article'].astype(str).str.zfill(12)
        
        print("\nData preview:")
        print(df.head())
        
        return df
        
    except Exception as e:
        print(f"Data loading failed: {str(e)}")
        return None

def generate_transfer_recommendations(df, mode='A'):
    """Generate transfer recommendations based on mode"""
    print(f"\nStep 2: Generating transfer recommendations (Mode {mode})...")
    
    recommendations = []
    
    # Group by Article to process each product separately
    for article, group in df.groupby('Article'):
        # Separate ND and RF stores
        nd_stores = group[group['RP Type'] == 'ND']
        rf_stores = group[group['RP Type'] == 'RF']
        
        if len(nd_stores) == 0 or len(rf_stores) == 0:
            continue  # Skip if no ND or RF stores for this article
        
        # Calculate effective sales (using Last Month Sold Qty and MTD Sold Qty)
        for _, store in rf_stores.iterrows():
            store['Effective Sold Qty'] = store['Last Month Sold Qty'] + store['MTD Sold Qty']
        
        # Find stores with highest effective sales
        rf_stores = rf_stores.copy()
        rf_stores['Effective Sold Qty'] = rf_stores['Last Month Sold Qty'] + rf_stores['MTD Sold Qty']
        max_sales = rf_stores['Effective Sold Qty'].max()
        
        # Transfer out rules based on mode
        transfer_out_stores = []
        
        # Priority 1: ND stores
        for _, nd_store in nd_stores.iterrows():
            if nd_store['SaSa Net Stock'] > 0:
                transfer_out_stores.append({
                    'store': nd_store,
                    'available_qty': nd_store['SaSa Net Stock'],
                    'transfer_type': 'ND Transfer Out'
                })
        
        # Priority 2: RF stores based on mode
        for _, rf_store in rf_stores.iterrows():
            total_stock = rf_store['SaSa Net Stock'] + rf_store['Pending Received']
            
            if mode == 'A':  # Conservative mode
                # RF surplus transfer
                if (total_stock > rf_store['Safety Stock'] and 
                    rf_store['Effective Sold Qty'] < max_sales):
                    base_transfer = total_stock - rf_store['Safety Stock']
                    upper_limit = max(total_stock * 0.4, 2)
                    actual_transfer = min(base_transfer, upper_limit)
                    
                    if actual_transfer > 0:
                        transfer_out_stores.append({
                            'store': rf_store,
                            'available_qty': actual_transfer,
                            'transfer_type': 'RF Surplus Transfer Out'
                        })
            
            elif mode == 'B':  # Enhanced mode
                # RF transfer (can go below safety stock)
                if (total_stock > rf_store['MOQ'] and 
                    rf_store['Effective Sold Qty'] < max_sales):
                    base_transfer = total_stock - rf_store['MOQ']
                    upper_limit = max(total_stock * 0.8, 2)
                    actual_transfer = min(base_transfer, upper_limit)
                    
                    if actual_transfer > 0:
                        remaining_stock = total_stock - actual_transfer
                        transfer_type = 'RF Enhanced Transfer Out' if remaining_stock < rf_store['Safety Stock'] else 'RF Surplus Transfer Out'
                        
                        transfer_out_stores.append({
                            'store': rf_store,
                            'available_qty': actual_transfer,
                            'transfer_type': transfer_type
                        })
        
        # Transfer in rules
        transfer_in_stores = []
        
        for _, rf_store in rf_stores.iterrows():
            total_stock = rf_store['SaSa Net Stock'] + rf_store['Pending Received']
            
            # Priority 1: Emergency shortage
            if (total_stock == 0 and rf_store['Effective Sold Qty'] > 0):
                transfer_in_stores.append({
                    'store': rf_store,
                    'needed_qty': rf_store['Safety Stock'],
                    'priority': 1,
                    'reason': 'Emergency Shortage'
                })
            
            # Priority 2: Potential shortage
            elif (total_stock < rf_store['Safety Stock'] and 
                  rf_store['Effective Sold Qty'] == max_sales and max_sales > 0):
                transfer_in_stores.append({
                    'store': rf_store,
                    'needed_qty': rf_store['Safety Stock'] - total_stock,
                    'priority': 2,
                    'reason': 'Potential Shortage'
                })
        
        # Match transfers
        for transfer_out in transfer_out_stores:
            for transfer_in in transfer_in_stores:
                if transfer_out['available_qty'] <= 0:
                    break
                
                transfer_qty = min(transfer_out['available_qty'], transfer_in['needed_qty'])
                
                if transfer_qty > 0:
                    # Apply store restrictions (HA/H/HC to HD allowed, but HD to HA/H/HC not allowed)
                    transfer_site = transfer_out['store']['Site']
                    receive_site = transfer_in['store']['Site']
                    
                    # Check if transfer is allowed
                    transfer_allowed = True
                    
                    # If transfer site is HD and receive site is HA/H/HC, not allowed
                    if transfer_site.startswith('HD') and receive_site.startswith(('HA', 'H', 'HC')):
                        transfer_allowed = False
                    
                    # If both sites are HA/H/HC, check if same OM
                    elif (transfer_site.startswith(('HA', 'H', 'HC')) and 
                          receive_site.startswith(('HA', 'H', 'HC')) and
                          transfer_out['store']['OM'] != transfer_in['store']['OM']):
                        transfer_allowed = False
                    
                    if transfer_allowed:
                        # Optimize transfer quantity (try to increase to 2 if it's 1)
                        if transfer_qty == 1 and transfer_out['available_qty'] >= 2:
                            transfer_qty = 2
                        
                        # Create recommendation
                        recommendation = {
                            'Article': article,
                            'Product Desc': transfer_out['store'].get('Article Description', transfer_out['store'].get('Article Long Text (60 Chars)', '')),
                            'Transfer OM': transfer_out['store']['OM'],
                            'Transfer Site': transfer_out['store']['Site'],
                            'Receive OM': transfer_in['store']['OM'],
                            'Receive Site': transfer_in['store']['Site'],
                            'Transfer Qty': int(transfer_qty),
                            'Transfer Site Original Stock': transfer_out['store']['SaSa Net Stock'],
                            'Transfer Site After Transfer Stock': transfer_out['store']['SaSa Net Stock'] - transfer_qty,
                            'Transfer Site Safety Stock': transfer_out['store']['Safety Stock'],
                            'Transfer Site MOQ': transfer_out['store']['MOQ'],
                            'Transfer Site Last Month Sold Qty': transfer_out['store']['Last Month Sold Qty'],
                            'Transfer Site MTD Sold Qty': transfer_out['store']['MTD Sold Qty'],
                            'Receive Site Last Month Sold Qty': transfer_in['store']['Last Month Sold Qty'],
                            'Receive Site MTD Sold Qty': transfer_in['store']['MTD Sold Qty'],
                            'Receive Original Stock': transfer_in['store']['SaSa Net Stock'],
                            'Remark': transfer_out['transfer_type'],
                            'Notes': f"Transfer reason: {transfer_in['reason']}"
                        }
                        
                        recommendations.append(recommendation)
                        
                        # Update available quantities
                        transfer_out['available_qty'] -= transfer_qty
                        transfer_in['needed_qty'] -= transfer_qty
    
    print(f"Generated {len(recommendations)} transfer recommendations")
    return recommendations

def generate_excel_report(recommendations, mode='A'):
    """Generate Excel report with recommendations"""
    print("\nStep 3: Generating Excel report...")
    
    if not recommendations:
        print("No recommendations to export")
        return None
    
    # Create DataFrame from recommendations
    df_recommendations = pd.DataFrame(recommendations)
    
    # Create summary statistics
    total_recommendations = len(df_recommendations)
    total_transfer_quantity = df_recommendations['Transfer Qty'].sum()
    
    # Transfer type statistics
    transfer_type_stats = df_recommendations['Remark'].value_counts().to_dict()
    
    # Create summary DataFrame
    summary_data = [
        ['Total Recommendations', total_recommendations, 'items'],
        ['Total Transfer Quantity', total_transfer_quantity, 'pieces'],
        ['Average Transfer Quantity', round(total_transfer_quantity / total_recommendations, 1) if total_recommendations > 0 else 0, 'pieces'],
    ]
    
    for transfer_type, count in transfer_type_stats.items():
        summary_data.append([f"{transfer_type} Count", count, 'items'])
    
    df_summary = pd.DataFrame(summary_data, columns=['Metric', 'Value', 'Unit'])
    
    # Generate Excel file
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"Transfer_Recommendations_Mode_{mode}_{timestamp}.xlsx"
    
    with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
        # Write recommendations
        df_recommendations.to_excel(writer, sheet_name='Transfer Recommendations', index=False)
        
        # Write summary
        df_summary.to_excel(writer, sheet_name='Summary Dashboard', index=False)
        
        # Get workbook and worksheets for formatting
        workbook = writer.book
        worksheet_rec = writer.sheets['Transfer Recommendations']
        worksheet_sum = writer.sheets['Summary Dashboard']
        
        # Add formats
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#D7E4BC',
            'border': 1
        })
        
        # Set column widths for recommendations
        column_widths = {
            'A': 12,  # Article
            'B': 25,  # Product Desc
            'C': 12,  # Transfer OM
            'D': 12,  # Transfer Site
            'E': 12,  # Receive OM
            'F': 12,  # Receive Site
            'G': 10,  # Transfer Qty
            'H': 18,  # Transfer Site Original Stock
            'I': 20,  # Transfer Site After Transfer Stock
            'J': 18,  # Transfer Site Safety Stock
            'K': 12,  # Transfer Site MOQ
            'L': 25,  # Remark
            'M': 18,  # Transfer Site Last Month Sold Qty
            'N': 15,  # Transfer Site MTD Sold Qty
            'O': 18,  # Receive Site Last Month Sold Qty
            'P': 15,  # Receive Site MTD Sold Qty
            'Q': 15,  # Receive Original Stock
            'R': 200  # Notes
        }
        
        for col, width in column_widths.items():
            worksheet_rec.set_column(f'{col}:{col}', width)
        
        # Apply header format to recommendations
        for col_num, value in enumerate(df_recommendations.columns.values):
            worksheet_rec.write(0, col_num, value, header_format)
        
        # Set column widths for summary
        worksheet_sum.set_column('A:A', 30)
        worksheet_sum.set_column('B:B', 15)
        worksheet_sum.set_column('C:C', 10)
        
        # Apply header format to summary
        for col_num, value in enumerate(df_summary.columns.values):
            worksheet_sum.write(0, col_num, value, header_format)
    
    print(f"Excel report generated: {filename}")
    return filename

def main():
    """Main processing function"""
    # Load and process data
    df = load_and_process_data('Windy_PIP_10Dec2025.XLSX')
    if df is None:
        return
    
    # Generate recommendations for both modes
    for mode in ['A', 'B']:
        print(f"\n{'='*50}")
        print(f"Processing Mode {mode}: {'Conservative' if mode == 'A' else 'Enhanced'} Transfer")
        print(f"{'='*50}")
        
        recommendations = generate_transfer_recommendations(df, mode)
        filename = generate_excel_report(recommendations, mode)
        
        if filename:
            print(f"Mode {mode} report completed: {filename}")
    
    print("\n" + "="*50)
    print("Data processing completed!")
    print("="*50)

if __name__ == "__main__":
    main()
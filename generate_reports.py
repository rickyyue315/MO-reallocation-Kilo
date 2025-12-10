#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Generate inventory reports, sales reports and inventory trend analysis
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime
import os

# Set matplotlib to use a non-interactive backend
plt.switch_backend('Agg')

def load_data():
    """Load the original data"""
    print("Loading original data...")
    df = pd.read_excel('Windy_PIP_10Dec2025.XLSX')
    
    # Select relevant columns
    relevant_columns = [
        'Article', 'Article Description', 'Article Long Text (60 Chars)',
        'OM', 'RP Type', 'Site', 'MOQ', 'SaSa Net Stock', 
        'Pending Received', 'Safety Stock', 'Last Month Sold Qty', 'MTD Sold Qty'
    ]
    
    # Filter to only relevant columns
    available_cols = [col for col in relevant_columns if col in df.columns]
    df = df[available_cols]
    
    # Handle missing values
    df = df.fillna(0)
    
    # Convert Article to string with leading zeros if needed
    if 'Article' in df.columns:
        df['Article'] = df['Article'].astype(str).str.zfill(12)
    
    return df

def calculate_metrics(df):
    """Calculate additional metrics for analysis"""
    print("Calculating metrics...")
    
    # Calculate total stock (net stock + pending received)
    df['Total Stock'] = df['SaSa Net Stock'] + df['Pending Received']
    
    # Calculate effective sales (last month + MTD)
    df['Effective Sold Qty'] = df['Last Month Sold Qty'] + df['MTD Sold Qty']
    
    # Calculate stock days (if we have sales data)
    # Using average daily sales from last month (assuming 30 days)
    df['Avg Daily Sales'] = df['Last Month Sold Qty'] / 30
    df['Stock Days'] = np.where(df['Avg Daily Sales'] > 0, 
                               df['Total Stock'] / df['Avg Daily Sales'], 
                               999)  # High value for no sales
    
    # Calculate safety stock days
    df['Safety Stock Days'] = np.where(df['Avg Daily Sales'] > 0, 
                                     df['Safety Stock'] / df['Avg Daily Sales'], 
                                     999)
    
    # Calculate inventory status
    df['Inventory Status'] = np.where(df['Total Stock'] < df['Safety Stock'], 'Below Safety',
                                     np.where(df['Total Stock'] > df['Safety Stock'] * 2, 'Excess', 'Normal'))
    
    # Calculate sales rate (sales vs stock ratio)
    df['Sales Rate'] = np.where(df['Total Stock'] > 0, 
                               df['Effective Sold Qty'] / df['Total Stock'], 
                               0)
    
    return df

def generate_inventory_report(df):
    """Generate inventory report"""
    print("\nGenerating inventory report...")
    
    # Create summary statistics by store
    store_summary = df.groupby(['Site', 'OM', 'RP Type']).agg({
        'Article': 'count',
        'SaSa Net Stock': 'sum',
        'Pending Received': 'sum',
        'Total Stock': 'sum',
        'Safety Stock': 'sum',
        'Last Month Sold Qty': 'sum',
        'MTD Sold Qty': 'sum',
        'Effective Sold Qty': 'sum',
        'Stock Days': 'mean',
        'Sales Rate': 'mean'
    }).reset_index()
    
    store_summary.columns = [
        'Site', 'OM', 'RP Type', 'Article Count', 'Net Stock', 'Pending Received', 
        'Total Stock', 'Safety Stock', 'Last Month Sold', 'MTD Sold', 
        'Effective Sold', 'Avg Stock Days', 'Avg Sales Rate'
    ]
    
    # Create summary by article
    article_summary = df.groupby(['Article', 'Article Description']).agg({
        'Site': 'count',
        'SaSa Net Stock': 'sum',
        'Pending Received': 'sum',
        'Total Stock': 'sum',
        'Safety Stock': 'sum',
        'Last Month Sold Qty': 'sum',
        'MTD Sold Qty': 'sum',
        'Effective Sold Qty': 'sum'
    }).reset_index()
    
    article_summary.columns = [
        'Article', 'Article Description', 'Store Count', 'Net Stock', 
        'Pending Received', 'Total Stock', 'Safety Stock', 
        'Last Month Sold', 'MTD Sold', 'Effective Sold'
    ]
    
    # Create inventory status summary
    status_summary = df.groupby(['Inventory Status']).agg({
        'Article': 'count',
        'SaSa Net Stock': 'sum',
        'Total Stock': 'sum',
        'Safety Stock': 'sum'
    }).reset_index()
    
    status_summary.columns = ['Inventory Status', 'Article Count', 'Net Stock', 'Total Stock', 'Safety Stock']
    
    # Generate Excel file
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"Inventory_Report_{timestamp}.xlsx"
    
    with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
        # Write main data
        df.to_excel(writer, sheet_name='Full Data', index=False)
        
        # Write summaries
        store_summary.to_excel(writer, sheet_name='Store Summary', index=False)
        article_summary.to_excel(writer, sheet_name='Article Summary', index=False)
        status_summary.to_excel(writer, sheet_name='Inventory Status', index=False)
        
        # Get workbook and worksheets for formatting
        workbook = writer.book
        worksheets = {name: writer.sheets[name] for name in writer.sheets}
        
        # Add formats
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#D7E4BC',
            'border': 1
        })
        
        # Apply header format to all worksheets
        for sheet_name, worksheet in worksheets.items():
            for col_num, value in enumerate(df.columns.values if sheet_name == 'Full Data' else 
                                          store_summary.columns.values if sheet_name == 'Store Summary' else
                                          article_summary.columns.values if sheet_name == 'Article Summary' else
                                          status_summary.columns.values):
                worksheet.write(0, col_num, value, header_format)
            
            # Set column widths
            worksheet.set_column(0, 0, 15)  # First column
            worksheet.set_column(1, 1, 25)  # Second column
            for col in range(2, len(df.columns) if sheet_name == 'Full Data' else len(store_summary.columns) if sheet_name == 'Store Summary' else len(article_summary.columns) if sheet_name == 'Article Summary' else len(status_summary.columns)):
                worksheet.set_column(col, col, 15)
    
    print(f"Inventory report generated: {filename}")
    return filename

def generate_sales_report(df):
    """Generate sales report"""
    print("\nGenerating sales report...")
    
    # Create sales summary by store
    store_sales = df.groupby(['Site', 'OM', 'RP Type']).agg({
        'Article': 'count',
        'Last Month Sold Qty': 'sum',
        'MTD Sold Qty': 'sum',
        'Effective Sold Qty': 'sum',
        'SaSa Net Stock': 'sum',
        'Total Stock': 'sum',
        'Sales Rate': 'mean'
    }).reset_index()
    
    store_sales.columns = [
        'Site', 'OM', 'RP Type', 'Article Count', 'Last Month Sold', 
        'MTD Sold', 'Effective Sold', 'Net Stock', 'Total Stock', 'Avg Sales Rate'
    ]
    
    # Create sales summary by article
    article_sales = df.groupby(['Article', 'Article Description']).agg({
        'Site': 'count',
        'Last Month Sold Qty': 'sum',
        'MTD Sold Qty': 'sum',
        'Effective Sold Qty': 'sum',
        'SaSa Net Stock': 'sum',
        'Total Stock': 'sum'
    }).reset_index()
    
    article_sales.columns = [
        'Article', 'Article Description', 'Store Count', 'Last Month Sold', 
        'MTD Sold', 'Effective Sold', 'Net Stock', 'Total Stock'
    ]
    
    # Calculate sales performance indicators
    article_sales['Sales Velocity'] = article_sales['Effective Sold'] / article_sales['Store Count']
    article_sales['Stock Turnover'] = np.where(article_sales['Total Stock'] > 0, 
                                             article_sales['Effective Sold'] / article_sales['Total Stock'], 
                                             0)
    
    # Sort by sales velocity
    article_sales = article_sales.sort_values('Sales Velocity', ascending=False)
    
    # Generate Excel file
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"Sales_Report_{timestamp}.xlsx"
    
    with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
        # Write summaries
        store_sales.to_excel(writer, sheet_name='Store Sales Summary', index=False)
        article_sales.to_excel(writer, sheet_name='Article Sales Summary', index=False)
        
        # Get workbook and worksheets for formatting
        workbook = writer.book
        worksheets = {name: writer.sheets[name] for name in writer.sheets}
        
        # Add formats
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#D7E4BC',
            'border': 1
        })
        
        # Apply header format to all worksheets
        for sheet_name, worksheet in worksheets.items():
            cols = store_sales.columns.values if sheet_name == 'Store Sales Summary' else article_sales.columns.values
            for col_num, value in enumerate(cols):
                worksheet.write(0, col_num, value, header_format)
            
            # Set column widths
            worksheet.set_column(0, 0, 15)  # First column
            worksheet.set_column(1, 1, 25)  # Second column
            for col in range(2, len(cols)):
                worksheet.set_column(col, col, 15)
    
    print(f"Sales report generated: {filename}")
    return filename

def generate_trend_analysis(df):
    """Generate inventory trend analysis"""
    print("\nGenerating inventory trend analysis...")
    
    # Create trend analysis by store type
    trend_data = df.groupby(['RP Type']).agg({
        'Article': 'count',
        'SaSa Net Stock': ['sum', 'mean'],
        'Pending Received': ['sum', 'mean'],
        'Total Stock': ['sum', 'mean'],
        'Safety Stock': ['sum', 'mean'],
        'Last Month Sold Qty': ['sum', 'mean'],
        'MTD Sold Qty': ['sum', 'mean'],
        'Effective Sold Qty': ['sum', 'mean'],
        'Stock Days': 'mean',
        'Sales Rate': 'mean'
    }).reset_index()
    
    # Flatten column names
    trend_data.columns = ['RP Type', 'Article Count', 'Total Net Stock', 'Avg Net Stock',
                         'Total Pending', 'Avg Pending', 'Total Stock', 'Avg Stock',
                         'Total Safety Stock', 'Avg Safety Stock', 'Total Last Month Sold',
                         'Avg Last Month Sold', 'Total MTD Sold', 'Avg MTD Sold',
                         'Total Effective Sold', 'Avg Effective Sold', 'Avg Stock Days', 'Avg Sales Rate']
    
    # Calculate inventory efficiency metrics
    trend_data['Inventory Efficiency'] = np.where(trend_data['Avg Stock'] > 0,
                                                trend_data['Avg Effective Sold'] / trend_data['Avg Stock'],
                                                0)
    
    trend_data['Safety Stock Coverage'] = np.where(trend_data['Avg Safety Stock'] > 0,
                                                 trend_data['Avg Stock'] / trend_data['Avg Safety Stock'],
                                                 0)
    
    # Create charts
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    chart_filename = f"Inventory_Trend_Charts_{timestamp}.png"
    
    # Set up the figure
    plt.figure(figsize=(15, 10))
    
    # Chart 1: Stock vs Sales by RP Type
    plt.subplot(2, 2, 1)
    sns.barplot(x='RP Type', y='Avg Stock', data=trend_data, color='blue', alpha=0.7, label='Avg Stock')
    sns.barplot(x='RP Type', y='Avg Effective Sold', data=trend_data, color='red', alpha=0.7, label='Avg Effective Sold')
    plt.title('Average Stock vs Effective Sales by RP Type')
    plt.ylabel('Quantity')
    plt.legend()
    
    # Chart 2: Inventory Efficiency by RP Type
    plt.subplot(2, 2, 2)
    sns.barplot(x='RP Type', y='Inventory Efficiency', data=trend_data, color='green')
    plt.title('Inventory Efficiency by RP Type')
    plt.ylabel('Efficiency Ratio')
    
    # Chart 3: Safety Stock Coverage
    plt.subplot(2, 2, 3)
    sns.barplot(x='RP Type', y='Safety Stock Coverage', data=trend_data, color='orange')
    plt.title('Safety Stock Coverage by RP Type')
    plt.ylabel('Coverage Ratio')
    
    # Chart 4: Stock Days
    plt.subplot(2, 2, 4)
    sns.barplot(x='RP Type', y='Avg Stock Days', data=trend_data, color='purple')
    plt.title('Average Stock Days by RP Type')
    plt.ylabel('Days')
    
    plt.tight_layout()
    plt.savefig(chart_filename, dpi=300, bbox_inches='tight')
    plt.close()
    
    # Generate Excel file with trend analysis
    excel_filename = f"Inventory_Trend_Analysis_{timestamp}.xlsx"
    
    with pd.ExcelWriter(excel_filename, engine='xlsxwriter') as writer:
        # Write trend data
        trend_data.to_excel(writer, sheet_name='Trend Analysis', index=False)
        
        # Get workbook and worksheet for formatting
        workbook = writer.book
        worksheet = writer.sheets['Trend Analysis']
        
        # Add formats
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#D7E4BC',
            'border': 1
        })
        
        # Apply header format
        for col_num, value in enumerate(trend_data.columns.values):
            worksheet.write(0, col_num, value, header_format)
        
        # Set column widths
        worksheet.set_column(0, 0, 15)  # RP Type
        for col in range(1, len(trend_data.columns)):
            worksheet.set_column(col, col, 18)
        
        # Insert chart
        worksheet.insert_image('H2', chart_filename)
    
    print(f"Trend analysis generated: {excel_filename}")
    print(f"Charts saved as: {chart_filename}")
    
    return excel_filename, chart_filename

def main():
    """Main function"""
    print("="*50)
    print("Inventory Reports and Trend Analysis")
    print("="*50)
    
    # Load data
    df = load_data()
    
    # Calculate metrics
    df = calculate_metrics(df)
    
    # Generate reports
    inventory_file = generate_inventory_report(df)
    sales_file = generate_sales_report(df)
    trend_file, chart_file = generate_trend_analysis(df)
    
    print("\n" + "="*50)
    print("All reports generated successfully!")
    print("="*50)
    print(f"Inventory Report: {inventory_file}")
    print(f"Sales Report: {sales_file}")
    print(f"Trend Analysis: {trend_file}")
    print(f"Trend Charts: {chart_file}")

if __name__ == "__main__":
    main()
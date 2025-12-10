#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Check data file structure
"""

import pandas as pd
import sys

def main():
    try:
        # Read Excel file
        df = pd.read_excel('Windy_PIP_10Dec2025.XLSX')
        
        print("File check successful!")
        print(f"Number of rows: {len(df)}")
        print(f"Number of columns: {len(df.columns)}")
        print("\nColumn names:")
        for i, col in enumerate(df.columns):
            print(f"{i+1}. {col}")
        
        print("\nFirst 5 rows:")
        print(df.head())
        
        print("\nData types:")
        print(df.dtypes)
        
        return df
    except Exception as e:
        print(f"Error: {str(e)}")
        return None

if __name__ == "__main__":
    df = main()
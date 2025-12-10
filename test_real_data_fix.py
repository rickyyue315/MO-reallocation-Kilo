import pandas as pd
import os
from business_logic import BusinessLogic
from data_processor import DataProcessor

def test_specific_case():
    """
    Test specific case: 10029151001 HA30 transfer quantity vs SaSa Net Stock
    """
    file_path = 'Windy_PIP_10Dec2025.XLSX'
    if not os.path.exists(file_path):
        print(f"File not found: {file_path}")
        return

    print(f"Reading file: {file_path}")
    
    # Use DataProcessor to process data
    processor = DataProcessor()
    try:
        # Directly read using pandas and do necessary preprocessing
        df = pd.read_excel(file_path)
        
        # Ensure necessary columns exist and convert types
        numeric_cols = ['SaSa Net Stock', 'Pending Received', 'Safety Stock', 'Last Month Sold Qty', 'MTD Sold Qty', 'MOQ']
        for col in numeric_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        
        # Calculate effective sold qty
        if 'Effective Sold Qty' not in df.columns:
            df['Effective Sold Qty'] = df['Last Month Sold Qty'] + df['MTD Sold Qty']
            
        # Ensure Article is string
        df['Article'] = df['Article'].astype(str)
        
    except Exception as e:
        print(f"Data processing failed: {e}")
        return

    # Filter specific article and site
    target_article = '100291510001' # Note: Excel might have 100291510001 (12 digits)
    
    # Check Article format in data
    sample_article = df['Article'].iloc[0]
    print(f"Sample Article in data: {sample_article} (Length: {len(str(sample_article))})")
    
    # Try to find target article
    possible_articles = [a for a in df['Article'].unique() if '10029151001' in str(a) or '100291510001' in str(a)]
    print(f"Found related Articles: {possible_articles}")
    
    if not possible_articles:
        print("Target article not found")
        return
        
    target_article = possible_articles[0]
    print(f"Using target Article: {target_article}")

    # Check HA30 stock status
    ha30_row = df[(df['Article'] == target_article) & (df['Site'] == 'HA30')]
    if len(ha30_row) == 0:
        print(f"No record found for {target_article} at HA30")
    else:
        print(f"HA30 Original Data:\n{ha30_row[['Article', 'Site', 'SaSa Net Stock', 'Pending Received', 'Safety Stock', 'Effective Sold Qty']].to_string()}")

    # Run business logic to generate recommendations
    business_logic = BusinessLogic()
    
    # Test Mode A
    print("\n--- Testing Mode A ---")
    success, recommendations_a, _ = business_logic.generate_transfer_recommendations(df, "A")
    
    if success:
        # Check all recommendations
        print(f"Mode A generated {len(recommendations_a)} recommendations")
        errors = []
        for rec in recommendations_a:
            if rec['Transfer Qty'] > rec['Transfer Site Original Stock']:
                errors.append(rec)
        
        if errors:
            print(f"FAIL: Found {len(errors)} recommendations with Transfer Qty > Original Stock in Mode A")
            for rec in errors[:5]:
                print(f"  Article: {rec['Article']}, Site: {rec['Transfer Site']}, Qty: {rec['Transfer Qty']}, Stock: {rec['Transfer Site Original Stock']}")
        else:
            print("PASS: All Mode A recommendations have Transfer Qty <= Original Stock")
            
        # Detailed check for any stock issues
        stock_issues = []
        for rec in recommendations_a:
            if rec['Transfer Qty'] > rec['Transfer Site Original Stock']:
                stock_issues.append(rec)
        
        if stock_issues:
            print(f"\nDETAILED STOCK ISSUES FOUND ({len(stock_issues)}):")
            for rec in stock_issues:
                print(f"  VIOLATION: Article={rec['Article']}, Transfer Site={rec['Transfer Site']}, Receive Site={rec['Receive Site']}, Transfer Qty={rec['Transfer Qty']}, Original Stock={rec['Transfer Site Original Stock']}")
            
        # Check HA30 recommendations
        ha30_recs = [r for r in recommendations_a if r['Article'] == target_article and r['Transfer Site'] == 'HA30']
        if ha30_recs:
            for rec in ha30_recs:
                print(f"Mode A Recommendation for HA30: Transfer {rec['Transfer Qty']}, Original Stock {rec['Transfer Site Original Stock']}")
    else:
        print(f"Mode A generation failed: {recommendations_a}")

    # Test Mode B
    print("\n--- Testing Mode B ---")
    success, recommendations_b, _ = business_logic.generate_transfer_recommendations(df, "B")
    
    if success:
        # Check all recommendations
        print(f"Mode B generated {len(recommendations_b)} recommendations")
        errors = []
        for rec in recommendations_b:
            if rec['Transfer Qty'] > rec['Transfer Site Original Stock']:
                errors.append(rec)
        
        if errors:
            print(f"FAIL: Found {len(errors)} recommendations with Transfer Qty > Original Stock in Mode B")
            for rec in errors[:5]:
                print(f"  Article: {rec['Article']}, Site: {rec['Transfer Site']}, Qty: {rec['Transfer Qty']}, Stock: {rec['Transfer Site Original Stock']}")
        else:
            print("PASS: All Mode B recommendations have Transfer Qty <= Original Stock")
            
        # Detailed check for any stock issues
        stock_issues = []
        for rec in recommendations_b:
            if rec['Transfer Qty'] > rec['Transfer Site Original Stock']:
                stock_issues.append(rec)
        
        if stock_issues:
            print(f"\nDETAILED STOCK ISSUES FOUND ({len(stock_issues)}):")
            for rec in stock_issues:
                print(f"  VIOLATION: Article={rec['Article']}, Transfer Site={rec['Transfer Site']}, Receive Site={rec['Receive Site']}, Transfer Qty={rec['Transfer Qty']}, Original Stock={rec['Transfer Site Original Stock']}")

        # Check HA30 recommendations
        ha30_recs = [r for r in recommendations_b if r['Article'] == target_article and r['Transfer Site'] == 'HA30']
        if ha30_recs:
            for rec in ha30_recs:
                print(f"Mode B Recommendation for HA30: Transfer {rec['Transfer Qty']}, Original Stock {rec['Transfer Site Original Stock']}")
    else:
        print(f"Mode B generation failed: {recommendations_b}")

    # Check for any violations against demand quantity
    print("\n--- Checking Demand Quantity Constraint ---")
    
    def check_demand_constraint(recommendations, mode_name):
        # Re-identify receive candidates to get original demand quantities
        rf_candidates = df[df['RP Type'] == 'RF'].copy()
        
        # Calculate max sold per article for potential shortage identification
        rf_candidates['max_sold_per_article'] = rf_candidates.groupby('Article')['Effective Sold Qty'].transform('max')
        
        # Identify urgent shortage candidates
        urgent_candidates = rf_candidates[
            (rf_candidates['SaSa Net Stock'] + rf_candidates['Pending Received'] == 0) &
            (rf_candidates['Effective Sold Qty'] > 0)
        ].copy()
        urgent_candidates['demand_quantity'] = urgent_candidates['Safety Stock']
        urgent_candidates['receive_priority'] = '緊急缺貨'
        
        # Identify potential shortage candidates
        potential_candidates = rf_candidates[
            (rf_candidates['SaSa Net Stock'] + rf_candidates['Pending Received'] < rf_candidates['Safety Stock']) &
            (rf_candidates['Effective Sold Qty'] == rf_candidates['max_sold_per_article'])
        ].copy()
        total_stock = potential_candidates['SaSa Net Stock'] + potential_candidates['Pending Received']
        potential_candidates['demand_quantity'] = potential_candidates['Safety Stock'] - total_stock
        potential_candidates['receive_priority'] = '潛在缺貨'
        
        # Combine all receive candidates
        receive_candidates = pd.concat([urgent_candidates, potential_candidates], ignore_index=True)
        
        violations = []
        for rec in recommendations:
            # Find the corresponding receive candidate
            receive_candidate = receive_candidates[
                (receive_candidates['Article'] == rec['Article']) &
                (receive_candidates['Site'] == rec['Receive Site'])
            ]
            
            if len(receive_candidate) > 0:
                original_demand = receive_candidate['demand_quantity'].iloc[0]
                if rec['Transfer Qty'] > original_demand:
                    violations.append({
                        'rec': rec,
                        'original_demand': original_demand
                    })
        
        if violations:
            print(f"FAIL: Found {len(violations)} {mode_name} recommendations with Transfer Qty > Original Demand")
            for v in violations[:5]:  # Print first 5 violations
                r = v['rec']
                print(f"  Article: {r['Article']}, Transfer: {r['Transfer Qty']}, Demand: {v['original_demand']}, From: {r['Transfer Site']}, To: {r['Receive Site']}")
        else:
            print(f"PASS: All {mode_name} recommendations have Transfer Qty <= Original Demand")
    
    check_demand_constraint(recommendations_a, "Mode A")
    check_demand_constraint(recommendations_b, "Mode B")

if __name__ == "__main__":
    test_specific_case()
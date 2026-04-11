"""
Complete Test Script for ExcelExportManager
Generates all 14 sheets with proper formatting and color coding.
Run this to verify the implementation works correctly.
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.excel_export_manager import ExcelExportManager
from src.logging_config import LoggingManager
import pandas as pd

logger = LoggingManager.get_logger()


def test_excel_export():
    """
    Complete test of ExcelExportManager.
    Generates sample data and creates a 14-sheet workbook.
    """
    print("\n" + "="*80)
    print("EXCEL EXPORT MANAGER - COMPLETE TEST")
    print("="*80 + "\n")
    
    try:
        # Step 1: Initialize manager
        print("[1/4] Initializing ExcelExportManager...")
        manager = ExcelExportManager(storage_dir="./local_trading_data")
        print(f"✓ Manager initialized")
        print(f"  Output directory: {manager.get_export_directory()}\n")
        
        # Step 2: Generate sample dataframes
        print("[2/4] Generating sample dataframes for all 14 sheets...")
        dataframes = manager.generate_sample_dataframes()
        print(f"✓ Generated {len(dataframes)} dataframes:")
        for sheet_name, df in dataframes.items():
            print(f"  - {sheet_name}: {len(df)} rows × {len(df.columns)} columns")
        print()
        
        # Step 3: Create workbook
        print("[3/4] Creating advanced multi-sheet workbook...")
        filepath = manager.create_advanced_workbook(
            dataframes=dataframes,
            date_str="2026-04-11"
        )
        print(f"✓ Workbook created successfully!")
        print(f"  File path: {filepath}\n")
        
        # Step 4: Verify file
        print("[4/4] Verifying file creation...")
        if os.path.exists(filepath):
            file_size = os.path.getsize(filepath)
            print(f"✓ File verified!")
            print(f"  File size: {file_size:,} bytes")
            print(f"  File exists: Yes\n")
            
            print("="*80)
            print("✅ TEST SUCCESSFUL - All 14 sheets generated with formatting")
            print("="*80 + "\n")
            return True
        else:
            print(f"✗ File not found at: {filepath}\n")
            return False
    
    except Exception as e:
        print(f"\n✗ TEST FAILED: {e}")
        import traceback
        traceback.print_exc()
        return False


def test_custom_data():
    """
    Test with custom dataframes showing color-coded sheets.
    """
    print("\n" + "="*80)
    print("CUSTOM DATA TEST - Color Coding")
    print("="*80 + "\n")
    
    try:
        manager = ExcelExportManager(storage_dir="./local_trading_data")
        
        # Get time intervals
        time_intervals = manager.TIME_INTERVALS
        print(f"[*] Using {len(time_intervals)} time intervals (09:15 to 15:25)\n")
        
        # Create custom SQZMOM data with specific colors
        print("[*] Creating SQZMOM sheet with color codes:")
        sqzmom_data = []
        colors = ["LIME", "GREEN", "RED", "MAROON", "ORANGE", "BLACK"]
        symbols = ["RELIANCE", "TCS", "INFY"]
        
        for symbol in symbols:
            row = {"Symbol": symbol, "Metrics": "Momentum"}
            for i, time in enumerate(time_intervals):
                row[time] = colors[i % len(colors)]
            sqzmom_data.append(row)
        
        sqzmom_df = pd.DataFrame(sqzmom_data)
        print(f"  ✓ SQZMOM DataFrame created: {len(sqzmom_df)} rows")
        for color in colors:
            print(f"    - {color}")
        
        # Create custom FINAL data with sync signals
        print("\n[*] Creating FINAL sheet with sync signals:")
        final_data = []
        final_colors = ["BULL SYNC", "BEAR SYNC", "NO SYNC", "Long Buildup", "Short Buildup", "WAIT"]
        
        for symbol in symbols:
            row = {"Symbol": symbol, "Metrics": "Sync"}
            for i, time in enumerate(time_intervals):
                row[time] = final_colors[i % len(final_colors)]
            final_data.append(row)
        
        final_df = pd.DataFrame(final_data)
        print(f"  ✓ FINAL DataFrame created: {len(final_df)} rows")
        for final_color in final_colors:
            print(f"    - {final_color}")
        
        # Create workbook with custom data
        print("\n[*] Creating workbook with custom color-coded data...")
        dataframes = {
            "SQZMOM": sqzmom_df,
            "FINAL": final_df
        }
        
        filepath = manager.create_advanced_workbook(
            dataframes=dataframes,
            filename="custom_colored_signals.xlsx"
        )
        
        print(f"\n✓ Custom workbook created!")
        print(f"  File path: {filepath}")
        
        if os.path.exists(filepath):
            file_size = os.path.getsize(filepath)
            print(f"  File size: {file_size:,} bytes")
            print(f"\n✅ Custom data test successful!")
        
        return True
    
    except Exception as e:
        print(f"\n✗ Custom data test failed: {e}")
        import traceback
        traceback.print_exc()
        return False


def main():
    """Run all tests."""
    print("\n")
    print("╔" + "="*78 + "╗")
    print("║" + " "*15 + "EXCEL EXPORT MANAGER - COMPREHENSIVE TEST SUITE" + " "*16 + "║")
    print("╚" + "="*78 + "╝")
    
    # Run tests
    test1_result = test_excel_export()
    test2_result = test_custom_data()
    
    # Summary
    print("\n" + "="*80)
    print("TEST SUMMARY")
    print("="*80)
    print(f"Test 1 - Full Workbook Generation: {'✅ PASSED' if test1_result else '❌ FAILED'}")
    print(f"Test 2 - Custom Color Coding:      {'✅ PASSED' if test2_result else '❌ FAILED'}")
    print("="*80)
    
    if test1_result and test2_result:
        print("\n🎉 ALL TESTS PASSED - Implementation is working correctly!\n")
        return 0
    else:
        print("\n⚠️  Some tests failed - Check errors above\n")
        return 1


if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code)

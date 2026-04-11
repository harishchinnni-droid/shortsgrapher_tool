"""
Excel Export Manager - Usage Examples and Integration Guide

This module demonstrates how to use the advanced ExcelExportManager class
to generate sophisticated multi-sheet workbooks with time-series data.
"""

from src.excel_export_manager import ExcelExportManager
import pandas as pd
from datetime import datetime


class ExcelExportExamples:
    """
    Comprehensive examples of using ExcelExportManager for various use cases.
    """
    
    @staticmethod
    def example_generate_simple_workbook():
        """
        Example 1: Generate workbook with sample/demo data.
        This is useful for testing and understanding the structure.
        """
        # Initialize manager
        manager = ExcelExportManager(storage_dir="./local_trading_data")
        
        # Generate sample dataframes (includes all 13 sheets)
        dataframes = manager.generate_sample_dataframes()
        
        # Create workbook
        filepath = manager.create_advanced_workbook(
            dataframes=dataframes,
            date_str=datetime.now().strftime("%Y-%m-%d")
        )
        
        print(f"✓ Sample workbook created at: {filepath}")
        return filepath
    
    @staticmethod
    def example_create_with_custom_data():
        """
        Example 2: Create workbook with your own custom data.
        This shows how to structure dataframes for the ExcelExportManager.
        """
        manager = ExcelExportManager(storage_dir="./local_trading_data")
        
        # Define time intervals (provided by manager)
        time_intervals = manager.TIME_INTERVALS  # 76 intervals from 09:15 to 15:25
        
        # Create a Reference sheet
        reference_df = pd.DataFrame({
            "Symbol": ["RELIANCE", "TCS", "INFY"],
            "Sector": ["Energy", "IT", "IT"],
            "Market_Cap": ["7.5L Cr", "13.8L Cr", "7.2L Cr"]
        })
        
        # Create a SQZMOM sheet with color codes
        sqzmom_data = []
        for symbol in ["RELIANCE", "TCS"]:
            row = {"Symbol": symbol, "Metrics": "Momentum"}
            # Populate with color values: LIME, GREEN, RED, MAROON, ORANGE, BLACK
            for time in time_intervals:
                import random
                row[time] = random.choice(["LIME", "GREEN", "RED", "MAROON", "ORANGE", "BLACK"])
            sqzmom_data.append(row)
        sqzmom_df = pd.DataFrame(sqzmom_data)
        
        # Create a FINAL sheet with sync recommendations
        final_data = []
        for symbol in ["RELIANCE", "TCS"]:
            row = {"Symbol": symbol, "Metrics": "Sync"}
            # Populate with FINAL status values
            for time in time_intervals:
                import random
                row[time] = random.choice([
                    "BULL SYNC", "BEAR SYNC", "NO SYNC",
                    "Long Buildup", "Short Buildup", "WAIT"
                ])
            final_data.append(row)
        final_df = pd.DataFrame(final_data)
        
        # Prepare all dataframes
        dataframes = {
            "Reference": reference_df,
            "SQZMOM": sqzmom_df,
            "FINAL": final_df,
            # Other sheets can be empty or populated similarly
        }
        
        # Create workbook
        filepath = manager.create_advanced_workbook(
            dataframes=dataframes,
            filename="custom_data.xlsx"
        )
        
        print(f"✓ Custom workbook created at: {filepath}")
        return filepath
    
    @staticmethod
    def example_create_colored_sheet():
        """
        Example 3: Deep dive into creating colored sheets (SQZMOM & FINAL).
        Demonstrates the color mapping system.
        """
        manager = ExcelExportManager(storage_dir="./local_trading_data")
        
        # Color mapping for SQZMOM sheet
        sqzmom_colors = manager.COLOR_SCHEMES["SQZMOM"]
        print("SQZMOM Available Colors:")
        for color_name, color_info in sqzmom_colors.items():
            print(f"  - {color_name}: BG={color_info['bg']}, FONT={color_info['font']}")
            print(f"    Description: {color_info['desc']}")
        
        # Color mapping for FINAL sheet
        final_colors = manager.COLOR_SCHEMES["FINAL"]
        print("\nFINAL Available Colors:")
        for color_name, color_info in final_colors.items():
            print(f"  - {color_name}: BG={color_info['bg']}, FONT={color_info['font']}")
            print(f"    Description: {color_info['desc']}")
        
        # Example: Create SQZMOM data with proper color values
        symbols = ["RELIANCE", "TCS", "INFY"]
        time_intervals = manager.TIME_INTERVALS
        
        sqzmom_rows = []
        for symbol in symbols:
            # Create multiple metrics for the same symbol
            metrics = ["Signal", "Indication", "Status"]
            for metric in metrics:
                row = {"Symbol": symbol, "Metrics": metric}
                # Populate with color values from SQZMOM color scheme
                color_values = list(sqzmom_colors.keys())  # ["LIME", "GREEN", "RED", ...]
                for i, time_interval in enumerate(time_intervals):
                    row[time_interval] = color_values[i % len(color_values)]
                sqzmom_rows.append(row)
        
        sqzmom_df = pd.DataFrame(sqzmom_rows)
        
        # Create workbook with just SQZMOM sheet for testing
        dataframes = {"SQZMOM": sqzmom_df}
        
        filepath = manager.create_advanced_workbook(
            dataframes=dataframes,
            filename="colored_sqzmom.xlsx"
        )
        
        print(f"✓ Colored SQZMOM workbook created at: {filepath}")
        return filepath
    
    @staticmethod
    def example_all_sheets_structure():
        """
        Example 4: Complete structure with all 14 sheets.
        Shows the expected structure for each sheet.
        """
        manager = ExcelExportManager(storage_dir="./local_trading_data")
        
        sheets_structure = {
            "Legend": [
                "Column A: Indicator/Sheet Category",
                "Column B: Status/Value",
                "Column C: Visual Color (Cell formatted with actual color)",
                "Column D: Description",
                "Auto-generated from COLOR_SCHEMES"
            ],
            "Reference": [
                "Column A: Symbol",
                "Column B onwards: Metadata (Sector, Market Cap, etc.)",
                "Can have any structure - no time-series required"
            ],
            "R1S1": [
                "Column A: Symbol",
                "Column B: Metrics (e.g., Open, High, Low, Close)",
                "Columns C-onwards: Time intervals (09:15, 09:20, ...)",
                "Time-series numeric data"
            ],
            "Pine": [
                "Column A: Symbol",
                "Column B: Metrics",
                "Columns C-onwards: Time intervals",
                "EMA tracking data"
            ],
            "SQZMOM": [
                "Column A: Symbol",
                "Column B: Metrics",
                "Columns C-onwards: Time intervals",
                "Color-coded values: LIME, GREEN, RED, MAROON, ORANGE, BLACK"
            ],
            "Breakout": [
                "Column A: Symbol",
                "Column B: Metrics",
                "Columns C-onwards: Time intervals",
                "Breakout signals and levels"
            ],
            "RSI": [
                "Column A: Symbol",
                "Column B: Metrics",
                "Columns C-onwards: Time intervals",
                "RSI values (0-100)"
            ],
            "ADX": [
                "Column A: Symbol",
                "Column B: Metrics",
                "Columns C-onwards: Time intervals",
                "ADX strength values"
            ],
            "VWAP": [
                "Column A: Symbol",
                "Column B: Metrics",
                "Columns C-onwards: Time intervals",
                "VWAP price levels"
            ],
            "MTF_Trend": [
                "Column A: Symbol",
                "Column B: Metrics",
                "Columns C-onwards: Time intervals",
                "Multi-timeframe trend analysis"
            ],
            "OI_Dynamics": [
                "Column A: Symbol",
                "Column B: Metrics",
                "Columns C-onwards: Time intervals",
                "Open Interest dynamics"
            ],
            "15m_5m_Sync": [
                "Column A: Symbol",
                "Column B: Metrics",
                "Columns C-onwards: Time intervals",
                "Synchronization between 15min and 5min"
            ],
            "FINAL": [
                "Column A: Symbol",
                "Column B: Metrics",
                "Columns C-onwards: Time intervals",
                "Color-coded final signals: BULL SYNC, BEAR SYNC, NO SYNC, Long Buildup, Short Buildup, WAIT"
            ],
            "Order": [
                "Column A: Symbol",
                "Columns B onwards: Trade execution details",
                "Can have free-form structure",
                "Example: Order_Type, Entry_Time, Entry_Price, Exit_Time, Exit_Price, PnL"
            ]
        }
        
        print("✓ Excel Workbook Structure (14 sheets):")
        print("=" * 70)
        
        for idx, (sheet_name, structure) in enumerate(sheets_structure.items(), 1):
            print(f"\n{idx}. {sheet_name}:")
            for detail in structure:
                print(f"   - {detail}")
    
    @staticmethod
    def example_export_from_main_tool():
        """
        Example 5: Integration with main StockAutomationTool.
        Shows how to use the export_advanced_timeseries_workbook method.
        """
        from src.main import StockAutomationTool
        
        # Initialize tool
        tool = StockAutomationTool(local_storage_dir="./local_trading_data")
        
        # Export advanced workbook
        filepath = tool.export_advanced_timeseries_workbook(
            date_str="2026-04-11"
        )
        
        print(f"✓ Advanced workbook exported from main tool: {filepath}")
        return filepath
    
    @staticmethod
    def example_batch_export():
        """
        Example 6: Batch export multiple workbooks with different datasets.
        Useful for exporting signals for multiple dates.
        """
        manager = ExcelExportManager(storage_dir="./local_trading_data")
        
        dates = ["2026-04-10", "2026-04-11", "2026-04-12"]
        
        exported_files = []
        for date in dates:
            # Generate or fetch your data for this date
            dataframes = manager.generate_sample_dataframes()
            
            filepath = manager.create_advanced_workbook(
                dataframes=dataframes,
                date_str=date
            )
            exported_files.append(filepath)
            print(f"✓ Exported workbook for {date}: {filepath}")
        
        return exported_files


# Example execution
if __name__ == "__main__":
    print("Excel Export Manager - Usage Examples\n")
    print("=" * 70)
    
    # Run examples
    print("\n### Example 1: Generate Simple Workbook ###")
    example1_path = ExcelExportExamples.example_generate_simple_workbook()
    
    print("\n### Example 3: Create Colored Sheet ###")
    example3_path = ExcelExportExamples.example_create_colored_sheet()
    
    print("\n### Example 4: All Sheets Structure ###")
    ExcelExportExamples.example_all_sheets_structure()
    
    print("\n" + "=" * 70)
    print("Examples completed! Check local_trading_data/excel_reports/ for output files.")

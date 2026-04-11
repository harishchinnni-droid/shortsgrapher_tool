"""
Advanced Excel Export Manager
Generates complex multi-sheet workbooks with time-series data and sophisticated formatting.
Uses xlsxwriter for advanced conditional formatting and cell validation.
Strictly isolated sheet creation for multi-indicator time-series analysis.
"""

import os
from datetime import datetime
from typing import Dict, List, Optional, Any
import pandas as pd
import xlsxwriter
from src.logging_config import LoggingManager

logger = LoggingManager.get_logger()


class ExcelExportManager:
    """
    Advanced Excel export manager for multi-sheet time-series workbooks.
    Handles complex formatting, color coding, and legend generation.
    """
    
    # Color Scheme Definitions
    COLOR_SCHEMES = {
        "SQZMOM": {
            "LIME": {"bg": "#00FF00", "font": "#000000", "desc": "Strong Bullish Momentum"},
            "GREEN": {"bg": "#008000", "font": "#FFFFFF", "desc": "Weakening Bullish Momentum"},
            "RED": {"bg": "#FF0000", "font": "#FFFFFF", "desc": "Strong Bearish Momentum"},
            "MAROON": {"bg": "#800000", "font": "#FFFFFF", "desc": "Weakening Bearish Momentum"},
            "ORANGE": {"bg": "#FFA500", "font": "#000000", "desc": "Squeeze On"},
            "BLACK": {"bg": "#000000", "font": "#FFFFFF", "desc": "Squeeze Off"}
        },
        "FINAL": {
            "BULL SYNC": {"bg": "#00FF00", "font": "#000000", "desc": "Bullish Synchronization"},
            "BEAR SYNC": {"bg": "#FF0000", "font": "#FFFFFF", "desc": "Bearish Synchronization"},
            "NO SYNC": {"bg": "#D3D3D3", "font": "#000000", "desc": "No Synchronization"},
            "Long Buildup": {"bg": "#C6EFCE", "font": "#006100", "desc": "Long Position Buildup"},
            "Short Buildup": {"bg": "#FFC7CE", "font": "#9C0006", "desc": "Short Position Buildup"},
            "WAIT": {"bg": "#FFEB9C", "font": "#9C6500", "desc": "Wait Signal"}
        }
    }
    
    # Sheet names in order (Legend will be inserted at position 0)
    SHEET_NAMES = [
        "Reference",
        "R1S1",
        "Pine",
        "SQZMOM",
        "Breakout",
        "RSI",
        "ADX",
        "VWAP",
        "MTF_Trend",
        "OI_Dynamics",
        "15m_5m_Sync",
        "FINAL",
        "Order"
    ]
    
    # Time intervals for trading (05:15 to 15:25 in 5-minute intervals)
    TIME_INTERVALS = [
        "09:15", "09:20", "09:25", "09:30", "09:35", "09:40", "09:45", "09:50", "09:55",
        "10:00", "10:05", "10:10", "10:15", "10:20", "10:25", "10:30", "10:35", "10:40",
        "10:45", "10:50", "10:55", "11:00", "11:05", "11:10", "11:15", "11:20", "11:25",
        "11:30", "11:35", "11:40", "11:45", "11:50", "11:55", "12:00", "12:05", "12:10",
        "12:15", "12:20", "12:25", "12:30", "12:35", "12:40", "12:45", "12:50", "12:55",
        "13:00", "13:05", "13:10", "13:15", "13:20", "13:25", "13:30", "13:35", "13:40",
        "13:45", "13:50", "13:55", "14:00", "14:05", "14:10", "14:15", "14:20", "14:25",
        "14:30", "14:35", "14:40", "14:45", "14:50", "14:55", "15:00", "15:05", "15:10",
        "15:15", "15:20", "15:25"
    ]
    
    def __init__(self, storage_dir: str = "./local_trading_data"):
        """
        Initialize Excel Export Manager.
        
        Args:
            storage_dir: Directory to store Excel workbooks
        """
        self.storage_dir = storage_dir
        self.excel_dir = os.path.join(storage_dir, "excel_reports")
        os.makedirs(self.excel_dir, exist_ok=True)
        logger.info(f"Excel Export Manager initialized at: {self.excel_dir}")
    
    def create_advanced_workbook(
        self,
        dataframes: Dict[str, pd.DataFrame],
        filename: Optional[str] = None,
        date_str: Optional[str] = None
    ) -> str:
        """
        Create a complex multi-sheet workbook with time-series data and formatting.
        
        Args:
            dataframes: Dictionary mapping sheet names to pandas DataFrames
                       Expected keys: Reference, R1S1, Pine, SQZMOM, Breakout, RSI, ADX, VWAP, MTF_Trend, OI_Dynamics, 15m_5m_Sync, FINAL, Order
            filename: Custom filename (optional)
            date_str: Date string for filename (optional)
            
        Returns:
            str: Path to created Excel file
        """
        try:
            # Generate filename
            if filename is None:
                if date_str is None:
                    date_str = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
                filename = f"{date_str}_Advanced_Signals.xlsx"
            
            filepath = os.path.join(self.excel_dir, filename)
            
            # Create workbook
            workbook = xlsxwriter.Workbook(filepath)
            
            # Create format objects for reuse
            formats = self._create_formats(workbook)
            
            # Create Legend sheet first (index 0)
            self._create_legend_sheet(workbook, formats)
            
            # Create all other sheets
            for sheet_name in self.SHEET_NAMES:
                if sheet_name in dataframes:
                    df = dataframes[sheet_name]
                    
                    if sheet_name in ["SQZMOM", "FINAL"]:
                        # Colored time-series sheets
                        self._create_timeseries_sheet_with_colors(
                            workbook, sheet_name, df, formats
                        )
                    else:
                        # Regular time-series sheets
                        self._create_timeseries_sheet(workbook, sheet_name, df, formats)
                else:
                    # Create empty sheet if data not provided
                    self._create_empty_sheet(workbook, sheet_name, formats)
            
            workbook.close()
            logger.info(f"✓ Advanced Excel workbook created: {filepath}")
            return filepath
        
        except Exception as e:
            logger.error(f"Failed to create advanced workbook: {e}")
            raise
    
    def _create_formats(self, workbook: xlsxwriter.Workbook) -> Dict[str, Any]:
        """
        Create reusable format objects for the workbook.
        
        Args:
            workbook: xlsxwriter Workbook object
            
        Returns:
            Dictionary of format objects
        """
        formats = {
            # Header formats
            "header": workbook.add_format({
                "bold": True,
                "bg_color": "#366092",
                "font_color": "#FFFFFF",
                "align": "center",
                "valign": "vcenter",
                "border": 1,
                "text_wrap": True,
                "font_size": 10
            }),
            "header_frozen": workbook.add_format({
                "bold": True,
                "bg_color": "#366092",
                "font_color": "#FFFFFF",
                "align": "center",
                "valign": "vcenter",
                "border": 1,
                "text_wrap": True,
                "font_size": 10,
                "locked": True
            }),
            # Data formats
            "data": workbook.add_format({
                "align": "center",
                "valign": "vcenter",
                "border": 1,
                "font_size": 9
            }),
            "data_left": workbook.add_format({
                "align": "left",
                "valign": "vcenter",
                "border": 1,
                "font_size": 9
            }),
            "data_numeric": workbook.add_format({
                "align": "center",
                "valign": "vcenter",
                "border": 1,
                "font_size": 9,
                "num_format": "0.00"
            }),
            # Legend formats
            "legend_header": workbook.add_format({
                "bold": True,
                "bg_color": "#366092",
                "font_color": "#FFFFFF",
                "align": "center",
                "valign": "vcenter",
                "border": 1,
                "font_size": 11
            }),
            "legend_category": workbook.add_format({
                "bold": True,
                "bg_color": "#E7E6E6",
                "align": "left",
                "valign": "vcenter",
                "border": 1,
                "font_size": 10
            }),
            "legend_data": workbook.add_format({
                "align": "left",
                "valign": "vcenter",
                "border": 1,
                "font_size": 9
            })
        }
        
        # Add color-specific formats for SQZMOM
        for color_name, color_info in self.COLOR_SCHEMES["SQZMOM"].items():
            formats[f"sqzmom_{color_name}"] = workbook.add_format({
                "bg_color": color_info["bg"],
                "font_color": color_info["font"],
                "align": "center",
                "valign": "vcenter",
                "border": 1,
                "bold": True,
                "font_size": 9
            })
        
        # Add color-specific formats for FINAL
        for color_name, color_info in self.COLOR_SCHEMES["FINAL"].items():
            formats[f"final_{color_name.replace(' ', '_')}"] = workbook.add_format({
                "bg_color": color_info["bg"],
                "font_color": color_info["font"],
                "align": "center",
                "valign": "vcenter",
                "border": 1,
                "bold": True,
                "font_size": 9
            })
        
        return formats
    
    def _create_legend_sheet(self, workbook: xlsxwriter.Workbook, formats: Dict[str, Any]):
        """
        Create the Legend sheet with color code explanations.
        
        Args:
            workbook: xlsxwriter Workbook object
            formats: Dictionary of format objects
        """
        worksheet = workbook.add_worksheet("Legend")
        
        # Set column widths
        worksheet.set_column("A:A", 25)
        worksheet.set_column("B:B", 20)
        worksheet.set_column("C:C", 25)
        worksheet.set_column("D:D", 50)
        
        # Write headers
        headers = ["Indicator/Sheet", "Status/Value", "Color", "Description"]
        for col, header in enumerate(headers):
            worksheet.write(0, col, header, formats["legend_header"])
        
        row = 1
        
        # SQZMOM Legend
        worksheet.write(row, 0, "SQZMOM (Momentum)", formats["legend_category"])
        row += 1
        
        for color_name, color_info in self.COLOR_SCHEMES["SQZMOM"].items():
            worksheet.write(row, 0, "SQZMOM")
            worksheet.write(row, 1, color_name, formats["legend_data"])
            
            # Color cell
            color_format = workbook.add_format({
                "bg_color": color_info["bg"],
                "font_color": color_info["font"],
                "align": "center",
                "valign": "vcenter",
                "border": 1,
                "bold": True
            })
            worksheet.write(row, 2, "■", color_format)
            worksheet.write(row, 3, color_info["desc"], formats["legend_data"])
            row += 1
        
        # Add spacing
        row += 1
        
        # FINAL Legend
        worksheet.write(row, 0, "FINAL (Sync & Buildup)", formats["legend_category"])
        row += 1
        
        for color_name, color_info in self.COLOR_SCHEMES["FINAL"].items():
            worksheet.write(row, 0, "FINAL")
            worksheet.write(row, 1, color_name, formats["legend_data"])
            
            # Color cell
            color_format = workbook.add_format({
                "bg_color": color_info["bg"],
                "font_color": color_info["font"],
                "align": "center",
                "valign": "vcenter",
                "border": 1,
                "bold": True
            })
            worksheet.write(row, 2, "■", color_format)
            worksheet.write(row, 3, color_info["desc"], formats["legend_data"])
            row += 1
        
        logger.info("✓ Legend sheet created")
    
    def _create_timeseries_sheet(
        self,
        workbook: xlsxwriter.Workbook,
        sheet_name: str,
        df: pd.DataFrame,
        formats: Dict[str, Any]
    ):
        """
        Create a time-series sheet with standard formatting.
        
        Args:
            workbook: xlsxwriter Workbook object
            sheet_name: Name of the sheet
            df: DataFrame with Symbol in col 0, Metrics in col 1, time data from col 2+
            formats: Dictionary of format objects
        """
        worksheet = workbook.add_worksheet(sheet_name)
        
        # Set column widths
        worksheet.set_column(0, 0, 15)  # Symbol column
        worksheet.set_column(1, 1, 18)  # Metrics column
        worksheet.set_column(2, len(self.TIME_INTERVALS) + 1, 12)  # Time columns
        
        # Write headers
        df_columns = list(df.columns)
        
        # Header row
        worksheet.write(0, 0, df_columns[0] if len(df_columns) > 0 else "Symbol", formats["header_frozen"])
        worksheet.write(0, 1, df_columns[1] if len(df_columns) > 1 else "Metrics", formats["header_frozen"])
        
        for col_idx, column in enumerate(df_columns[2:], start=2):
            worksheet.write(0, col_idx, column, formats["header"])
        
        # Write data
        for row_idx, (_, row_data) in enumerate(df.iterrows(), start=1):
            # Symbol (frozen)
            worksheet.write(row_idx, 0, row_data.iloc[0], formats["data_left"])
            # Metrics (frozen)
            worksheet.write(row_idx, 1, row_data.iloc[1], formats["data_left"])
            
            # Time-series data
            for col_idx, value in enumerate(row_data.iloc[2:], start=2):
                if isinstance(value, float):
                    worksheet.write(row_idx, col_idx, value, formats["data_numeric"])
                else:
                    worksheet.write(row_idx, col_idx, value, formats["data"])
        
        # Freeze panes: freeze first two columns and header row
        worksheet.freeze_panes(1, 2)
        
        logger.info(f"✓ Sheet '{sheet_name}' created with {len(df)} data rows")
    
    def _create_timeseries_sheet_with_colors(
        self,
        workbook: xlsxwriter.Workbook,
        sheet_name: str,
        df: pd.DataFrame,
        formats: Dict[str, Any]
    ):
        """
        Create a time-series sheet with conditional color formatting.
        Applies colors based on cell values (SQZMOM and FINAL sheets).
        
        Args:
            workbook: xlsxwriter Workbook object
            sheet_name: Name of the sheet (SQZMOM or FINAL)
            df: DataFrame with Symbol in col 0, Metrics in col 1, time data from col 2+
            formats: Dictionary of format objects
        """
        worksheet = workbook.add_worksheet(sheet_name)
        
        # Set column widths
        worksheet.set_column(0, 0, 15)  # Symbol column
        worksheet.set_column(1, 1, 18)  # Metrics column
        worksheet.set_column(2, len(self.TIME_INTERVALS) + 1, 12)  # Time columns
        
        # Get the appropriate color scheme
        color_scheme = self.COLOR_SCHEMES.get(sheet_name, {})
        
        # Write headers
        df_columns = list(df.columns)
        
        worksheet.write(0, 0, df_columns[0] if len(df_columns) > 0 else "Symbol", formats["header_frozen"])
        worksheet.write(0, 1, df_columns[1] if len(df_columns) > 1 else "Metrics", formats["header_frozen"])
        
        for col_idx, column in enumerate(df_columns[2:], start=2):
            worksheet.write(0, col_idx, column, formats["header"])
        
        # Write data with color mapping
        for row_idx, (_, row_data) in enumerate(df.iterrows(), start=1):
            # Symbol (frozen)
            worksheet.write(row_idx, 0, row_data.iloc[0], formats["data_left"])
            # Metrics (frozen)
            worksheet.write(row_idx, 1, row_data.iloc[1], formats["data_left"])
            
            # Time-series data with conditional formatting
            for col_idx, value in enumerate(row_data.iloc[2:], start=2):
                cell_value = str(value).strip() if pd.notna(value) else ""
                
                # Determine format based on cell value
                if cell_value in color_scheme:
                    format_key = f"{sheet_name.lower()}_{cell_value.replace(' ', '_')}"
                    cell_format = formats.get(format_key, formats["data"])
                else:
                    cell_format = formats["data"]
                
                worksheet.write(row_idx, col_idx, cell_value, cell_format)
        
        # Freeze panes: freeze first two columns and header row
        worksheet.freeze_panes(1, 2)
        
        logger.info(f"✓ Colored sheet '{sheet_name}' created with {len(df)} data rows")
    
    def _create_empty_sheet(
        self,
        workbook: xlsxwriter.Workbook,
        sheet_name: str,
        formats: Dict[str, Any]
    ):
        """
        Create an empty sheet with proper structure.
        
        Args:
            workbook: xlsxwriter Workbook object
            sheet_name: Name of the sheet
            formats: Dictionary of format objects
        """
        worksheet = workbook.add_worksheet(sheet_name)
        
        # Set column widths
        worksheet.set_column(0, 0, 15)
        worksheet.set_column(1, 1, 18)
        worksheet.set_column(2, 75, 12)
        
        # Write headers
        worksheet.write(0, 0, "Symbol", formats["header_frozen"])
        worksheet.write(0, 1, "Metrics", formats["header_frozen"])
        
        for col_idx, time_interval in enumerate(self.TIME_INTERVALS, start=2):
            worksheet.write(0, col_idx, time_interval, formats["header"])
        
        # Freeze panes
        worksheet.freeze_panes(1, 2)
        
        logger.info(f"✓ Empty sheet '{sheet_name}' created")
    
    def generate_sample_dataframes(self) -> Dict[str, pd.DataFrame]:
        """
        Generate sample dataframes for all sheets for testing purposes.
        
        Returns:
            Dictionary of sample DataFrames
        """
        sample_symbols = ["RELIANCE", "TCS", "INFY", "HDFC", "LT"]
        
        dataframes = {}
        
        # Reference sheet
        dataframes["Reference"] = pd.DataFrame({
            "Symbol": sample_symbols,
            "Sector": ["Energy", "IT", "IT", "Finance", "Engineering"],
            "Market_Cap": ["7.5L Cr", "13.8L Cr", "7.2L Cr", "4.8L Cr", "2.3L Cr"]
        })
        
        # Time-series sheets
        for sheet_name in ["R1S1", "Pine", "SQZMOM", "Breakout", "RSI", "ADX", "VWAP", "MTF_Trend", "OI_Dynamics", "15m_5m_Sync"]:
            data = []
            for symbol in sample_symbols:
                for metric in ["Open", "High", "Low", "Close"]:
                    row = {"Symbol": symbol, "Metrics": metric}
                    for time_interval in self.TIME_INTERVALS:
                        import random
                        if sheet_name == "SQZMOM":
                            row[time_interval] = random.choice(["LIME", "GREEN", "RED", "MAROON", "ORANGE", "BLACK"])
                        else:
                            row[time_interval] = round(random.uniform(100, 500), 2)
                    data.append(row)
            dataframes[sheet_name] = pd.DataFrame(data)
        
        # FINAL sheet with colors
        final_data = []
        for symbol in sample_symbols:
            for metric in ["Synchronization", "Buildup"]:
                row = {"Symbol": symbol, "Metrics": metric}
                for time_interval in self.TIME_INTERVALS:
                    import random
                    row[time_interval] = random.choice(["BULL SYNC", "BEAR SYNC", "NO SYNC", "Long Buildup", "Short Buildup", "WAIT"])
                final_data.append(row)
        dataframes["FINAL"] = pd.DataFrame(final_data)
        
        # Order sheet
        order_data = []
        for i, symbol in enumerate(sample_symbols):
            order_data.append({
                "Symbol": symbol,
                "Order_Type": "BUY" if i % 2 == 0 else "SELL",
                "Entry_Time": "09:30",
                "Entry_Price": round(random.uniform(100, 500), 2),
                "Exit_Time": "14:45",
                "Exit_Price": round(random.uniform(100, 500), 2),
                "PnL": round(random.uniform(-1000, 5000), 2)
            })
        dataframes["Order"] = pd.DataFrame(order_data)
        
        logger.info("✓ Sample dataframes generated")
        return dataframes
    
    def get_export_directory(self) -> str:
        """Get the Excel export directory."""
        return self.excel_dir

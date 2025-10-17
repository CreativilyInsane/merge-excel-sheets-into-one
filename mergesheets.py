#!/usr/bin/env python3
"""
Excel Sheet Consolidator
Converts multiple sheets from a specified range into a single consolidated sheet.
"""

import argparse
import sys
import os
import pandas as pd
from tqdm import tqdm
import signal
import subprocess
import platform
from typing import List, Tuple, Dict, Any, Optional
import emoji
import json
from datetime import datetime
import re

class ColumnConfig:
    """Class to handle column configuration and transformations"""
    
    @staticmethod
    def apply_column_properties(df: pd.DataFrame, col_config: Dict[str, Dict[str, Any]]) -> pd.DataFrame:
        """Apply column properties to dataframe"""
        if not col_config:
            return df
            
        result_df = df.copy()
        
        for col_name, properties in col_config.items():
            if col_name in result_df.columns:
                try:
                    # Apply word count if specified
                    if properties.get('word_count', False):
                        result_df[f'{col_name}_word_count'] = result_df[col_name].astype(str).apply(
                            lambda x: len(x.split()) if pd.notna(x) and x.strip() else 0
                        )
                    
                    # Apply data type conversion
                    dtype = properties.get('dtype')
                    if dtype:
                        result_df[col_name] = ColumnConfig.convert_dtype(result_df[col_name], dtype)
                        
                except Exception as e:
                    print(f"⚠️  Warning: Could not apply properties to column '{col_name}': {e}")
        
        return result_df
    
    @staticmethod
    def convert_dtype(series: pd.Series, dtype: str) -> pd.Series:
        """Convert series to specified data type"""
        dtype = dtype.lower()
        
        if dtype in ['string', 'str', 'text']:
            return series.astype(str)
        elif dtype in ['int', 'integer', 'number']:
            return pd.to_numeric(series, errors='coerce').astype('Int64')
        elif dtype in ['float', 'decimal']:
            return pd.to_numeric(series, errors='coerce').astype(float)
        elif dtype in ['bool', 'boolean']:
            return series.astype(str).str.lower().map({'true': True, 'false': False, '1': True, '0': False})
        elif dtype in ['date', 'datetime']:
            return pd.to_datetime(series, errors='coerce')
        elif dtype == 'category':
            return series.astype('category')
        else:
            return series

class SheetConsolidator:
    def __init__(self):
        self.interrupted = False
        self.column_config = {}
        self.setup_signal_handlers()
    
    def setup_signal_handlers(self):
        """Handle Ctrl+C gracefully"""
        signal.signal(signal.SIGINT, self.signal_handler)
    
    def signal_handler(self, signum, frame):
        """Handle interrupt signals"""
        self.interrupted = True
        print(f"\n\n{emoji.emojize(':warning:')}  Operation interrupted by user!")
        print(f"{emoji.emojize(':hourglass_not_done:')}  Cleaning up...")
        sys.exit(1)
    
    def load_column_config(self, config_path: Optional[str]) -> bool:
        """Load column configuration from JSON file"""
        if not config_path:
            self.column_config = {}
            return True
            
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                self.column_config = json.load(f)
            
            print(f"{emoji.emojize(':gear:')}  Loaded column configuration from: {config_path}")
            self.display_column_config()
            return True
            
        except FileNotFoundError:
            print(f"❌ Column config file not found: {config_path}")
            return False
        except json.JSONDecodeError as e:
            print(f"❌ Invalid JSON in config file: {e}")
            return False
        except Exception as e:
            print(f"❌ Error loading config file: {e}")
            return False
    
    def display_column_config(self):
        """Display the loaded column configuration"""
        if not self.column_config:
            return
            
        print(f"{emoji.emojize(':clipboard:')}  Column Configuration:")
        for col_name, properties in self.column_config.items():
            prop_str = []
            if properties.get('word_count'):
                prop_str.append("word_count")
            if properties.get('dtype'):
                prop_str.append(f"dtype={properties['dtype']}")
            print(f"    📊 {col_name}: {', '.join(prop_str) if prop_str else 'no transformations'}")
    
    def create_config_template(self, input_file: str, sheet_range: str) -> bool:
        """Create a template configuration file based on the input file structure"""
        try:
            sheet_names = self.get_sheet_names(input_file)
            target_sheets = self.parse_sheet_range(sheet_range, len(sheet_names))
            
            # Sample first sheet to get column structure
            sample_sheet = sheet_names[target_sheets[0]]
            df_sample = pd.read_excel(input_file, sheet_name=sample_sheet, nrows=5)
            
            config_template = {}
            for col in df_sample.columns:
                config_template[str(col)] = {
                    "word_count": False,
                    "dtype": "auto",  # auto, string, int, float, bool, date, category
                    "description": f"Column: {col}"
                }
            
            template_filename = f"column_config_template_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
            with open(template_filename, 'w', encoding='utf-8') as f:
                json.dump(config_template, f, indent=2, ensure_ascii=False)
            
            print(f"\n{emoji.emojize(':sparkles:')}  Created column configuration template: {template_filename}")
            print(f"{emoji.emojize(':information:')}  Edit this file and use it with --config option")
            return True
            
        except Exception as e:
            print(f"❌ Failed to create config template: {e}")
            return False
    
    def parse_sheet_range(self, range_str: str, total_sheets: int) -> List[int]:
        """Parse sheet range string into list of sheet indices"""
        try:
            if '-' in range_str:
                start, end = map(int, range_str.split('-'))
                if start < 1 or end > total_sheets or start > end:
                    raise ValueError("Invalid range")
                return list(range(start - 1, end))  # Convert to 0-based indexing
            else:
                sheets = [int(x.strip()) - 1 for x in range_str.split(',')]  # Convert to 0-based
                if any(sheet < 0 or sheet >= total_sheets for sheet in sheets):
                    raise ValueError("Sheet number out of range")
                return sheets
        except ValueError as e:
            raise ValueError(f"Invalid sheet range format: {range_str}. Use format like '1-5' or '1,3,5'") from e
    
    def validate_files(self, input_path: str, output_path: str):
        """Validate input and output file paths"""
        if not os.path.exists(input_path):
            raise FileNotFoundError(f"📁 Input file not found: {input_path}")
        
        output_dir = os.path.dirname(output_path)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir, exist_ok=True)
    
    def get_sheet_names(self, file_path: str) -> List[str]:
        """Get all sheet names from the Excel file"""
        try:
            return pd.ExcelFile(file_path).sheet_names
        except Exception as e:
            raise Exception(f"❌ Failed to read Excel file: {e}")
    
    def consolidate_sheets(self, input_file: str, output_file: str, sheet_range: str) -> bool:
        """Main function to consolidate sheets"""
        print(f"\n{emoji.emojize(':rocket:')}  Starting Sheet Consolidation")
        print(f"{emoji.emojize(':file_folder:')}  Input: {input_file}")
        print(f"{emoji.emojize(':floppy_disk:')}  Output: {output_file}")
        print(f"{emoji.emojize(':page_facing_up:')}  Sheet Range: {sheet_range}")
        
        if self.column_config:
            print(f"{emoji.emojize(':gear:')}  Column Properties: ENABLED")
        else:
            print(f"{emoji.emojize(':gear:')}  Column Properties: DISABLED (using raw data)")
        
        print("=" * 60)
        
        # Validate files
        self.validate_files(input_file, output_file)
        
        # Get sheet names
        try:
            sheet_names = self.get_sheet_names(input_file)
            total_sheets = len(sheet_names)
            print(f"{emoji.emojize(':spiral_notepad:')}  Total sheets found: {total_sheets}")
            
            # Parse sheet range
            target_sheets = self.parse_sheet_range(sheet_range, total_sheets)
            target_sheet_names = [sheet_names[i] for i in target_sheets]
            
            print(f"{emoji.emojize(':dart:')}  Processing {len(target_sheets)} sheets: {', '.join(target_sheet_names)}")
            
        except Exception as e:
            print(f"❌ Error: {e}")
            return False
        
        # Process sheets
        consolidated_data = []
        processed_count = 0
        
        with tqdm(total=len(target_sheets), 
                 bar_format="{l_bar}%s{bar}%s{r_bar}" % (emoji.emojize(':blue_square:'), emoji.emojize(':green_square:')),
                 desc=f"{emoji.emojize(':hourglass_flowing_sand:')} Processing sheets",
                 unit="sheet") as pbar:
            
            for sheet_idx in target_sheets:
                if self.interrupted:
                    break
                
                sheet_name = sheet_names[sheet_idx]
                try:
                    # Read sheet with progress update
                    df = pd.read_excel(input_file, sheet_name=sheet_name)
                    
                    # Apply column properties if configuration exists
                    if self.column_config:
                        df = ColumnConfig.apply_column_properties(df, self.column_config)
                    
                    # Add source sheet name as a column
                    df['_Source_Sheet'] = sheet_name
                    
                    consolidated_data.append(df)
                    processed_count += 1
                    
                    pbar.set_postfix_str(f"📊 {sheet_name}")
                    pbar.update(1)
                    
                except Exception as e:
                    print(f"\n⚠️  Warning: Failed to process sheet '{sheet_name}': {e}")
                    pbar.update(1)
                    continue
        
        if self.interrupted:
            return False
        
        if not consolidated_data:
            print(f"\n{emoji.emojize(':x:')}  No data was processed successfully!")
            return False
        
        # Combine all data
        print(f"\n{emoji.emojize(':card_file_box:')}  Combining data from {processed_count} sheets...")
        try:
            combined_df = pd.concat(consolidated_data, ignore_index=True)
            
            # Save to output file
            print(f"{emoji.emojize(':inbox_tray:')}  Saving to {output_file}...")
            combined_df.to_excel(output_file, index=False, sheet_name='Consolidated_Data')
            
        except Exception as e:
            print(f"❌ Error combining/saving data: {e}")
            return False
        
        return True
    
    def open_file(self, file_path: str):
        """Open the output file using system default application"""
        try:
            system = platform.system()
            if system == "Darwin":  # macOS
                subprocess.run(["open", file_path])
            elif system == "Windows":
                os.startfile(file_path)
            else:  # Linux and other Unix-like
                subprocess.run(["xdg-open", file_path])
            print(f"{emoji.emojize(':eyes:')}  Opening output file...")
        except Exception as e:
            print(f"⚠️  Note: Could not open file automatically: {e}")
            print(f"📁 File saved at: {file_path}")

def main():
    parser = argparse.ArgumentParser(
        description="📊 Excel Sheet Consolidator - Convert multiple sheets to a single consolidated sheet",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python sheet_consolidator.py input.xlsx output.xlsx 1-5
  python sheet_consolidator.py data.xlsx consolidated.xlsx 1,3,5,7 --config columns.json
  python sheet_consolidator.py input.xlsx output.xlsx 1-3 --create-template
  python sheet_consolidator.py "input file.xlsx" "output file.xlsx" 2-8 --no-open

Column Configuration JSON Format:
{
  "ColumnName1": {
    "word_count": true,
    "dtype": "string"
  },
  "ColumnName2": {
    "word_count": false,
    "dtype": "int"
  }
}

Supported Data Types: string, int, float, bool, date, category
        """
    )
    
    parser.add_argument('input_file', help='📁 Path to input Excel file')
    parser.add_argument('output_file', help='💾 Path to output Excel file')
    parser.add_argument('sheet_range', help='🔢 Sheet range (e.g., 1-5, 1,3,5)')
    parser.add_argument('--config', help='⚙️ Path to column configuration JSON file')
    parser.add_argument('--create-template', action='store_true', 
                       help='📝 Create a template column configuration file')
    parser.add_argument('--no-open', action='store_true', 
                       help="👀 Don't open the file after completion")
    
    args = parser.parse_args()
    
    # Create consolidator instance
    consolidator = SheetConsolidator()
    
    try:
        # Handle template creation
        if args.create_template:
            success = consolidator.create_config_template(args.input_file, args.sheet_range)
            sys.exit(0 if success else 1)
        
        # Load column configuration if provided
        if args.config and not consolidator.load_column_config(args.config):
            sys.exit(1)
        
        # Perform consolidation
        success = consolidator.consolidate_sheets(args.input_file, args.output_file, args.sheet_range)
        
        if success:
            print(f"\n{emoji.emojize(':party_popper:')}  SUCCESS!")
            print(f"{emoji.emojize(':white_heavy_check_mark:')}  Consolidated {args.sheet_range} sheets into: {args.output_file}")
            
            # Open file if not disabled
            if not args.no_open:
                consolidator.open_file(args.output_file)
            else:
                print(f"{emoji.emojize(':file_folder:')}  Output saved to: {args.output_file}")
            
            print(f"\n{emoji.emojize(':bar_chart:')}  Task completed successfully!")
            sys.exit(0)
        else:
            print(f"\n{emoji.emojize(':x:')}  FAILED!")
            sys.exit(1)
            
    except Exception as e:
        print(f"\n{emoji.emojize(':red_exclamation_mark:')}  ERROR: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()

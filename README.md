# merge-excel-sheets-into-one

📊 Excel Sheet Consolidator
A professional Python tool for consolidating multiple Excel sheets into a single sheet with advanced column configuration, progress tracking, and smart data type detection.

https://img.shields.io/badge/Python-3.7+-blue.svg
https://img.shields.io/badge/Pandas-1.3+-green.svg
https://img.shields.io/badge/OpenPyXL-3.0+-orange.svg

🚀 Features
📑 Multi-Sheet Consolidation: Combine multiple Excel sheets into a single sheet

🎯 Smart Column Detection: Automatic column type detection based on names and content

🔤 Space-Aware Processing: Intelligent handling of multi-word column names

📊 Advanced Word Count: Configurable word counting with ranges and filters

🔄 Data Type Conversion: Support for string, int, float, bool, date, and category types

📈 Progress Tracking: Beautiful progress bars with emoji indicators

🎨 Professional UI: Colorful terminal output with emojis

⚡ Graceful Interruption: Proper handling of Ctrl+C signals

🔧 Flexible Configuration: JSON-based column configuration

📁 Auto-File Opening: Automatically opens output file after completion

📦 Installation
Prerequisites
Python 3.7 or higher

pip (Python package manager)

Install Dependencies
bash
pip install pandas openpyxl tqdm emoji
Download Script
bash
git clone <repository-url>
cd excel-sheet-consolidator
🛠️ Usage
Basic Usage
bash
# Consolidate sheets 1-5
python sheet_consolidator.py input.xlsx output.xlsx 1-5

# Consolidate specific sheets
python sheet_consolidator.py data.xlsx consolidated.xlsx 1,3,5,7

# Files with spaces in names
python sheet_consolidator.py "my input file.xlsx" "consolidated output.xlsx" 2-8
Advanced Usage with Column Configuration
bash
# With custom column configuration
python sheet_consolidator.py input.xlsx output.xlsx 1-5 --config column_config.json

# Don't open file after completion
python sheet_consolidator.py input.xlsx output.xlsx 1-3 --no-open
Template Generation
bash
# Create smart detection template
python sheet_consolidator.py input.xlsx output.xlsx 1-3 --create-template smart_detect

# Create space-aware template
python sheet_consolidator.py input.xlsx output.xlsx 1-3 --create-template space_aware

# Create text-only template
python sheet_consolidator.py input.xlsx output.xlsx 1-3 --create-template text_only

# Create advanced template with examples
python sheet_consolidator.py input.xlsx output.xlsx 1-3 --advanced-template
📋 Command Line Arguments
Argument	Description	Example
input_file	Input Excel file path	data.xlsx
output_file	Output Excel file path	consolidated.xlsx
sheet_range	Sheet range to process	1-5 or 1,3,5
--config	Column configuration JSON file	--config columns.json
--create-template	Create configuration template	--create-template smart_detect
--advanced-template	Create advanced template with examples	--advanced-template
--no-open	Don't open file after completion	--no-open
📁 Configuration Files
Column Configuration JSON Structure
json
{
  "ColumnName": {
    "word_count": false,
    "dtype": "auto",
    "description": "Column description"
  }
}
Advanced Word Count Configuration
json
{
  "Product Description": {
    "word_count": {
      "min_length": 2,
      "max_length": 25,
      "start": 0,
      "end": 10,
      "allowed_chars": "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ -",
      "exclude_chars": "0123456789"
    },
    "dtype": "string"
  }
}
Configuration Options
Word Count Options
Simple: "word_count": true - Basic word counting

Advanced: Object with range and filter options:

min_length: Minimum word length (default: 0)

max_length: Maximum word length (default: infinite)

start: Start index for word range (0-based)

end: End index for word range (exclusive)

allowed_chars: Only count words containing these characters

exclude_chars: Exclude words containing these characters

Data Type Options
auto - Automatic detection (default)

string - Convert to string type

int - Convert to integer (handles missing values)

float - Convert to float

bool - Convert to boolean

date - Convert to datetime

category - Convert to categorical data

🎯 Template Types
1. smart_detect
Detects column types based on name patterns and content

Analyzes data statistics (average length, word count)

Provides intelligent defaults

2. space_aware
Focuses on multi-word column names

Automatically enables word count for descriptive columns

Perfect for distinguishing identifiers vs descriptions

3. text_only
Configures only text-like columns

Leaves other columns with minimal configuration

Optimized for text-heavy datasets

4. all
Basic configuration for all columns

Simple template for manual customization

🔍 Smart Detection Features
Column Name Pattern Recognition
Text Columns: description, note, comment, remark, text

Numeric Columns: price, cost, amount, quantity, number

Date Columns: date, time, created, modified

Boolean Columns: is_, has_, flag, status, active

Space-Based Intelligence
Multi-word names: Automatically detected and configured for word count

Single-word names: Minimal configuration for identifiers

Name analysis: Word count and space detection in column names

📊 Output Features
Consolidated Data: All selected sheets combined into one

Source Tracking: _Source_Sheet column added to track original sheet

Word Count Columns: {column_name}_word_count for configured columns

Data Type Consistency: Uniform data types across all sheets

Error Handling: Continues processing even if individual sheets fail

🛡️ Error Handling
File Validation: Checks input file existence and output directory

Sheet Range Validation: Validates sheet numbers and ranges

Graceful Interruption: Proper cleanup on Ctrl+C

Error Recovery: Continues processing other sheets if one fails

Detailed Logging: Clear error messages with emoji indicators

🎨 Progress Indicators
Progress Bars: Visual progress with emoji-enhanced bars

Real-time Updates: Current sheet being processed

Completion Summary: Success/failure statistics

Colorful Output: Professional terminal presentation

📝 Examples
Example 1: Basic Consolidation
bash
python sheet_consolidator.py sales_data.xlsx consolidated_sales.xlsx 1-12
Example 2: Advanced Configuration
bash
# Create template
python sheet_consolidator.py data.xlsx output.xlsx 1-3 --create-template smart_detect

# Edit generated template, then use it
python sheet_consolidator.py data.xlsx output.xlsx 1-3 --config column_config_smart_detect_20240115_103000.json
Example 3: Selective Processing
bash
# Process only quarterly sheets (Q1, Q2, Q3, Q4)
python sheet_consolidator.py annual_data.xlsx quarterly_summary.xlsx 1,4,7,10 --no-open
🔧 Advanced Configuration Examples
Text Analysis Configuration
json
{
  "Customer Feedback": {
    "word_count": {
      "min_length": 3,
      "max_length": 20,
      "start": 0,
      "end": 50,
      "exclude_chars": "!@#$%^&*()"
    },
    "dtype": "string"
  }
}
Mixed Data Types
json
{
  "Product ID": {
    "word_count": false,
    "dtype": "string"
  },
  "Product Name": {
    "word_count": true,
    "dtype": "string"
  },
  "Price": {
    "word_count": false,
    "dtype": "float"
  },
  "In Stock": {
    "word_count": false,
    "dtype": "bool"
  },
  "Last Updated": {
    "word_count": false,
    "dtype": "date"
  }
}
🐛 Troubleshooting
Common Issues
File Not Found

text
❌ Input file not found: data.xlsx
Solution: Check file path and permissions

Invalid Sheet Range

text
❌ Invalid sheet range format: 1-15. Use format like '1-5' or '1,3,5'
Solution: Verify sheet numbers exist in the file

JSON Configuration Error

text
❌ Invalid JSON in config file: Expecting property name enclosed in double quotes
Solution: Validate JSON syntax using a JSON validator

Permission Denied

text
❌ Error combining/saving data: [Errno 13] Permission denied: 'output.xlsx'
Solution: Check write permissions in output directory

Debug Mode
For detailed debugging, you can modify the script to add print statements or use Python's logging module.

📄 License
This project is open source and available under the MIT License.

🤝 Contributing
Contributions are welcome! Please feel free to submit pull requests or open issues for bugs and feature requests.

📞 Support
If you encounter any problems or have questions:

Check the troubleshooting section above

Open an issue on GitHub

Provide your Excel file structure and configuration

🎉 Acknowledgments
Built with ❤️ using Python

Powered by Pandas for data manipulation

Enhanced with emojis for better user experience

Professional progress tracking with tqdm

Happy Data Consolidating! 🎊📈

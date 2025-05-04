# ECS2XLS - XML to XLSX Converter for FLS ECS8 Trend Data

![Demo](resources/xml2xls.gif)

A powerful utility for converting XML trend data exports from FLS ECS8 SCADA system into Excel (XLSX) format. This tool efficiently processes large datasets and provides a user-friendly way to analyze trend data in Excel.

## Features

- Converts FLS ECS8 trend export XML files to Excel format
- Handles large datasets with automatic sheet splitting
- Preserves timestamps with Magadan timezone (+11 UTC) adjustment
- Progress tracking with visual indicators
- Detailed logging and error handling
- Customizable maximum rows per Excel sheet

## Requirements

- Python 3.x
- Required Python packages:
  - pandas
  - xlsxwriter
  - alive-progress
  - colorama

## Installation

1. Clone this repository:
```bash
git clone https://github.com/yourusername/xml2xls.git
cd xml2xls
```

2. Install required packages:
```bash
pip install -r requirements.txt
```

## Usage

1. Export trend data from ECS8:
   - Navigate to: `ECS8 -> Trend -> Trend Export -> Logged values`
   - Default export filename format: `Trend on YYYY-MM-DDTHH.MM.SS.xml`

2. Run the converter:
```bash
python xml2xls.py "path/to/your/Trend on 2024-08-30T21.23.17.xml"
```

Optional parameters:
- `-max_row`: Set maximum number of rows per Excel sheet (default: 1,000,000)

The converter will create an Excel file with the same name as the input XML file, appended with `.xlsx`.

## Output Format

- Each tag is saved in a separate Excel sheet
- For large datasets, sheets are automatically split to maintain Excel's row limit
- Timestamps are adjusted to Magadan timezone (+11 UTC)
- Includes metadata about the data range and tag information

## Development

This project is written in Python and uses the following main libraries:
- `pandas` for data manipulation
- `xlsxwriter` for Excel file creation
- `alive-progress` for progress visualization
- `colorama` for colored console output

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Author

7Art - 2023


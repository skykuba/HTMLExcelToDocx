# Word Document Generator from HTML data in Excel

## Description

This project allows you to automatically generate Word (`.docx`) files based on an HTML template and data from an Excel file. Each row from the Excel file generates a separate Word document with the fields filled in accordingly.

## Requirements

- Python 3
- pip

## Installation

1. Clone the repository.
2. Install the required libraries:
```python
pip install -r requirements.txt
```
## Usage

1. Place the Excel data file in the `dane/` directory (default: `testDane.xlsx`).
2. Create your own HTML template (template.html) with proper headers corresponding to the Excel sheet, and make sure it is placed in the main project directory.
3. Run the script:
```python
python main.py
```


## License

Project licensed under the MIT License.
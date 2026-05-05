# IT23562592 - Assignment 1 - Transliteration Accuracy Testing

## Prerequisites
- Python 3.11 or 3.12
- pip

## Installation
pip install -U pip
pip install playwright openpyxl
playwright install

## Running the Tests
cd D:\test_automation
py test_automation.py --excel "IT23562592_Assignment 1 - Test cases.xlsx" --url "https://www.pixelssuite.com/chat-translator" --wait-ms 8000 --type-delay-ms 80 --slow-mo-ms 200 --save-every 1 --keep-open

# Stock Support Level Calculator

This program calculates and displays three support levels for foreign stocks using historical price data.

## Features
- Calculates 3 support levels based on historical price data
- Real-time stock data using Yahoo Finance
- Simple and intuitive Windows GUI interface
- Automatic calculation when pressing Enter

## Installation

1. Make sure you have Python 3.7 or higher installed
2. Install the required packages:
```
pip install -r requirements.txt
```

## Usage

1. Run the program:
```
python stock_support_calculator.py
```

2. Enter a stock symbol (e.g., AAPL for Apple, GOOGL for Google)
3. Press Enter or click "Calculate Support Levels"
4. The program will display:
   - Three support levels
   - Current stock price

## Notes
- Stock symbols should be entered in their standard format (e.g., AAPL, GOOGL, MSFT)
- The program uses 1 year of historical data to calculate support levels
- Support levels are calculated based on recent price lows 
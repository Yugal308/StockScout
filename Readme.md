# Stock Analysis Web Application

A Flask-based web application for analyzing stock data from Excel files. The application fetches financial metrics using Yahoo Finance API and provides recommendations based on fundamental analysis.

## Features

- Upload Excel files containing stock symbols
- Fetch real-time financial data using yfinance
- Calculate Buy/Sell/Hold recommendations based on:
  - PE Ratio
  - PEG Ratio
  - Dividend Yield
  - Debt to Equity
  - Price relative to 52-week high
- Download processed Excel file with analysis results
- Clean and responsive user interface

## Requirements

- Python 3.8+
- Required packages (see requirements.txt)

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/stock-analysis-app.git
   cd stock-analysis-app
   ```

2. Create a virtual environment:
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows use: venv\Scripts\activate
   ```

3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

4. Run the application:
   ```bash
   python app.py
   ```

5. Access the application in your browser at:
   ```
   http://localhost:5000
   ```

## Usage

1. Prepare an Excel file with stock symbols in a column named "SYMBOL"
2. Access the web interface
3. Upload your Excel file
4. Wait for the analysis to complete
5. Download the processed file with analysis results

## File Format

Your input Excel file should have:
- A column named "SYMBOL" containing stock symbols
- Optional column "Exchange" (defaults to NSE if not provided)

Example:

| SYMBOL | Exchange |
|--------|----------|
| INFY   | NSE      |
| TCS    | BSE      |

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Contributing

Contributions are welcome! Please open an issue or submit a pull request.

## Disclaimer

This application is for educational purposes only. The recommendations provided are based on simple financial metrics and should not be considered as professional investment advice. Always do your own research before making investment decisions.
# Value Scorer

A Python-based stock analysis tool for value investing. Evaluates companies based on quality metrics with emphasis on Return on Equity (ROE), debt levels, and free cash flow generation.

## Features

- **Quality Scoring** - Score stocks from 0-100 based on fundamental metrics
- **Intrinsic Value Calculation** - DCF and Graham formula valuations
- **Backtesting** - Validate the scoring methodology against historical data
- **Caching** - Avoid rate limits with intelligent data caching
- **Excel Reports** - Color-coded, sortable spreadsheets

## Quick Start

```bash
# Install dependencies
pip install yfinance pandas openpyxl numpy

# Create a ticker file
echo -e "AAPL\nMSFT\nGOOG\nJNJ" > stocks.txt

# Run analysis
python value_scorer.py stocks.txt
```

## Installation

```bash
# Clone the repository
git clone https://github.com/sushil-shrestha/value-scorer.git
cd value-scorer

# (Optional) Create virtual environment
python -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate

# Install dependencies
pip install yfinance pandas openpyxl numpy
```

## Usage

### Basic Analysis

```bash
# Analyze stocks from a file
python value_scorer.py stocks.txt

# Custom output file
python value_scorer.py stocks.txt -o my_analysis.xlsx

# Parallel processing (faster for many stocks)
python value_scorer.py stocks.txt -w 4
```

### Backtesting

Test how the scoring system would have performed historically:

```bash
# Run 3-year backtest (default)
python value_scorer.py stocks.txt --backtest

# Custom lookback period
python value_scorer.py stocks.txt --backtest --years 2

# Custom benchmark (default: SPY)
python value_scorer.py stocks.txt --backtest --benchmark QQQ
```

### Cache Management

The tool caches data to avoid Yahoo Finance rate limits:

```bash
# View cache statistics
python value_scorer.py stocks.txt --cache-stats

# Clear cache and fetch fresh data
python value_scorer.py stocks.txt --clear-cache

# Disable caching entirely
python value_scorer.py stocks.txt --no-cache

# Set custom cache expiry (hours)
python value_scorer.py stocks.txt --cache-expiry 48
```

### Input File Format

Create a text file with one ticker per line. Comments start with `#`:

```
# Technology
AAPL
MSFT
GOOGL

# Healthcare
JNJ
UNH
```

## Scoring System

### Quality Score (0-100)

| Component | Weight | Description |
|-----------|--------|-------------|
| **ROE** | 30 pts | Return on Equity - management efficiency |
| **Debt** | 25 pts | Debt-to-Equity ratio - financial risk |
| **Cash Flow** | 20 pts | Free Cash Flow yield - cash generation |
| **Valuation** | 15 pts | P/E and P/B ratios - price vs value |
| **Profitability** | 10 pts | Net and operating margins |

### Quality Ratings

| Rating | Score | Interpretation |
|--------|-------|----------------|
| A+ | 80-100 | Exceptional quality |
| A | 70-79 | High quality |
| B+ | 60-69 | Above average |
| B | 50-59 | Average |
| C+ | 40-49 | Below average |
| C | 30-39 | Poor quality |
| D | 0-29 | Avoid |

### Intrinsic Value Methods

The tool calculates intrinsic value using two methods:

1. **DCF (Discounted Cash Flow)**
   - Projects free cash flow for 5 years
   - Uses revenue growth as a constraint
   - Adjusts discount rate for risk (debt, profitability)
   - Conservative terminal value calculation

2. **Graham Formula**
   - `V = EPS × (8.5 + 2g)`
   - Where g = expected growth rate (capped at 15%)
   - Classic value investing approach

### Value Ratings

| Rating | Upside | Interpretation |
|--------|--------|----------------|
| Strong Buy | ≥50% | Significantly undervalued |
| Buy | 25-50% | Attractive opportunity |
| Hold | 10-25% | Moderate upside |
| Fair Value | ±10% | Fairly priced |
| Overvalued | <-10% | Above intrinsic value |

## Output

### Excel Report

The tool generates a formatted Excel file with:

**Sheet 1: Value Investment Analysis**
- All metrics for each stock
- Quality Score and Rating
- Intrinsic value calculations
- Color-coded cells (green=good, red=concern)

**Sheet 2: Scoring Methodology**
- Detailed scoring criteria
- Rating definitions

### Backtest Report

When running `--backtest`, you get:
- Historical Quality Score for each stock
- Actual price returns over the period
- Alpha vs benchmark (SPY by default)
- Performance breakdown by quality rating
- Win rate statistics

## Example Output

```
======================================================================
BACKTEST SUMMARY
======================================================================

Overall Statistics (8 stocks analyzed):
  Average Return: +96.7%
  Benchmark Return: +87.6%
  Average Alpha: +9.1%
  Win Rate (beat benchmark): 4/8 (50.0%)

Performance by Historical Quality Rating:
--------------------------------------------------
  A  :   3 stocks | Avg Return:  +130.6% | Win Rate:  66.7%
  B  :   4 stocks | Avg Return:   +73.6% | Win Rate:  50.0%
  D  :   1 stocks | Avg Return:   +87.4% | Win Rate:   0.0%

KEY INSIGHT
======================================================================
  High Quality (Score ≥60): Avg Return: +130.6%
  Low Quality (Score <40):  Avg Return: +87.4%
  Quality Premium: +43.1% (High quality outperformed)
```

## CLI Reference

```
usage: value_scorer.py [-h] [-o OUTPUT] [-w WORKERS] [--fresh]
                       [--no-cache] [--clear-cache] [--cache-expiry HOURS]
                       [--cache-dir DIR] [--cache-stats]
                       [--backtest] [--years N] [--benchmark TICKER]
                       ticker_file

Arguments:
  ticker_file           Text file with stock tickers (one per line)

Options:
  -o, --output FILE     Output Excel file (default: value_investment_analysis.xlsx)
  -w, --workers N       Parallel workers (default: 1)
  --fresh               Ignore existing results, start fresh

Cache Options:
  --no-cache            Disable caching
  --clear-cache         Clear cache before running
  --cache-expiry HOURS  Cache expiry time (default: 24)
  --cache-dir DIR       Cache directory (default: ~/.value_scorer_cache)
  --cache-stats         Show cache statistics and exit

Backtest Options:
  --backtest            Run backtest analysis
  --years N             Lookback period in years (default: 3)
  --benchmark TICKER    Benchmark ticker (default: SPY)
```

## Investment Philosophy

This tool is built on proven value investing principles:

1. **Quality First** - High ROE indicates competitive advantages
2. **Avoid Debt** - Leverage amplifies risk in downturns
3. **Cash is King** - Free cash flow is the true measure of value
4. **Margin of Safety** - Buy below intrinsic value
5. **Long-term Focus** - Quality compounds over decades

### What to Look For

Ideal investment candidates typically have:
- Quality Score > 70
- Consistent ROE above 15%
- Debt/Equity < 0.5
- Positive free cash flow
- Trading below intrinsic value

## Limitations

- Data quality depends on Yahoo Finance
- Limited historical data for newer companies
- Financial sector may need adjusted metrics
- Cyclical companies show point-in-time snapshots
- Backtesting has survivorship bias (only existing stocks)

## Contributing

Contributions are welcome! Please feel free to submit issues or pull requests.

## Disclaimer

This tool is for educational and research purposes only. It does not constitute financial advice. Always conduct your own due diligence before making investment decisions. Past performance does not guarantee future results.

## License

MIT License - see LICENSE file for details.

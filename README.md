# Montgomery Value Investment Scorer

A Python-based stock analysis tool implementing Roger Montgomery's value investing methodology from his book "Value.able". This tool evaluates companies based on quality metrics with heavy emphasis on Return on Equity (ROE), debt levels, and free cash flow generation.

## Overview

Roger Montgomery's approach focuses on finding high-quality businesses that can compound wealth over long periods. His methodology prioritizes:
- **High and consistent ROE** (Return on Equity)
- **Low or no debt** 
- **Strong free cash flow generation**
- **Reasonable valuations relative to quality**

## Scoring Components

The scoring system allocates 100 points across five key areas:

### 1. Return on Equity (30 points)
- Montgomery's primary quality metric
- Measures management's ability to generate returns on shareholder equity
- Scoring:
  - ≥20% ROE: 30 points
  - ≥15% ROE: 25 points
  - ≥10% ROE: 15 points
  - Bonus points for consistency (low standard deviation)

### 2. Debt Levels (25 points)
- Montgomery strongly prefers companies with little to no debt
- Debt-free companies have more flexibility during downturns
- Scoring based on Debt-to-Equity ratio:
  - 0 (debt-free): 25 points
  - <0.1: 23 points
  - <0.3: 20 points
  - <0.5: 15 points

### 3. Free Cash Flow (20 points)
- Measures actual cash generation after capital expenditures
- FCF Yield = Free Cash Flow / Market Cap
- Scoring:
  - ≥10% FCF Yield: 20 points
  - ≥7% FCF Yield: 16 points
  - ≥5% FCF Yield: 12 points

### 4. Valuation (15 points)
- P/E ratio component
- P/B ratio relative to ROE (Montgomery's value approach)
- Lower multiples for quality companies indicate better value

### 5. Profitability (10 points)
- Net profit margin
- Operating margin
- Higher margins indicate pricing power and efficiency

## Quality Ratings

Based on the total score:
- **A+ (80-100)**: Exceptional quality, strong moat
- **A (70-79)**: High quality, competitive advantages
- **B+ (60-69)**: Good quality, above average
- **B (50-59)**: Average quality, acceptable
- **C+ (40-49)**: Below average, some concerns
- **C (30-39)**: Poor quality, significant issues
- **D (0-29)**: Very poor quality, avoid

## Installation

```bash
# Install required packages
pip install yfinance pandas openpyxl numpy --break-system-packages
```

## Usage

### Basic Usage
```bash
# Create a text file with tickers (one per line)
echo "AAPL
MSFT
BRK-B
JNJ" > my_stocks.txt

# Run the analysis
python montgomery_value_scorer.py my_stocks.txt
```

### Custom Output File
```bash
python montgomery_value_scorer.py my_stocks.txt -o my_analysis.xlsx
```

### Input File Format
Create a text file with one ticker symbol per line:
```
AAPL
MSFT
GOOGL
BRK-B
JNJ
```

## Output

The tool generates an Excel file with:

### Sheet 1: Value Investment Analysis
- All financial metrics for each stock
- Montgomery Score (0-100)
- Quality Rating (A+ to D)
- Color-coded cells for easy visual analysis:
  - Green: Excellent metrics
  - Yellow: Average metrics
  - Red: Poor metrics

### Sheet 2: Scoring Methodology
- Detailed explanation of scoring components
- Quality rating definitions
- Key principles

## Key Metrics Analyzed

- **ROE (%)**: Return on Equity
- **5Y Avg ROE (%)**: 5-year average for consistency
- **Debt to Equity**: Financial leverage
- **Free Cash Flow**: Actual cash generation
- **Net Margin (%)**: Profitability
- **Operating Margin (%)**: Operational efficiency
- **P/E Ratio**: Price to Earnings
- **P/B Ratio**: Price to Book
- **PEG Ratio**: Price/Earnings to Growth
- **Dividend Yield**: Income component
- **Payout Ratio**: Dividend sustainability

## Investment Philosophy

Montgomery's approach aligns with Warren Buffett's principles:
1. **Quality over everything**: High ROE indicates competitive advantages
2. **Margin of Safety**: Buy quality companies at reasonable prices
3. **Long-term focus**: Hold for 10-20+ years
4. **Avoid debt-heavy companies**: Debt amplifies risk
5. **Cash is king**: Free cash flow is the ultimate measure of value

## Ideal Candidates for Long-term Investment

Look for companies with:
- Montgomery Score > 70
- Consistent ROE above 15%
- Debt/Equity < 0.5
- Strong free cash flow generation
- Reasonable valuations (P/E < 20 for quality companies)
- Sustainable competitive advantages (moats)

## Limitations

- Data quality depends on Yahoo Finance availability
- Historical data may be limited for newer companies
- Financial companies may need adjusted metrics
- Cyclical companies may show misleading point-in-time metrics

## Examples of High-Quality Companies

Historically high-scoring companies often include:
- Consumer staples with strong brands (P&G, Coca-Cola)
- Technology leaders with network effects (Microsoft, Apple)
- Healthcare companies with patents (Johnson & Johnson)
- Financial services with scale advantages (Berkshire Hathaway)

## Further Reading

- "Value.able" by Roger Montgomery
- "The Intelligent Investor" by Benjamin Graham
- "Common Stocks and Uncommon Profits" by Philip Fisher
- Montgomery Investment Management website

## Disclaimer

This tool is for educational and research purposes only. Always conduct your own due diligence before making investment decisions. Past performance does not guarantee future results.
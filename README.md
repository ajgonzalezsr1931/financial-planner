# Financial Planner CC - Credit Card Payment Planner

A comprehensive Google Apps Script solution for planning and analyzing credit card payments and payoff strategies.

## 🎯 Project Overview

This project provides a complete credit card payment planning system that works within Google Sheets, offering powerful calculations and scenario analysis to help users make informed decisions about debt payoff strategies.

## 📁 Project Structure

```
financial planner CC/
├── README.md                               # This file
├── credit-card-planner.gs                  # Main Google Apps Script code
├── sidebar.html                            # Interactive sidebar interface  
├── calculator.html                         # Quick payment calculator dialog
└── CREDIT_CARD_PLANNER_INSTRUCTIONS.md    # Complete setup and usage guide
```

## 🚀 Quick Start

1. Open [Google Sheets](https://sheets.google.com) and create a new spreadsheet
2. Go to **Extensions** → **Apps Script**
3. Copy the code from `credit-card-planner.gs` into the script editor
4. Add the HTML files (`sidebar.html` and `calculator.html`) as HTML files in the project
5. Save, deploy, and grant permissions
6. Return to Google Sheets - you'll see a new "Credit Card Planner" menu

## ✨ Key Features

### Core Functionality
- **Multiple Credit Card Tracking**: Manage multiple cards simultaneously
- **Payment Calculation**: Calculate required payments for target payoff dates
- **Month-by-Month Timeline**: Detailed breakdown of each payment
- **Scenario Comparison**: Compare different payment strategies
- **Interest Analysis**: See exactly how much interest you'll pay

### Bill Tracking System
- **📋 Monthly Bills**: Track fixed monthly bills
- **📊 Variable Monthly Bills**: Handle bills with fixed due dates but varying amounts (electric, gas, water)
- **🔄 28-Day Cycle Bills**: Manage bills with 28-day cycles
- **⏳ Limited Duration Bills**: Track loan payments with progress tracking
- **💰 Bi-Weekly Income**: Calculate monthly estimates from bi-weekly pay
- **📅 Bill Schedule**: 6-month calendar view of upcoming bills
- **📊 Budget Summary**: Net cash flow analysis

### Electric Bill Analysis Tool
- **PDF Analysis**: Extract current charges from electric bill PDFs
- **Seasonal Patterns**: Identify summer vs winter usage variations  
- **Accurate Budgeting**: Calculate estimates based on actual usage (excludes previous balances)
- **Historical Analysis**: Analyze years of billing data for trends

#### To Update Electric Bill Analysis:
1. Add new PDF bills to `bills/chariot/` folder
2. Run: `python3 extract_bills.py`  
3. Review updated `billing_summary.txt` for new estimates
4. Update Variable Monthly Bill entries as needed

### User Interface
- **Interactive Sidebar**: Easy card management within Google Sheets
- **Quick Calculator**: Pop-up calculator for fast calculations
- **Automated Sheet Creation**: Organized results in separate sheets
- **Professional Styling**: Clean, user-friendly interface

### Advanced Analysis
- **Debt Avalanche Planning**: Optimize payment order for multiple cards
- **What-If Scenarios**: Test different payment amounts
- **Total Cost Analysis**: Understand the true cost of carrying debt
- **Savings Calculator**: See money saved with higher payments

## 🔧 Technical Details

- **Platform**: Google Apps Script (JavaScript V8)
- **Deployment**: Google Sheets Add-on
- **Authentication**: Google OAuth2
- **Storage**: Google Sheets (no external databases)
- **UI Framework**: Google Apps Script HtmlService

## 📊 File Descriptions

### `credit-card-planner.gs`
Main script containing:
- Credit card calculation algorithms
- Payment timeline generation
- Scenario comparison logic
- Google Sheets integration
- Menu and dialog management

### `sidebar.html`  
Interactive sidebar interface featuring:
- Credit card input forms
- Existing card management
- Analysis tool controls
- Real-time calculation results
- Data management options

### `calculator.html`
Quick calculator dialog providing:
- Payment amount calculations
- Payoff timeline generation
- Dual-mode interface (payment vs timeline)
- Instant results display
- Input validation

### `CREDIT_CARD_PLANNER_INSTRUCTIONS.md`
Comprehensive documentation including:
- Step-by-step installation guide
- Complete usage instructions
- Feature explanations
- Troubleshooting tips
- Advanced customization options

## 🎯 Use Cases

- **Personal Finance Planning**: Individual debt payoff strategies
- **Financial Counseling**: Tools for financial advisors
- **Budgeting**: Understanding debt payment impacts
- **Education**: Learning about interest and debt mechanics
- **Business**: Employee financial wellness programs

## ✅ Recently Completed Features (March 2026)

### Interactive Payment Validation System
- **Monthly Payment Tracking**: Added columns I-T (Month 1-12) to Credit Cards sheet for actual payment tracking
- **Validation Checkboxes**: Interactive timeline with checkboxes in column L to validate when payments are made
- **Automatic Updates**: Checkbox clicks automatically update corresponding monthly columns in Credit Cards sheet
- **Card-Specific Updates**: Precise targeting ensures only the correct card's payment is updated
- **Visual Feedback**: Validated payments highlighted with green formatting and user confirmation alerts

### Advanced Timeline Features  
- **Mixed Payment Support**: Users can customize payment amounts in timeline column H
- **Dynamic Formulas**: Interest, principal, and balance calculations automatically adjust to payment changes
- **Final Payment Logic**: Last payments automatically calculate exact amount needed to pay off debt
- **Extension Handling**: Additional timeline rows appear automatically if balance remains after planned months
- **Progress Visualization**: Color-coded balance cells show payoff progress (25%, 50%, 75% milestones)

### Robust Trigger System
- **Simple Triggers**: Implemented Google Sheets' built-in `onEdit()` function for reliable validation
- **Automatic Detection**: System detects schedule sheets and validation column changes automatically  
- **Error Recovery**: Comprehensive debugging tools and manual validation options
- **Trigger Management**: Menu options for trigger setup, debugging, and cleanup

### Enhanced Debugging & Testing
- **Manual Testing**: `Test Validation Manually` function for troubleshooting
- **Simulation Tools**: `Simulate Checkbox Click` for testing without actual checkbox interaction
- **Trigger Debugging**: `Debug Triggers` function shows all active triggers and their status
- **Manual Validation Mode**: Backup instructions for manual payment validation when needed
- **Comprehensive Logging**: Detailed console logging for troubleshooting validation issues

### User Experience Improvements
- **Smart Card Name Extraction**: Handles custom schedule timestamps and sheet naming variations
- **Input Validation**: Prevents invalid month ranges and missing payment data
- **User Feedback**: Clear success/error messages for all validation operations
- **Emergency Functions**: `Clear All Triggers` option for troubleshooting trigger issues

## 💡 Future Work Items / Next Session Plans

### Dynamic Payment Schedule Integration (Planned - Next Session)
- **Actual Payment Reader**: Read validated payments from Credit Cards sheet columns I-T (Month 1-12)
- **Remaining Balance Calculator**: Calculate current debt after actual payments with interest
- **Smart Schedule Recalculation**: Adjust remaining payment schedule to meet target payoff dates
- **Enhanced Timeline Display**: Show actual vs planned payments with visual distinction
- **Adaptive Planning**: Handle ahead/behind schedule scenarios with optimized recommendations

### Additional Enhancement Ideas

Priority development tasks:

### 1. Payment Tracking System
- [ ] Add checkboxes to indicate that the stated payment was made
- [ ] Automatically update the remaining balance when payments are marked as completed
- [ ] Track actual payment dates vs. planned dates
- [ ] Visual indicators for on-time/late payments

### 2. Sheet 2 Rebuild
- [ ] Start rebuilding sheet 2 to be more intuitive
- [ ] Ensure seamless integration with CC Planner functionality
- [ ] Improve user experience and workflow between sheets
- [ ] Add cross-sheet data validation and consistency

### 3. Page 1 Redesign
- [ ] Rebuild page 1 to be cleaner and more user-friendly
- [ ] Make it easier to update card information
- [ ] Improve visual layout for better readability
- [ ] Streamline the data entry and follow-up process

Additional potential enhancements:
- Integration with banking APIs
- Advanced debt consolidation analysis
- Credit score impact modeling
- Export to other financial tools

## 📝 Version History

- **v1.0** (March 2026): Initial release with core functionality

---

**Created**: March 2026  
**Author**: Custom Google Apps Script Solution  
**Platform**: Google Workspace (Google Sheets)  
**License**: Personal Use
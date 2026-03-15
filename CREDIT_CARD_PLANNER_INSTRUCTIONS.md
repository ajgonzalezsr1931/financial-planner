# Credit Card Payment Planner for Google Sheets

A comprehensive Google Apps Script tool for planning credit card payoffs with multiple payment scenarios and detailed analysis.

## Features

- **Multiple Credit Cards**: Track and analyze multiple credit cards simultaneously
- **Payment Calculation**: Calculate required monthly payments for target payoff dates
- **Payoff Timeline**: Month-by-month breakdown of payments, interest, and remaining balance
- **Scenario Comparison**: Compare different payment amounts and their impact
- **Custom Analysis**: Generate schedules for any payment amount
- **Interactive Sidebar**: Easy-to-use interface within Google Sheets

## Installation

### Step 1: Create a New Google Sheets File
1. Go to [sheets.google.com](https://sheets.google.com)
2. Create a new blank spreadsheet
3. Give it a name like "Credit Card Payment Planner"

### Step 2: Open Apps Script Editor
1. In your Google Sheet, click **Extensions** > **Apps Script**
2. This will open the Google Apps Script editor in a new tab

### Step 3: Add the Script Files
1. **Replace Code.gs**: Delete the default code and paste the contents of `credit-card-planner.gs`
2. **Add HTML Files**:
   - Click the **+** button next to "Files"
   - Select **HTML** 
   - Name it `sidebar` and paste the contents of `sidebar.html`
   - Repeat for `calculator` naming it `calculator` and using `calculator.html` content

### Step 4: Save and Deploy
1. Click the **Save** button (disk icon)
2. Give your project a name like "Credit Card Planner"
3. Click **Deploy** > **New Deployment**
4. Choose **Execute as**: Me
5. Choose **Who has access**: Anyone
6. Click **Deploy**
7. Grant necessary permissions when prompted

**Important**: If you make any changes to the code files, you must create a **New Deployment** for the changes to take effect. Don't use "Manage Deployments" - always create a new one.

### Step 5: Return to Google Sheets
1. Go back to your Google Sheets tab
2. Refresh the page
3. You should see a new menu called **Credit Card Planner**

## Usage

### Adding Credit Cards

1. Click **Credit Card Planner** > **Open Payment Planner**
2. In the sidebar that appears:
   - Enter card name (e.g., "Chase Freedom", "Capital One")
   - Enter current balance owed
   - Enter APR (Annual Percentage Rate)
   - Enter current minimum payment
   - Enter target months to pay off
3. Click **Add Credit Card**

The tool will automatically:
- Calculate the required monthly payment
- Create a summary in the "Credit Cards" sheet
- Show total interest and total amount you'll pay

### Analyzing Payment Scenarios

#### Generate Payment Schedule
1. Select a card from the dropdown
2. Click **Payment Schedule**
3. A new sheet will be created showing month-by-month breakdown

#### Compare Multiple Scenarios
1. Select a card from the dropdown
2. Click **Compare Scenarios**
3. View different payment amounts and their impact on:
   - Total payoff time
   - Total interest paid
   - Money saved vs minimum payments

#### Custom Payment Analysis
1. Select a card
2. Enter a custom payment amount
3. Click **Generate Custom Schedule**
4. See exactly how long it will take with that payment

### Quick Calculator

Use **Credit Card Planner** > **Calculate Payment Amount** for quick calculations without saving data.

## Understanding the Results

### Key Metrics

- **Required Payment**: Monthly payment needed to pay off in target months
- **Total Interest**: Extra money paid beyond the original balance
- **Payoff Time**: How many months until the balance reaches $0
- **Interest as % of Balance**: Shows the "cost" of carrying the debt

### Reading Payment Schedules

Each month shows:
- **Payment**: Your monthly payment amount
- **Interest**: Portion going to interest charges
- **Principal**: Portion reducing your balance
- **Remaining Balance**: What you still owe

### Scenario Comparisons

Compare different strategies:
- **Minimum Payment**: Longest time, highest interest
- **Double Minimum**: Faster payoff, less interest
- **Fixed Timeframes**: 12, 24, 36 month payoffs
- **Custom Amounts**: Your preferred payment level

## Tips for Best Results

### Accurate Data Entry
- Use exact current balance (check latest statement)
- Enter the actual APR, not promotional rates
- Include any annual fees in your calculations

### Strategic Planning
- **Higher payments early** save the most interest
- **Target 24-36 months** for most cards is often optimal
- **Pay highest APR cards first** if you have multiple
- **Consider debt consolidation** if you have many high-APR cards

### Regular Updates
- Update balances monthly as they change
- Recalculate when APR changes
- Adjust if you miss or make extra payments

## Advanced Features

### Multiple Card Strategy
1. Add all your credit cards
2. Use the comparison tools to see which should be paid first
3. Consider "debt avalanche" (highest APR first) vs "debt snowball" (smallest balance first)

### What-If Analysis
- Try different payment amounts to see the impact
- Compare aggressive vs conservative payoff strategies
- Calculate break-even points for debt consolidation

### Export and Sharing
- All data stays in your Google Sheet
- Share with financial advisors or family
- Export to PDF for record keeping

## Troubleshooting

### Common Issues

**Menu doesn't appear**: Refresh the page and wait a few seconds

**Permission errors**: Make sure you granted all requested permissions during setup

**Calculations seem wrong**: Verify you entered APR as annual rate (not monthly)

**Can't save cards**: Check that all required fields are filled out

### Getting Help

The tool includes built-in validation and error messages. If you encounter issues:

1. Check the Browser Console (F12) for error details
2. Verify all inputs are positive numbers
3. Make sure APR is between 0-50%
4. Ensure target months is at least 1

## Privacy and Security

- All data stays in your Google Drive
- No external services or third parties involved
- You control who has access to your spreadsheet
- Google's standard security applies

## Customization

The tool is built with Google Apps Script - you can modify:
- Add new calculation methods
- Change the user interface
- Include additional fee calculations
- Export data to other formats

Feel free to modify the code to match your specific needs!

---

**Version**: 1.0  
**Created**: March 2026  
**Compatible with**: Google Sheets, Google Apps Script
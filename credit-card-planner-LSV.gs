/**
 * Credit Card Payment Planner for Google Sheets
 * A comprehensive tool for planning credit card payoffs with multiple scenarios
 */

// Main menu function to add to Google Sheets
function onInstall() {
  onOpen();
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Credit Card Planner')
    .addItem('Open Payment Planner', 'showSidebar')
    .addItem('Calculate Payment Amount', 'showPaymentCalculator')
    .addItem('Update Credit Cards Sheet', 'updateCreditCardsSheet')
    .addSeparator()
    .addItem('🏠 Generate Mortgage Amortization', 'showMortgageDialog')
    .addItem('🏠 Update Mortgage Status', 'updateMortgageStatus')
    .addItem('🏠 Generate Simulation Table', 'generateExtraPaymentSimulation')
    .addSeparator()
    .addItem('Setup Validation System', 'setupValidationTrigger')
    .addItem('Debug Triggers', 'debugTriggers')
    .addItem('Test Validation Manually', 'testValidationManually')
    .addItem('Simulate Checkbox Click', 'simulateCheckboxClick')
    .addItem('Manual Validation Mode', 'enableManualValidation')
    .addItem('Clear All Triggers', 'clearAllTriggers')
    .addItem('Clear All Data', 'clearAllData')
    .addToUi();
}

/**
 * Manual test function to test validation without triggers
 */
function testValidationManually() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const sheetName = sheet.getName();
    
    SpreadsheetApp.getUi().alert('Test Info', `Current sheet: ${sheetName}\nTesting validation on current sheet...`, SpreadsheetApp.getUi().ButtonSet.OK);
    
    // Test if this looks like a schedule sheet
    if (!sheetName.includes(' Schedule')) {
      SpreadsheetApp.getUi().alert('Error', `This doesn't appear to be a schedule sheet. Sheet name: "${sheetName}"\nPlease switch to a schedule sheet first.`, SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // Extract card name
    let cardName = sheetName.replace(' Schedule', '');
    cardName = cardName.replace(/\s*\(Custom\s+\d{2}:\d{2}\)/, '').trim();
    
    // Test with first row (month 1)
    const testRow = 9; // Row 9 should be first month
    const monthCell = sheet.getRange(testRow, 7); // Column G
    const paymentCell = sheet.getRange(testRow, 8); // Column H
    const month = monthCell.getValue();
    const paymentAmount = paymentCell.getValue();
    
    const testInfo = `Card Name: "${cardName}"\nMonth: ${month} (${typeof month})\nPayment Amount: ${paymentAmount} (${typeof paymentAmount})\n\nWill test updating Credit Cards sheet...`;
    
    const result = SpreadsheetApp.getUi().alert('Test Validation', testInfo, SpreadsheetApp.getUi().ButtonSet.OK_CANCEL);
    
    if (result === SpreadsheetApp.getUi().Button.OK) {
      updateCreditCardPayment(cardName, month, paymentAmount);
      SpreadsheetApp.getUi().alert('Test Complete', 'Manual validation test completed. Check the Credit Cards sheet and the Apps Script logs for details.', SpreadsheetApp.getUi().ButtonSet.OK);
    }
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Test Error', 'Error during manual test: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Simulate the onEdit trigger for testing
 */
function simulateCheckboxClick() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const selection = sheet.getActiveRange();
    
    if (!selection || selection.getColumn() !== 12) {
      SpreadsheetApp.getUi().alert('Instructions', 'Please select a validation checkbox cell (column L) in a schedule sheet first, then run this test.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // Create a mock event object
    const mockEvent = {
      range: selection,
      source: SpreadsheetApp.getActiveSpreadsheet()
    };
    
    const currentValue = selection.getValue();
    SpreadsheetApp.getUi().alert('Simulating Click', `Simulating checkbox click on ${selection.getA1Notation()}\nCurrent value: ${currentValue}\nSheet: ${sheet.getName()}`, SpreadsheetApp.getUi().ButtonSet.OK);
    
    // Call the validation function directly
    onValidationChange(mockEvent);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Simulation Error', 'Error during simulation: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Enable manual validation mode with instructions
 */
function enableManualValidation() {
  const instructions = `MANUAL VALIDATION MODE\n\n` +
    `If automatic validation isn't working, you can validate payments manually:\n\n` +
    `1. Go to any Schedule sheet\n` +
    `2. Select a checkbox cell in column L (validation column)\n` +
    `3. Go to Credit Card Planner → Simulate Checkbox Click\n\n` +
    `OR\n\n` +
    `1. Go to any Schedule sheet\n` +
    `2. Go to Credit Card Planner → Test Validation Manually\n\n` +
    `This will manually update the Credit Cards sheet with the payment information.`;
  
  SpreadsheetApp.getUi().alert('Manual Validation Instructions', instructions, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Debug function to see all triggers
 */
function debugTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  let message = `Total triggers: ${triggers.length}\n\n`;
  
  triggers.forEach((trigger, index) => {
    message += `Trigger ${index + 1}:\n`;
    message += `  Function: ${trigger.getHandlerFunction()}\n`;
    message += `  Event Type: ${trigger.getEventType()}\n`;
    message += `  Source: ${trigger.getTriggerSource()}\n`;
    message += `  Unique ID: ${trigger.getUniqueId()}\n\n`;
  });
  
  if (triggers.length === 0) {
    message += "No triggers found. You may need to run 'Setup Validation Trigger'.";
  }
  
  SpreadsheetApp.getUi().alert('Trigger Debug Info', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Clear all triggers (emergency function)
 */
function clearAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  let deletedCount = 0;
  
  triggers.forEach(trigger => {
    ScriptApp.deleteTrigger(trigger);
    deletedCount++;
  });
  
  SpreadsheetApp.getUi().alert('Triggers Cleared', `Deleted ${deletedCount} triggers. You will need to run 'Setup Validation Trigger' again to re-enable validation.`, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Setup validation system using Google Sheets' built-in trigger
 */
function setupValidationTrigger() {
  try {
    // Clear any old installable triggers first
    const triggers = ScriptApp.getProjectTriggers();
    let deletedCount = 0;
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'onValidationChange' || trigger.getHandlerFunction() === 'onEdit') {
        ScriptApp.deleteTrigger(trigger);
        deletedCount++;
      }
    });
    
    console.log(`Deleted ${deletedCount} existing triggers`);
    
    // Note: Google Sheets will automatically call onEdit() function when any cell is edited
    // We don't need to create an installable trigger for this
    
    SpreadsheetApp.getUi().alert('Success', `Validation system is now active! Google Sheets will automatically detect checkbox changes.\n\n(Cleaned up ${deletedCount} old triggers)\n\nThe onEdit() function will run automatically when you check validation checkboxes.`, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (error) {
    console.error('setupValidationTrigger error:', error);
    SpreadsheetApp.getUi().alert('Error', 'Setup error: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// Show the main sidebar
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('sidebar')
    .setTitle('Credit Card Payment Planner')
    .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}

// Show payment calculator dialog
function showPaymentCalculator() {
  const html = HtmlService.createHtmlOutputFromFile('calculator')
    .setWidth(400)
    .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'Payment Calculator');
}

// Show mortgage amortization dialog
function showMortgageDialog() {
  const ui = SpreadsheetApp.getUi();
  
  const result = ui.prompt(
    '🏠 Mortgage Amortization Table',
    'Enter your mortgage details (comma-separated):\n' +
    'Format: OriginDate, OriginalLoanAmount, InterestRate%, MaturityDate, CurrentPaymentDate\n\n' +
    'Example: 07/2021, 171615.18, 2.75, 07/2061, 03/2026',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (result.getSelectedButton() === ui.Button.OK) {
    const input = result.getResponseText().trim();
    try {
      const parts = input.split(',').map(p => p.trim());
      if (parts.length !== 5) {
        throw new Error('Please provide exactly 5 values separated by commas');
      }
      
      const [originDate, originalLoanAmount, interestRate, maturityDate, currentPaymentDate] = parts;
      
      generateMortgageAmortization(
        originDate,
        parseFloat(originalLoanAmount),
        parseFloat(interestRate),
        maturityDate,
        currentPaymentDate
      );
    } catch (error) {
      ui.alert('Error', 'Invalid input: ' + error.message, ui.ButtonSet.OK);
    }
  }
}

/**
 * Calculate required monthly payment to pay off debt in specified months
 */
function calculateMonthlyPayment(balance, apr, months) {
  if (balance <= 0 || months <= 0) return 0;
  
  const monthlyRate = apr / 100 / 12;
  
  if (monthlyRate === 0) {
    return balance / months;
  }
  
  const payment = balance * (monthlyRate * Math.pow(1 + monthlyRate, months)) / 
                  (Math.pow(1 + monthlyRate, months) - 1);
  
  return Math.ceil(payment * 100) / 100; // Round up to nearest cent
}

/**
 * Calculate payoff timeline with month-by-month breakdown
 */
function calculatePayoffTimeline(balance, apr, monthlyPayment) {
  const timeline = [];
  const monthlyRate = apr / 100 / 12;
  let currentBalance = balance;
  let month = 0;
  let totalInterestPaid = 0;
  
  while (currentBalance > 0 && month < 600) { // Safety limit of 50 years
    month++;
    
    const interestPayment = currentBalance * monthlyRate;
    const principalPayment = Math.min(monthlyPayment - interestPayment, currentBalance);
    
    if (principalPayment <= 0) {
      // Payment doesn't cover interest - will never pay off
      return {
        timeline: [],
        error: "Payment amount too low to cover interest charges",
        totalMonths: 0,
        totalInterest: 0
      };
    }
    
    currentBalance -= principalPayment;
    totalInterestPaid += interestPayment;
    
    // Use monthly payment for all months except adjust final payment if needed
    let actualPayment = monthlyPayment;
    
    // If this payment would result in balance <= 0, adjust to exact amount needed
    if (currentBalance <= 0.01) {
      actualPayment = interestPayment + principalPayment;
    }
    
    timeline.push({
      month: month,
      payment: Math.round(actualPayment * 100) / 100,
      interest: Math.round(interestPayment * 100) / 100,
      principal: Math.round(principalPayment * 100) / 100,
      balance: Math.round(currentBalance * 100) / 100
    });
    
    if (currentBalance < 0.01) break;
  }
  
  return {
    timeline: timeline,
    totalMonths: month,
    totalInterest: Math.round(totalInterestPaid * 100) / 100,
    totalPaid: Math.round((balance + totalInterestPaid) * 100) / 100
  };
}

/**
 * Update existing Credit Cards sheet to include monthly payment columns
 */
function updateCreditCardsSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Credit Cards');
  if (!sheet) {
    SpreadsheetApp.getUi().alert('No Credit Cards sheet found to update.');
    return;
  }
  
  // Check if monthly columns already exist
  const lastCol = sheet.getLastColumn();
  if (lastCol >= 20) { // 8 original + 12 monthly = 20 columns
    return; // Already has monthly columns
  }
  
  // Add monthly payment headers
  const monthlyHeaders = [];
  for (let i = 1; i <= 12; i++) {
    monthlyHeaders.push(`Month ${i}`);
  }
  
  // Add headers starting from column I (column 9)
  sheet.getRange(1, 9, 1, 12).setValues([monthlyHeaders]);
  sheet.getRange(1, 9, 1, 12).setFontWeight('bold');
  
  // Update existing card data with monthly payments
  if (sheet.getLastRow() > 1) {
    const existingData = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
    
    existingData.forEach((row, index) => {
      const targetMonths = row[4]; // Target Months (column E)
      const requiredPayment = row[5]; // Required Payment (column F)
      
      // Add monthly payment data for this card
      const monthlyPayments = [];
      for (let i = 1; i <= 12; i++) {
        if (i <= targetMonths) {
          monthlyPayments.push(requiredPayment);
        } else {
          monthlyPayments.push('');
        }
      }
      
      // Update the row with monthly payment data
      sheet.getRange(index + 2, 9, 1, 12).setValues([monthlyPayments]);
    });
  }
  
  // Auto-resize all columns
  sheet.autoResizeColumns(1, 20);
  
  SpreadsheetApp.getUi().alert('Credit Cards sheet updated with monthly payment columns!');
}

/**
 * Helper function to ensure Credit Cards sheet has correct structure
 */
function updateCreditCardsSheetStructure(sheet) {
  // Add headers if this is the first card or if headers are missing
  if (sheet.getLastRow() === 0) {
    const headers = ['Card Name', 'Current Balance', 'APR (%)', 'Minimum Payment', 'Target Months', 'Required Payment', 'Total Interest', 'Total Paid'];
    // Add monthly payment columns
    const monthlyHeaders = [];
    for (let i = 1; i <= 12; i++) {
      monthlyHeaders.push(`Month ${i}`);
    }
    const allHeaders = headers.concat(monthlyHeaders);
    sheet.getRange(1, 1, 1, allHeaders.length).setValues([allHeaders]);
    sheet.getRange(1, 1, 1, allHeaders.length).setFontWeight('bold');
  } else {
    // Check if monthly columns exist and add them if missing
    const lastCol = sheet.getLastColumn();
    if (lastCol < 20) { // 8 original + 12 monthly = 20 columns
      const monthlyHeaders = [];
      for (let i = 1; i <= 12; i++) {
        monthlyHeaders.push(`Month ${i}`);
      }
      sheet.getRange(1, 9, 1, 12).setValues([monthlyHeaders]);
      sheet.getRange(1, 9, 1, 12).setFontWeight('bold');
    }
  }
}

/**
 * Add a new credit card to the spreadsheet
 */
function addCreditCard(cardData) {
  const sheet = getOrCreateSheet('Credit Cards');
  
  // Always ensure the sheet has the correct structure
  updateCreditCardsSheetStructure(sheet);
  
  const requiredPayment = calculateMonthlyPayment(cardData.balance, cardData.apr, cardData.targetMonths);
  const timeline = calculatePayoffTimeline(cardData.balance, cardData.apr, requiredPayment);
  
  const row = [
    cardData.name,
    cardData.balance,
    cardData.apr,
    cardData.minPayment,
    cardData.targetMonths,
    requiredPayment,
    timeline.totalInterest,
    timeline.totalPaid
  ];
  
  // Add monthly payment columns (12 months)
  const monthlyPayments = [];
  for (let i = 1; i <= 12; i++) {
    if (i <= cardData.targetMonths) {
      monthlyPayments.push(requiredPayment);
    } else {
      monthlyPayments.push('');
    }
  }
  const fullRow = row.concat(monthlyPayments);
  
  sheet.appendRow(fullRow);
  
  // Auto-resize columns including monthly payment columns
  const totalColumns = 8 + 12; // 8 original columns + 12 monthly payment columns
  sheet.autoResizeColumns(1, totalColumns);
  
  return {
    success: true,
    requiredPayment: requiredPayment,
    totalInterest: timeline.totalInterest,
    totalPaid: timeline.totalPaid
  };
}

/**
 * Generate detailed payoff schedule for a specific card
 * TODO: Next session - integrate actual payments from Credit Cards sheet
 * - Read validated payments from columns I-T
 * - Calculate remaining balance after actual payments
 * - Adjust payment schedule for remaining months to meet target
 */
function generatePayoffSchedule(cardName, balance, apr, monthlyPayment, minPayment) {
  const timeline = calculatePayoffTimeline(balance, apr, monthlyPayment);
  
  if (timeline.error) {
    throw new Error(timeline.error);
  }
  
  const sheetName = `${cardName} Schedule`.substring(0, 30); // Google Sheets name limit
  const sheet = getOrCreateSheet(sheetName);
  
  // Clear existing data
  sheet.clear();
  
  // Add summary info with vibrant styling
  const summaryRange = sheet.getRange(1, 1, 5, 2);
  summaryRange.setValues([
    ['🏦 Card Name:', cardName],
    ['💰 Starting Balance:', balance],
    ['📈 APR:', apr],
    ['💳 Minimum Payment:', minPayment],
    ['⏱️ Payoff Time:', `${timeline.totalMonths} months`]
  ]);
  
  // Format the numeric values properly
  sheet.getRange(2, 2).setNumberFormat('"$"#,##0.00'); // Starting Balance with 2 decimals
  sheet.getRange(3, 2).setNumberFormat('0.00"%"'); // APR
  sheet.getRange(4, 2).setNumberFormat('"$"#,##0.00'); // Minimum Payment
  
  // Style summary section
  sheet.getRange(1, 1, 5, 1).setBackground('#4ECDC4')  // Energetic teal
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');
  sheet.getRange(1, 2, 5, 1).setBackground('#45B7D1')  // Energetic blue
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');
  
  // Add timeline data starting from row 8
  const startRow = 8;
  const headers = ['Estimated Payoff', 'Payment', 'Interest', 'Principal', 'Remaining Balance'];
  const timelineHeaderRange = sheet.getRange(startRow, 1, 1, headers.length);
  timelineHeaderRange.setValues([headers]);
  timelineHeaderRange.setBackground('#E74C3C')  // Energetic red
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setFontSize(11)
    .setBorder(true, true, true, true, true, true, '#C0392B', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  
  const timelineData = timeline.timeline.map(row => [
    row.month,
    row.payment,
    row.interest,
    row.principal,
    row.balance
  ]);
  
  if (timelineData.length > 0) {
    const dataRange = sheet.getRange(startRow + 1, 1, timelineData.length, headers.length);
    dataRange.setValues(timelineData);
    
    // Apply alternating row colors for energetic look
    for (let i = 0; i < timelineData.length; i++) {
      const rowRange = sheet.getRange(startRow + 1 + i, 1, 1, headers.length);
      if (i % 2 === 0) {
        rowRange.setBackground('#FFF3E0');  // Light orange
      } else {
        rowRange.setBackground('#E8F5E8');  // Light green
      }
      
      // Add borders for definition
      rowRange.setBorder(false, false, true, false, false, false, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);
      
      // Highlight milestone months (every 6 months) with bold formatting
      if (timelineData[i][0] % 6 === 0) {
        rowRange.setFontWeight('bold')
          .setBackground('#FFE082');  // Golden yellow for milestones
      }
      
      // Color-code remaining balance for motivation
      const remainingBalance = timelineData[i][4];
      const originalBalance = balance;
      const percentComplete = ((originalBalance - remainingBalance) / originalBalance) * 100;
      
      if (percentComplete >= 75) {
        sheet.getRange(startRow + 1 + i, 5, 1, 1).setBackground('#66BB6A').setFontColor('#FFFFFF');  // Green - almost done!
      } else if (percentComplete >= 50) {
        sheet.getRange(startRow + 1 + i, 5, 1, 1).setBackground('#FFA726').setFontColor('#FFFFFF');  // Orange - halfway there!
      } else if (percentComplete >= 25) {
        sheet.getRange(startRow + 1 + i, 5, 1, 1).setBackground('#FFCA28');  // Yellow - making progress!
      }
    }
  }
  
  // Format currency columns with energetic styling
  const currencyColumns = [2, 3, 4, 5]; // Payment, Interest, Principal, Balance
  currencyColumns.forEach(col => {
    const currencyRange = sheet.getRange(startRow + 1, col, timelineData.length, 1);
    currencyRange.setNumberFormat('"$"#,##0.00')
      .setHorizontalAlignment('right');
  });
  
  // Add motivational final row if debt is paid off
  if (timeline.totalMonths < timelineData.length || timelineData[timelineData.length - 1][4] === 0) {
    const finalRow = startRow + 1 + timelineData.length;
    sheet.getRange(finalRow, 1, 1, 5).setValues([['', 'DEBT FREE!', '🎈', 'CONGRATULATIONS!', '🎊']]);
    sheet.getRange(finalRow, 1, 1, 5)
      .setBackground('#4CAF50')  // Victory green
      .setFontColor('#FFFFFF')
      .setFontWeight('bold')
      .setFontSize(14)
      .setHorizontalAlignment('center');
  }
  
  // Add interactive timeline starting at G8 with interconnected formulas
  addInteractiveTimeline(sheet, balance, apr, monthlyPayment, timeline.totalMonths, false, cardName); // false for regular schedule
  
  sheet.autoResizeColumns(1, headers.length);
  sheet.autoResizeColumns(7, 5); // Auto-resize timeline columns G-K
  
  return timeline;
}

/**
 * Add interactive timeline with interconnected formulas starting at G8
 */
function addInteractiveTimeline(sheet, originalBalance, apr, defaultPayment, maxMonths, isCustomSchedule = false, cardName = '') {
  const timelineStartCol = 7; // Column G
  const timelineStartRow = 8;
  
  // Timeline headers in G8:L8 - different label for custom schedules
  const timelineLabel = isCustomSchedule ? 'Actual Custom' : 'Actual Payoff';
  const timelineHeaders = [timelineLabel, 'Payment', 'Interest', 'Principal', 'Remaining Balance', 'Paid ✓'];
  const headerRange = sheet.getRange(timelineStartRow, timelineStartCol, 1, 6);
  headerRange.setValues([timelineHeaders]);
  
  // Different colors for custom schedule timeline
  const headerColor = isCustomSchedule ? '#8E24AA' : '#FF6B35'; // Purple for custom, orange for regular
  const borderColor = isCustomSchedule ? '#7B1FA2' : '#E55B2B';
  
  headerRange.setBackground(headerColor)
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setFontSize(11)
    .setBorder(true, true, true, true, true, true, borderColor, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  
  // Starting values and formulas from row 9
  const dataStartRow = timelineStartRow + 1; // Row 9
  
  // Match the exact number of months from the estimated payoff schedule
  const monthsToAdd = maxMonths;
  
  for (let month = 1; month <= monthsToAdd; month++) {
    const currentRow = dataStartRow + month - 1;
    
    // Column G: Month number
    sheet.getRange(currentRow, timelineStartCol).setValue(month);
    
    // Column H: Payment (automatic adjustment for final payments)
    const paymentCell = sheet.getRange(currentRow, timelineStartCol + 1);
    let paymentFormula;
    if (month === 1) {
      // First month: Use default payment (user can manually change this to $75)
      paymentFormula = `=${defaultPayment}`;
    } else {
      // Subsequent months: Use full default payment, but adjust for final payment
      const prevBalanceCell = sheet.getRange(currentRow - 1, timelineStartCol + 4).getA1Notation();
      // If previous balance + interest is less than default payment, use exact amount needed
      paymentFormula = `=IF(${prevBalanceCell}<=0,0,IF(${prevBalanceCell}+(${prevBalanceCell}*$B$3/100/12)<=${defaultPayment},${prevBalanceCell}+(${prevBalanceCell}*$B$3/100/12),${defaultPayment}))`;
    }
    paymentCell.setFormula(paymentFormula);
    paymentCell.setBackground('#E8F5E8')  // Light green to indicate it's dynamic
      .setBorder(true, true, true, true, false, false, '#4CAF50', SpreadsheetApp.BorderStyle.SOLID)
      .setNumberFormat('"$"#,##0.00');
    
    // Column I: Interest calculation formula
    let interestFormula;
    if (month === 1) {
      // First month: Interest = Starting Balance * Monthly Rate
      interestFormula = `=IF(B2<=0,0,B2*$B$3/100/12)`;
    } else {
      // Subsequent months: Interest = Previous Balance * Monthly Rate
      const prevBalanceCell = sheet.getRange(currentRow - 1, timelineStartCol + 4).getA1Notation();
      interestFormula = `=IF(${prevBalanceCell}<=0,0,${prevBalanceCell}*$B$3/100/12)`;
    }
    const interestCell = sheet.getRange(currentRow, timelineStartCol + 2);
    interestCell.setFormula(interestFormula);
    interestCell.setNumberFormat('"$"#,##0.00');
    
    // Column J: Principal calculation formula
    const paymentCellRef = sheet.getRange(currentRow, timelineStartCol + 1).getA1Notation();
    const interestCellRef = interestCell.getA1Notation();
    let principalFormula;
    if (month === 1) {
      // First month: Principal = MIN(Payment - Interest, Starting Balance)
      principalFormula = `=MIN(MAX(0,${paymentCellRef}-${interestCellRef}),B2)`;
    } else {
      // Subsequent months: Principal = MIN(Payment - Interest, Previous Remaining Balance)
      const prevBalanceCell = sheet.getRange(currentRow - 1, timelineStartCol + 4).getA1Notation();
      principalFormula = `=MIN(MAX(0,${paymentCellRef}-${interestCellRef}),${prevBalanceCell})`;
    }
    const principalCell = sheet.getRange(currentRow, timelineStartCol + 3);
    principalCell.setFormula(principalFormula);
    principalCell.setNumberFormat('"$"#,##0.00');
    
    // Column K: Remaining balance calculation formula
    let balanceFormula;
    if (month === 1) {
      // First month: New Balance = Starting Balance - Principal Payment
      const principalCellRef = principalCell.getA1Notation();
      balanceFormula = `=MAX(0,B2-${principalCellRef})`;
    } else {
      // Subsequent months: New Balance = Previous Balance - Principal Payment
      const prevBalanceCell = sheet.getRange(currentRow - 1, timelineStartCol + 4).getA1Notation();
      const principalCellRef = principalCell.getA1Notation();
      balanceFormula = `=MAX(0,${prevBalanceCell}-${principalCellRef})`;
    }
    const balanceCell = sheet.getRange(currentRow, timelineStartCol + 4);
    balanceCell.setFormula(balanceFormula);
    balanceCell.setNumberFormat('"$"#,##0.00');
    
    // Column L: Validation checkbox
    const validationCell = sheet.getRange(currentRow, timelineStartCol + 5);
    validationCell.insertCheckboxes()
      .setBackground('#FFF8DC')  // Light yellow background
      .setBorder(true, true, true, true, false, false, '#FFD700', SpreadsheetApp.BorderStyle.SOLID);
    
    // Apply alternating row colors
    const rowRange = sheet.getRange(currentRow, timelineStartCol, 1, 6);
    if (month % 2 === 0) {
      rowRange.setBackground('#FFF3E0');  // Light orange
    } else {
      rowRange.setBackground('#F3E5F5');  // Light purple
    }
    
    // Override validation cell background
    validationCell.setBackground('#FFF8DC');  // Keep validation cell light yellow
    
    // Add borders for definition
    rowRange.setBorder(false, false, true, false, false, false, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);
    
    // Highlight milestone months (every 6 months) with bold formatting
    if (month % 6 === 0) {
      rowRange.setFontWeight('bold')
        .setBackground('#FFE082');  // Golden yellow for milestones
    }
  }
  
  // Always add dynamic row handling to extend timeline if needed
  // The function will determine if additional rows are actually needed
  addDynamicBalanceHandling(sheet, timelineStartCol, dataStartRow, monthsToAdd, originalBalance, apr, defaultPayment);
  
  // Add instructions in a note with troubleshooting info
  const scheduleType = isCustomSchedule ? "CUSTOM" : "REGULAR";
  const instructionText = 
    `INTERACTIVE TIMELINE INSTRUCTIONS (${scheduleType} SCHEDULE):\\n\\n` +
    "• MIXED PAYMENTS SUPPORTED: You can manually change payment amounts in column H\\n" +
    "• Example: Change H9 from $" + defaultPayment.toFixed(2) + " to $75 for different first payment\\n" +
    "• All other columns (Interest, Principal, Balance) will recalculate automatically\\n" +
    "• Final payments automatically adjust to exact amount needed\\n" +
    "• Monthly interest rate: " + (apr/12).toFixed(4) + "%\\n" +
    "• Balance reference: B2 (" + originalBalance + ")\\n" +
    "• APR reference: B3 (" + apr + "%)\\n" +
    "• Default payment: $" + defaultPayment.toFixed(2) + " (but can be customized)\\n" +
    (isCustomSchedule ? "• This is a CUSTOM schedule - original schedules remain unchanged\\n" : "") +
    "• TO CREATE MIXED SCHEDULE: Edit payment amounts in column H as needed\\n\\n" +
    "TROUBLESHOOTING: If calculations seem wrong, verify B2 contains balance and B3 contains APR percentage.";
  
  sheet.getRange(timelineStartRow, timelineStartCol).setNote(instructionText);
  
  // Add conditional formatting to highlight zero balances
  const balanceColumnRange = sheet.getRange(dataStartRow, timelineStartCol + 4, monthsToAdd, 1);
  const conditionalFormatRule = SpreadsheetApp.newConditionalFormatRule()
    .setRanges([balanceColumnRange])
    .whenNumberLessThanOrEqualTo(0.01)
    .setBackground('#4CAF50')  // Green background for paid off
    .setFontColor('#FFFFFF')   // White text
    .build();
  const rules = sheet.getConditionalFormatRules();
  rules.push(conditionalFormatRule);
  sheet.setConditionalFormatRules(rules);
  
  // Add a helper note for mixed payment schedules
  const firstPaymentCell = sheet.getRange(dataStartRow, timelineStartCol + 1);
  firstPaymentCell.setNote("MIXED PAYMENTS: You can edit this cell to set a different first payment amount (e.g., $75). All other calculations will update automatically.");
  
  // Add validation note
  const firstValidationCell = sheet.getRange(dataStartRow, timelineStartCol + 5);
  firstValidationCell.setNote("PAYMENT VALIDATION: Check this box when you make the payment. This will update the Credit Cards sheet to track your actual payments.");
}

/**
 * Add dynamic row handling to extend timeline if balance remains
 */
function addDynamicBalanceHandling(sheet, timelineStartCol, dataStartRow, initialMonths, originalBalance, apr, defaultPayment) {
  // Guard against invalid initial months
  if (initialMonths <= 0) {
    return; // Nothing to handle if no initial months
  }
  
  // Add exactly ONE extension row that only shows if there's a remaining balance
  const currentRow = dataStartRow + initialMonths;
  const month = initialMonths + 1;
  const prevBalanceCell = sheet.getRange(currentRow - 1, timelineStartCol + 4).getA1Notation();
  
  // Column G: Month number (conditional) - show month only if balance remains
  const monthCell = sheet.getRange(currentRow, timelineStartCol);
  monthCell.setFormula(`=IF(${prevBalanceCell}>0.01,${month},"")`);
  
  // Column H: Payment - only show if balance remains
  const paymentCell = sheet.getRange(currentRow, timelineStartCol + 1);
  paymentCell.setFormula(`=IF(${prevBalanceCell}>0.01,${prevBalanceCell}+(${prevBalanceCell}*$B$3/100/12),"")`);
  paymentCell.setBackground('#E8FFE8')  // Light green to indicate final payment
    .setBorder(true, true, true, true, false, false, '#4CAF50', SpreadsheetApp.BorderStyle.SOLID_THICK);
  
  // Column I: Interest calculation (conditional)
  const interestCell = sheet.getRange(currentRow, timelineStartCol + 2);
  interestCell.setFormula(`=IF(${prevBalanceCell}>0.01,${prevBalanceCell}*$B$3/100/12,"")`);
  
  // Column J: Principal calculation (conditional)
  const paymentCellRef = paymentCell.getA1Notation();
  const interestCellRef = interestCell.getA1Notation();
  const principalCell = sheet.getRange(currentRow, timelineStartCol + 3);
  principalCell.setFormula(`=IF(AND(${prevBalanceCell}>0.01,ISNUMBER(${paymentCellRef})),MIN(${paymentCellRef}-${interestCellRef},${prevBalanceCell}),"")`);
  
  // Column K: Remaining balance (should be 0 for final payment)
  const principalCellRef = principalCell.getA1Notation();
  const balanceCell = sheet.getRange(currentRow, timelineStartCol + 4);
  balanceCell.setFormula(`=IF(AND(${prevBalanceCell}>0.01,ISNUMBER(${principalCellRef})),MAX(0,${prevBalanceCell}-${principalCellRef}),"")`);
  
  // Apply conditional formatting to currency columns - only format if there's a value
  [paymentCell, interestCell, principalCell, balanceCell].forEach(cell => {
    // Use a custom format that only applies currency formatting when there's a numeric value
    cell.setNumberFormat('[>0]"$"#,##0.00;[<0]"$"#,##0.00;""');
  });
  
  // Add validation checkbox for extension row
  const extensionValidationCell = sheet.getRange(currentRow, timelineStartCol + 5);
  extensionValidationCell.setFormula(`=IF(${prevBalanceCell}>0.01,FALSE,"")`);
  extensionValidationCell.insertCheckboxes()
    .setBackground('#FFF8DC');  // Light yellow background
  
  // Apply final payment styling
  const rowRange = sheet.getRange(currentRow, timelineStartCol, 1, 6);
  rowRange.setBackground('#E8F5E8');  // Light green background for extension row
  rowRange.setBorder(false, false, true, false, false, false, '#4CAF50', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  rowRange.setFontWeight('bold'); // Make extension row bold
  
  // Override validation cell background
  extensionValidationCell.setBackground('#FFF8DC');  // Keep validation cell light yellow
}

/**
 * Compare multiple payment scenarios
 */
function compareScenarios(cardName, balance, apr, scenarios) {
  const sheetName = `${cardName} Scenarios`.substring(0, 30);
  const sheet = getOrCreateSheet(sheetName);
  sheet.clear();
  
  // Headers with energetic styling
  const headers = ['Scenario', 'Monthly Payment', 'Payoff Time (Months)', 'Total Interest', 'Total Paid', 'Monthly Savings vs Min'];
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#9C27B0')  // Energetic purple
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setFontSize(12)
    .setBorder(true, true, true, true, true, true, '#7B1FA2', SpreadsheetApp.BorderStyle.SOLID_THICK);
  
  const results = [];
  let minPaymentResult = null;
  
  scenarios.forEach((scenario, index) => {
    const timeline = calculatePayoffTimeline(balance, apr, scenario.payment);
    
    if (!timeline.error) {
      const result = {
        name: scenario.name,
        payment: scenario.payment,
        months: timeline.totalMonths,
        interest: timeline.totalInterest,
        totalPaid: timeline.totalPaid
      };
      
      results.push(result);
      
      if (scenario.name.toLowerCase().includes('minimum') || index === 0) {
        minPaymentResult = result;
      }
    }
  });
  
  // Calculate savings vs minimum payment
  const scenarioData = results.map(result => {
    const savings = minPaymentResult ? minPaymentResult.totalPaid - result.totalPaid : 0;
    return [
      result.name,
      result.payment,
      result.months,
      result.interest,
      result.totalPaid,
      Math.round(savings)
    ];
  });
  
  if (scenarioData.length > 0) {
    sheet.getRange(2, 1, scenarioData.length, headers.length).setValues(scenarioData);
    
    // Format currency columns
    const currencyColumns = [2, 4, 5, 6];
    currencyColumns.forEach(col => {
      sheet.getRange(2, col, scenarioData.length, 1).setNumberFormat('"$"#,##0.00');
    });
    
    // Find and highlight the row with the least total paid amount (column 5)
    if (scenarioData.length > 1) {
      let minTotalPaid = Infinity;
      let bestScenarioRowIndex = -1;
      
      results.forEach((result, index) => {
        if (result.totalPaid < minTotalPaid) {
          minTotalPaid = result.totalPaid;
          bestScenarioRowIndex = index;
        }
      });
      
      if (bestScenarioRowIndex >= 0) {
        // Highlight the best scenario row in green (row index + 2 because data starts at row 2)
        const rowToHighlight = bestScenarioRowIndex + 2;
        sheet.getRange(rowToHighlight, 1, 1, headers.length)
          .setBackground('#d4edda')  // Light green background
          .setFontWeight('bold');    // Make it bold for emphasis
      }
    }
  }
  
  sheet.autoResizeColumns(1, headers.length);
  return results;
}

/**
 * Helper function to get or create a sheet
 */
/**
 * Helper function to get or create a sheet with unique naming for custom schedules
 */
function getOrCreateCustomSheet(baseName, isCustomSchedule = false) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  if (!isCustomSchedule) {
    // For regular schedules, use existing logic
    let sheet = ss.getSheetByName(baseName);
    if (!sheet) {
      sheet = ss.insertSheet(baseName);
    }
    return sheet;
  }
  
  // For custom schedules, create unique names to avoid overwriting
  const originalSheetName = baseName;
  let existingSheet = ss.getSheetByName(originalSheetName);
  
  if (!existingSheet) {
    // No existing sheet, create the first one
    return ss.insertSheet(originalSheetName);
  }
  
  // Sheet exists, create a new one with timestamp
  const timestamp = new Date().toLocaleTimeString('en-US', { 
    hour12: false, 
    hour: '2-digit', 
    minute: '2-digit' 
  });
  const customSheetName = `${baseName} (Custom ${timestamp})`.substring(0, 30);
  
  return ss.insertSheet(customSheetName);
}

function getOrCreateSheet(name) {
  return getOrCreateCustomSheet(name, false);
}

/**
 * Clear all data from all sheets created by this tool
 */
function clearAllData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  
  const sheetsToDelete = sheets.filter(sheet => {
    const name = sheet.getName();
    return name === 'Credit Cards' || 
           name.includes('Schedule') || 
           name.includes('Scenarios') ||
           name.includes('Summary');
  });
  
  sheetsToDelete.forEach(sheet => {
    if (sheets.length > 1) { // Don't delete if it's the only sheet
      ss.deleteSheet(sheet);
    } else {
      sheet.clear();
    }
  });
  
  SpreadsheetApp.getUi().alert('All credit card planner data has been cleared.');
}

/**
 * Get list of existing credit cards
 */
function getCreditCards() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Credit Cards');
  if (!sheet || sheet.getLastRow() <= 1) {
    return [];
  }
  
  // Read the first 8 columns (original card data)
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
  return data.map(row => ({
    name: row[0],
    balance: row[1],
    apr: row[2],
    minPayment: row[3],
    targetMonths: row[4]
  }));
}

/**
 * TODO: Next session - Read actual validated payments from Credit Cards sheet
 * @param {string} cardName - Name of the credit card
 * @returns {Object} Object containing actual payments and calculated remaining balance
 */
function getActualPayments(cardName) {
  // PLACEHOLDER - To be implemented next session
  // Will read columns I-T (months 1-12) from Credit Cards sheet
  // Calculate remaining balance after actual payments
  // Return: { actualPayments: [], remainingBalance: number, monthsPaid: number }
  return {
    actualPayments: [],
    remainingBalance: null,
    monthsPaid: 0
  };
}

/**
 * TODO: Next session - Calculate adjusted payment schedule based on actual payments
 * @param {string} cardName - Name of the credit card  
 * @param {number} originalBalance - Starting balance
 * @param {number} apr - Annual percentage rate
 * @param {number} targetMonths - Target months for payoff
 * @returns {Object} Adjusted payment schedule and timeline
 */
function generateAdjustedPayoffSchedule(cardName, originalBalance, apr, targetMonths) {
  // PLACEHOLDER - To be implemented next session
  // Will integrate getActualPayments() results
  // Recalculate required payments for remaining months
  // Generate timeline showing actual vs planned payments
  return null;
}

/**
 * Calculate mortgage payment using standard amortization formula
 */
function calculateMortgagePayment(principal, annualRate, totalMonths) {
  if (annualRate === 0) {
    return principal / totalMonths;
  }
  
  const monthlyRate = annualRate / 100 / 12;
  const payment = principal * (monthlyRate * Math.pow(1 + monthlyRate, totalMonths)) / 
                  (Math.pow(1 + monthlyRate, totalMonths) - 1);
  
  return Math.round(payment * 100) / 100;
}

/**
 * Calculate original loan amount from current balance and payments made
 */
function calculateOriginalLoanAmount(currentBalance, annualRate, monthsPaid, totalMonths) {
  if (annualRate === 0) {
    return currentBalance + (currentBalance * monthsPaid / (totalMonths - monthsPaid));
  }
  
  const monthlyRate = annualRate / 100 / 12;
  const remainingMonths = totalMonths - monthsPaid;
  
  // Calculate what the monthly payment would be for the original loan
  // Working backwards from current balance and remaining payments
  const monthlyPayment = currentBalance * (monthlyRate * Math.pow(1 + monthlyRate, remainingMonths)) / 
                        (Math.pow(1 + monthlyRate, remainingMonths) - 1);
  
  // Calculate original principal from monthly payment and total term
  const originalPrincipal = monthlyPayment * (Math.pow(1 + monthlyRate, totalMonths) - 1) / 
                           (monthlyRate * Math.pow(1 + monthlyRate, totalMonths));
  
  return Math.round(originalPrincipal * 100) / 100;
}

/**
 * Generate complete mortgage amortization table
 */
function generateMortgageAmortization(originDate, originalLoanAmount, annualRate, maturityDate, currentPaymentDate) {
  try {
    // Parse dates
    const [startMonth, startYear] = originDate.split('/');
    const [endMonth, endYear] = maturityDate.split('/');
    const [currentMonth, currentYear] = currentPaymentDate.split('/');
    
    const startDateObj = new Date(parseInt(startYear), parseInt(startMonth) - 1, 1);
    const endDateObj = new Date(parseInt(endYear), parseInt(endMonth) - 1, 1);
    const currentDateObj = new Date(parseInt(currentYear), parseInt(currentMonth) - 1, 1);
    
    // Calculate months
    const totalMonths = (endDateObj.getFullYear() - startDateObj.getFullYear()) * 12 + 
                       (endDateObj.getMonth() - startDateObj.getMonth());
    const monthsPaid = (currentDateObj.getFullYear() - startDateObj.getFullYear()) * 12 + 
                      (currentDateObj.getMonth() - startDateObj.getMonth()) + 1;
    
    // Calculate monthly payment from original loan amount
    const monthlyRate = annualRate / 100 / 12;
    const monthlyPayment = calculateMortgagePayment(originalLoanAmount, annualRate, totalMonths);
    
    // Create amortization sheet
    const sheet = getOrCreateSheet('Mortgage Amortization');
    sheet.clear();
    
    // Header section with mortgage details
    const headerRange = sheet.getRange(1, 1, 8, 2);
    headerRange.setValues([
      ['🏠 Mortgage Loan Details', ''],
      ['Original Loan Amount:', originalLoanAmount],
      ['Current Balance (after ' + currentPaymentDate + '):', 0], // Will be updated after calculation
      ['Annual Interest Rate:', annualRate + '%'],
      ['Monthly Payment (P&I):', monthlyPayment],
      ['Loan Origin Date:', originDate],
      ['Loan Maturity Date:', maturityDate],
      ['Payments Made / Total:', `${monthsPaid} / ${totalMonths}`]
    ]);
    
    // Format header
    sheet.getRange(1, 1).setBackground('#2E7D32').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(14);
    sheet.getRange(2, 1, 7, 1).setBackground('#4CAF50').setFontColor('#FFFFFF').setFontWeight('bold');
    sheet.getRange(2, 2, 7, 1).setBackground('#81C784').setFontWeight('bold');
    
    // Format currency cells
    sheet.getRange(2, 2).setNumberFormat('"$"#,##0.00'); // Original loan
    sheet.getRange(5, 2).setNumberFormat('"$"#,##0.00'); // Monthly payment
    
    // Amortization table headers
    const tableStartRow = 10;
    const headers = ['Payment #', 'Date', 'Payment', 'Principal', 'Interest', 'Balance', 'Status'];
    const headerRange2 = sheet.getRange(tableStartRow, 1, 1, headers.length);
    headerRange2.setValues([headers]);
    headerRange2.setBackground('#1976D2').setFontColor('#FFFFFF').setFontWeight('bold');
    
    // Generate amortization schedule
    const scheduleData = [];
    let balance = originalLoanAmount;
    let currentBalance = 0; // Balance after current payment date
    
    for (let i = 1; i <= totalMonths; i++) {
      const paymentDate = new Date(startDateObj.getFullYear(), startDateObj.getMonth() + i - 1, 1);
      let interestPayment = balance * monthlyRate;
      let principalPayment = monthlyPayment - interestPayment;
      
      // Round to cents for precision
      interestPayment = Math.round(interestPayment * 100) / 100;
      principalPayment = Math.round(principalPayment * 100) / 100;
      
      // Update balance
      balance = Math.max(0, balance - principalPayment);
      balance = Math.round(balance * 100) / 100;
      
      // Capture balance after the current payment date
      if (i === monthsPaid) {
        currentBalance = balance;
      }
      
      // Determine status based on current payment date
      let status = '';
      if (i <= monthsPaid) { // Through current payment date
        status = 'PAID';
      } else if (i === monthsPaid + 1) { // Next month after current payment
        status = 'CURRENT';
      } else {
        status = 'REMAINING';
      }
      
      scheduleData.push([
        i,
        Utilities.formatDate(paymentDate, Session.getScriptTimeZone(), 'MMM yyyy'),
        monthlyPayment,
        principalPayment,
        interestPayment,
        balance,
        status
      ]);
    }
    
    // Add data to sheet with old payments moved to bottom
    const threeMonthsAgo = new Date(currentDateObj.getFullYear(), currentDateObj.getMonth() - 3, 1);
    
    // Separate recent payments (last 3 months + current + future) from old payments
    const recentPayments = [];
    const oldPayments = [];
    
    scheduleData.forEach(payment => {
      const paymentNum = payment[0];
      const paymentDate = new Date(startDateObj.getFullYear(), startDateObj.getMonth() + paymentNum - 1, 1);
      
      if (paymentDate >= threeMonthsAgo || payment[6] === 'CURRENT' || payment[6] === 'REMAINING') {
        recentPayments.push(payment);
      } else {
        oldPayments.push(payment);
      }
    });
    
    // Combine recent payments first, then old payments
    const organizedData = [...recentPayments, ...oldPayments];
    
    // Add separator row between recent and old payments if there are old payments
    let separatorRowIndex = -1;
    if (oldPayments.length > 0) {
      separatorRowIndex = recentPayments.length;
    }
    
    const dataRange = sheet.getRange(tableStartRow + 1, 1, organizedData.length, headers.length);
    dataRange.setValues(organizedData);
    
    // Format currency columns
    const currencyColumns = [3, 4, 5, 6]; // Payment, Principal, Interest, Balance
    currencyColumns.forEach(col => {
      sheet.getRange(tableStartRow + 1, col, organizedData.length, 1).setNumberFormat('"$"#,##0.00');
    });
    
    // Add visual separator between recent and old payments
    if (separatorRowIndex > 0) {
      const separatorRow = tableStartRow + 1 + separatorRowIndex;
      sheet.getRange(separatorRow, 1, 1, headers.length)
        .setBorder(true, false, false, false, false, false, '#FF9800', SpreadsheetApp.BorderStyle.SOLID_THICK);
      
      // Add a note about the separator
      sheet.getRange(separatorRow - 1, 7).setNote('Payments below this line are older than 3 months and have been moved to the bottom for organization.');
    }
    
    // Color code status column with organized data
    for (let i = 0; i < organizedData.length; i++) {
      const statusCell = sheet.getRange(tableStartRow + 1 + i, 7);
      const status = organizedData[i][6];
      
      if (status === 'PAID') {
        statusCell.setBackground('#C8E6C9').setFontColor('#2E7D32').setFontWeight('bold');
      } else if (status === 'CURRENT') {
        statusCell.setBackground('#FFE082').setFontColor('#F57F17').setFontWeight('bold');
      } else if (status === 'REMAINING') {
        statusCell.setBackground('#FFCDD2').setFontColor('#D32F2F');
      }
      
      // Highlight current payment row
      if (status === 'CURRENT') {
        sheet.getRange(tableStartRow + 1 + i, 1, 1, headers.length)
          .setBorder(true, true, true, true, false, false, '#FF9800', SpreadsheetApp.BorderStyle.SOLID_THICK);
      }
      
      // Add subtle background for old payments
      const paymentNum = organizedData[i][0];
      const paymentDate = new Date(startDateObj.getFullYear(), startDateObj.getMonth() + paymentNum - 1, 1);
      if (paymentDate < threeMonthsAgo && status === 'PAID') {
        const rowRange = sheet.getRange(tableStartRow + 1 + i, 1, 1, headers.length);
        rowRange.setBackground('#F5F5F5'); // Light grey background for old payments
      }
    }
    
    // Update the current balance in the header
    sheet.getRange(3, 2).setValue(currentBalance);
    sheet.getRange(3, 2).setNumberFormat('"$"#,##0.00');
    
    // Auto-resize columns
    sheet.autoResizeColumns(1, headers.length);
    
    // Add summary at bottom
    const summaryRow = tableStartRow + organizedData.length + 2;
    const totalInterest = scheduleData.reduce((sum, row) => sum + row[4], 0);
    const totalPaid = monthsPaid * monthlyPayment;
    const interestPaid = scheduleData.slice(0, monthsPaid).reduce((sum, row) => sum + row[4], 0);
    
    sheet.getRange(summaryRow, 1, 5, 2).setValues([
      ['📊 Summary', ''],
      ['Total Interest (Full Loan):', Math.round(totalInterest * 100) / 100],
      ['Payments Made:', `${monthsPaid} of ${totalMonths}`],
      ['Amount Paid to Date:', Math.round(totalPaid * 100) / 100],
      ['Interest Paid to Date:', Math.round(interestPaid * 100) / 100]
    ]);
    
    // Format summary
    sheet.getRange(summaryRow, 1).setBackground('#FF6B35').setFontColor('#FFFFFF').setFontWeight('bold');
    sheet.getRange(summaryRow + 1, 1, 4, 1).setBackground('#FFA726').setFontColor('#FFFFFF').setFontWeight('bold');
    sheet.getRange(summaryRow + 1, 2, 4, 1).setNumberFormat('"$"#,##0.00').setFontWeight('bold');
    
    SpreadsheetApp.getUi().alert(
      'Mortgage Amortization Complete!',
      `Generated amortization table for $${Math.round(originalLoanAmount).toLocaleString()} loan\n` +
      `Payments Made: ${monthsPaid} of ${totalMonths} (through ${Utilities.formatDate(currentDateObj, Session.getScriptTimeZone(), 'MMM yyyy')})\n` +
      `Current Balance: $${currentBalance.toLocaleString()}\n` +
      `Monthly Payment: $${monthlyPayment} (P&I only)\n` +
      `Next Payment Due: ${Utilities.formatDate(new Date(currentDateObj.getFullYear(), currentDateObj.getMonth() + 1, 1), Session.getScriptTimeZone(), 'MMM yyyy')} (#${monthsPaid + 1})\n\n` +
      `📋 Note: Recent payments (last 3 months + future) are shown first,\nolder payments are moved to the bottom for better organization.`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    console.error('Mortgage amortization error:', error);
    SpreadsheetApp.getUi().alert('Error', 'Failed to generate mortgage amortization: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Update mortgage status with current payment information
 */
function updateMortgageStatus() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Mortgage Amortization');
  
  if (!sheet) {
    ui.alert('No Mortgage Sheet', 'Please generate a mortgage amortization table first.', ui.ButtonSet.OK);
    return;
  }
  
  const result = ui.prompt(
    'Update Mortgage Status',
    'Enter the last month you paid (e.g., "Mar 2026", "April 2026", "05/2026"):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (result.getSelectedButton() === ui.Button.OK) {
    const lastPaidInput = result.getResponseText().trim();
    
    // Convert input to standard format
    let standardizedInput = lastPaidInput;
    
    // Handle MM/YYYY format (like "05/2026")
    const mmyyyyMatch = lastPaidInput.match(/^(\d{1,2})\/(\d{4})$/);
    if (mmyyyyMatch) {
      const monthNum = parseInt(mmyyyyMatch[1]);
      const year = mmyyyyMatch[2];
      const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 
                         'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
      if (monthNum >= 1 && monthNum <= 12) {
        standardizedInput = `${monthNames[monthNum - 1]} ${year}`;
      }
    }
    
    // Find the data range - starting from row 11 (headers are at row 10)
    const dataRange = sheet.getDataRange();
    const allData = dataRange.getValues();
    
    // Find the header row (should be row 10, index 9)
    let headerRowIndex = -1;
    for (let i = 0; i < allData.length; i++) {
      if (allData[i][0] === 'Payment #') {
        headerRowIndex = i;
        break;
      }
    }
    
    if (headerRowIndex === -1) {
      ui.alert('Error', 'Cannot find mortgage payment table.', ui.ButtonSet.OK);
      return;
    }
    
    // Find the matching payment row
    let lastPaidRowIndex = -1;
    const availableDates = []; // For debugging
    
    for (let i = headerRowIndex + 1; i < allData.length; i++) {
      const rowData = allData[i];
      
      // Skip empty rows
      if (!rowData[0] || !rowData[1]) continue;
      
      const dateValue = rowData[1]; // Date column
      let dateStr = '';
      
      // Handle both Date objects and string values
      if (dateValue instanceof Date) {
        dateStr = Utilities.formatDate(dateValue, Session.getScriptTimeZone(), 'MMM yyyy');
      } else {
        dateStr = String(dateValue);
      }
      
      availableDates.push(dateStr); // Collect for debugging
      
      // Try exact match first
      if (dateStr === standardizedInput) {
        lastPaidRowIndex = i;
        break;
      }
      
      // Try flexible matching as fallback
      const inputLower = standardizedInput.toLowerCase().replace(/[^\w]/g, '');
      const dateLower = dateStr.toLowerCase().replace(/[^\w]/g, '');
      
      if (dateLower.includes(inputLower) || inputLower.includes(dateLower)) {
        lastPaidRowIndex = i;
        break;
      }
    }
    
    if (lastPaidRowIndex === -1) {
      // Show available dates for debugging
      const first10Dates = availableDates.slice(0, 10).join(', ');
      const totalDates = availableDates.length;
      ui.alert('Error', `Cannot find payment for "${lastPaidInput}" (searched as "${standardizedInput}").\n\nFirst 10 available dates: ${first10Dates}\n\nTotal payments found: ${totalDates}\n\nPlease use exact format like "Apr 2026".`, ui.ButtonSet.OK);
      return;
    }
    
    // Validate the found row has proper data
    const foundRow = allData[lastPaidRowIndex];
    const foundPaymentNumber = foundRow[0];
    const foundDate = foundRow[1];
    const foundBalance = foundRow[5];
    
    if (!foundPaymentNumber || !foundDate || foundBalance === undefined || foundBalance === null) {
      // Show debug info to understand the data structure
      const debugInfo = `Row Index: ${lastPaidRowIndex}\nRow Data: [${foundRow.join(', ')}]\nExpected columns: Payment#, Date, Payment, Principal, Interest, Balance, Status`;
      ui.alert('Debug Info', debugInfo, ui.ButtonSet.OK);
      return;
    }
    
    // Update all payment statuses
    let updatesCount = 0;
    for (let i = headerRowIndex + 1; i < allData.length; i++) {
      const paymentNumber = allData[i][0];
      const currentStatus = allData[i][6]; // Status column
      let newStatus = '';
      
      if (paymentNumber <= foundPaymentNumber) {
        newStatus = 'PAID';
      } else if (paymentNumber === foundPaymentNumber + 1) {
        newStatus = 'CURRENT';
      } else {
        newStatus = 'REMAINING';
      }
      
      // Update if status changed
      if (currentStatus !== newStatus) {
        sheet.getRange(i + 1, 7).setValue(newStatus); // +1 because sheet rows are 1-indexed
        
        // Update color coding
        const statusCell = sheet.getRange(i + 1, 7);
        if (newStatus === 'PAID') {
          statusCell.setBackground('#C8E6C9').setFontColor('#2E7D32').setFontWeight('bold');
        } else if (newStatus === 'CURRENT') {
          statusCell.setBackground('#FFE082').setFontColor('#F57F17').setFontWeight('bold');
        } else if (newStatus === 'REMAINING') {
          statusCell.setBackground('#FFCDD2').setFontColor('#D32F2F').setFontWeight('bold');
        }
        updatesCount++;
      }
    }
    
    // Update the current balance in header based on the last paid row's balance
    const lastPaidBalance = foundBalance; // Balance column
    const lastPaidDate = foundDate; // Date column
    const lastPaidPaymentNum = foundPaymentNumber; // Payment number
    
    // Format the date properly for display
    let displayDate = '';
    if (lastPaidDate instanceof Date) {
      displayDate = Utilities.formatDate(lastPaidDate, Session.getScriptTimeZone(), 'MM/yyyy');
    } else {
      displayDate = String(lastPaidDate).replace(/[^\w\s]/g, ''); // Remove special chars
      // Convert "May 2026" to "05/2026" format
      const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 
                         'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
      const parts = displayDate.split(' ');
      if (parts.length === 2) {
        const monthIndex = monthNames.indexOf(parts[0]);
        if (monthIndex !== -1) {
          const monthNum = String(monthIndex + 1).padStart(2, '0');
          displayDate = `${monthNum}/${parts[1]}`;
        }
      }
    }
    
    // Update current balance amount (B3)
    sheet.getRange(3, 2).setValue(lastPaidBalance);
    
    // Update current balance description to show new date (A3)
    sheet.getRange(3, 1).setValue(`Current Balance (after ${displayDate}):`);
    
    // Update payments made count (B8)
    // Calculate total payments from loan parameters instead of hardcoding
    const currentPaymentsStr = sheet.getRange(8, 2).getValue();
    let totalPayments = '480'; // fallback
    
    // Try to get total from existing data first
    const existingMatch = String(currentPaymentsStr).match(/\d+\s*\/\s*(\d+)/);
    if (existingMatch) {
      totalPayments = existingMatch[1];
    } else {
      // Calculate total payments from loan parameters if not available
      try {
        const originDate = sheet.getRange('B6').getDisplayValue();
        const maturityDate = sheet.getRange('B7').getDisplayValue();
        const [startMonth, startYear] = originDate.split('/');
        const [endMonth, endYear] = maturityDate.split('/');
        const startDateObj = new Date(parseInt(startYear), parseInt(startMonth) - 1, 1);
        const endDateObj = new Date(parseInt(endYear), parseInt(endMonth) - 1, 1);
        const calculatedTotalMonths = (endDateObj.getFullYear() - startDateObj.getFullYear()) * 12 + 
                                     (endDateObj.getMonth() - startDateObj.getMonth());
        totalPayments = String(calculatedTotalMonths);
      } catch (error) {
        console.log('Could not calculate total payments, using fallback');
      }
    }
    sheet.getRange(8, 2).setValue(`${lastPaidPaymentNum} / ${totalPayments}`);
    
    ui.alert('Updated', 
      `Mortgage status updated!\n` +
      `Last paid: ${displayDate.includes('/') ? displayDate : standardizedInput}\n` +
      `Payments made: ${lastPaidPaymentNum}\n` +
      `Current balance: $${Number(lastPaidBalance).toLocaleString()}\n` +
      `Updated ${updatesCount} payment statuses.`, 
      ui.ButtonSet.OK);
  }
}

/**
 * Generate payoff schedule for a selected card (called from sidebar)
 */
function generatePayoffScheduleForCard(cardName, targetMonths) {
  const cards = getCreditCards();
  const card = cards.find(c => c.name === cardName);
  
  if (!card) {
    throw new Error('Card not found: ' + cardName);
  }
  
  // Use provided target months or fallback to stored value
  const months = targetMonths || card.targetMonths || 24;
  const requiredPayment = calculateMonthlyPayment(card.balance, card.apr, months);
  return generatePayoffSchedule(cardName, card.balance, card.apr, requiredPayment, card.minPayment);
}

/**
 * Generate custom payment schedule with unique naming (called from sidebar)
 */
function generateCustomPaymentSchedule(cardName, payment) {
  const cards = getCreditCards();
  const card = cards.find(c => c.name === cardName);
  
  if (!card) {
    throw new Error('Card not found: ' + cardName);
  }
  
  return generateCustomPayoffSchedule(cardName, card.balance, card.apr, payment, card.minPayment);
}

/**
 * Generate custom payoff schedule that doesn't overwrite existing schedules
 */
function generateCustomPayoffSchedule(cardName, balance, apr, customPayment, minPayment) {
  const timeline = calculatePayoffTimeline(balance, apr, customPayment);
  
  if (timeline.error) {
    throw new Error(timeline.error);
  }
  
  const baseSheetName = `${cardName} Schedule`.substring(0, 30);
  const sheet = getOrCreateCustomSheet(baseSheetName, true); // Use custom sheet logic
  
  // Clear existing data (only on the new custom sheet)
  sheet.clear();
  
  // Add summary info with vibrant styling for custom schedule
  const summaryRange = sheet.getRange(1, 1, 5, 2);
  summaryRange.setValues([
    ['🏦 Card Name:', cardName],
    ['💰 Starting Balance:', balance],
    ['📈 APR:', apr],
    ['💳 Custom Payment:', customPayment], // Show custom payment instead of minimum
    ['⏱️ Payoff Time:', `${timeline.totalMonths} months`]
  ]);
  
  // Format the numeric values properly
  sheet.getRange(2, 2).setNumberFormat('"$"#,##0.00'); // Starting Balance with 2 decimals
  sheet.getRange(3, 2).setNumberFormat('0.00"%"'); // APR
  sheet.getRange(4, 2).setNumberFormat('"$"#,##0.00'); // Custom Payment
  
  // Style summary section with different colors to distinguish custom schedule
  sheet.getRange(1, 1, 5, 1).setBackground('#FF6B35')  // Energetic orange for custom
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');
  sheet.getRange(1, 2, 5, 1).setBackground('#E74C3C')  // Energetic red for custom
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');
  
  // Add timeline data starting from row 7 (moved up since we removed a summary row)
  const startRow = 7;
  const headers = ['Custom Payoff', 'Payment', 'Interest', 'Principal', 'Remaining Balance'];
  const timelineHeaderRange = sheet.getRange(startRow, 1, 1, headers.length);
  timelineHeaderRange.setValues([headers]);
  timelineHeaderRange.setBackground('#9C27B0')  // Purple for custom schedule
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setFontSize(11)
    .setBorder(true, true, true, true, true, true, '#7B1FA2', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  
  const timelineData = timeline.timeline.map(row => [
    row.month,
    row.payment,
    row.interest,
    row.principal,
    row.balance
  ]);
  
  if (timelineData.length > 0) {
    const dataRange = sheet.getRange(startRow + 1, 1, timelineData.length, headers.length);
    dataRange.setValues(timelineData);
    
    // Apply custom styling for custom schedules
    for (let i = 0; i < timelineData.length; i++) {
      const rowRange = sheet.getRange(startRow + 1 + i, 1, 1, headers.length);
      if (i % 2 === 0) {
        rowRange.setBackground('#FFF8E1');  // Light amber for custom
      } else {
        rowRange.setBackground('#F3E5F5');  // Light purple for custom
      }
      
      rowRange.setBorder(false, false, true, false, false, false, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);
      
      if (timelineData[i][0] % 6 === 0) {
        rowRange.setFontWeight('bold')
          .setBackground('#FFCC80');  // Orange for milestones
      }
    }
  }
  
  // Format currency columns
  const currencyColumns = [2, 3, 4, 5];
  currencyColumns.forEach(col => {
    const currencyRange = sheet.getRange(startRow + 1, col, timelineData.length, 1);
    currencyRange.setNumberFormat('"$"#,##0.00')
      .setHorizontalAlignment('right');
  });
  
  // Add completion message
  if (timeline.totalMonths < timelineData.length || timelineData[timelineData.length - 1][4] === 0) {
    const finalRow = startRow + 1 + timelineData.length;
    sheet.getRange(finalRow, 1, 1, 5).setValues([['', 'CUSTOM GOAL ACHIEVED!', '🎈', 'WELL DONE!', '🎊']]);
    sheet.getRange(finalRow, 1, 1, 5)
      .setBackground('#FF9800')  // Orange for custom achievement
      .setFontColor('#FFFFFF')
      .setFontWeight('bold')
      .setFontSize(14)
      .setHorizontalAlignment('center');
  }
  
  // Add interactive timeline for the custom schedule
  addInteractiveTimeline(sheet, balance, apr, customPayment, timeline.totalMonths, true, cardName); // true for custom schedule
  
  sheet.autoResizeColumns(1, headers.length);
  sheet.autoResizeColumns(7, 5);
  
  return timeline;
}

/**
 * Generate default scenarios for comparison (called from sidebar)
 */
function generateDefaultScenarios(cardName, targetMonths) {
  const cards = getCreditCards();
  const card = cards.find(c => c.name === cardName);
  
  if (!card) {
    throw new Error('Card not found: ' + cardName);
  }
  
  // Use provided target months or fallback to stored value
  const months = targetMonths || card.targetMonths || 24;
  
  const scenarios = [
    { name: 'Minimum Payment', payment: card.minPayment },
    { name: 'Double Minimum', payment: card.minPayment * 2 },
    { name: `${months} Month Target`, payment: calculateMonthlyPayment(card.balance, card.apr, months) },
    { name: '12 Month Payoff', payment: calculateMonthlyPayment(card.balance, card.apr, 12) },
    { name: '36 Month Payoff', payment: calculateMonthlyPayment(card.balance, card.apr, 36) }
  ];
  
  // Filter out scenarios with payments less than minimum payment
  const validScenarios = scenarios.filter(scenario => scenario.payment >= card.minPayment);
  
  return compareScenarios(cardName, card.balance, card.apr, validScenarios);
}

/**
 * Run custom scenarios from sidebar input (called from sidebar)
 */
function runCustomScenarios(cardName, scenarios) {
  const cards = getCreditCards();
  const card = cards.find(c => c.name === cardName);
  
  if (!card) {
    throw new Error('Card not found: ' + cardName);
  }
  
  // Filter out scenarios with payments less than minimum payment
  const validScenarios = scenarios.filter(scenario => scenario.payment >= card.minPayment);
  
  // Add minimum payment as baseline if not already included
  const hasMinPayment = validScenarios.some(s => s.payment === card.minPayment);
  if (!hasMinPayment) {
    validScenarios.unshift({ name: 'Minimum Payment', payment: card.minPayment });
  }
  
  if (validScenarios.length === 0) {
    throw new Error('No valid scenarios - all payments are less than minimum payment of $' + card.minPayment);
  }
  
  return compareScenarios(cardName, card.balance, card.apr, validScenarios);
}

/**
 * Google Sheets simple trigger - automatically called when any cell is edited
 * This replaces the need for installable triggers
 */
function onEdit(e) {
  // Call our validation logic
  onValidationChange(e);
}

/**
 * Handle validation checkbox changes - called when checkboxes in column L are clicked
 * This function updates the Credit Cards sheet when payments are validated
 */
function onValidationChange(e) {
  try {
    console.log('onValidationChange triggered');
    console.log('Event range:', e.range.getA1Notation());
    console.log('Event column:', e.range.getColumn());
    console.log('Event row:', e.range.getRow());
    console.log('Event value:', e.range.getValue());
    
    // Only process changes in column L (validation column)
    if (e.range.getColumn() !== 12) {
      console.log('Not column L, ignoring');
      return; // Column L = 12
    }
    
    const sheet = e.source.getActiveSheet();
    const sheetName = sheet.getName();
    console.log('Sheet name:', sheetName);
    
    // Only process schedule sheets (containing " Schedule")
    if (!sheetName.includes(' Schedule')) {
      console.log('Not a schedule sheet, ignoring');
      return;
    }
    
    // Extract card name from sheet name (remove " Schedule" and any trailing numbers/timestamps)
    let cardName = sheetName.replace(' Schedule', '');
    // Remove custom schedule timestamps like " (Custom 14:30)"
    cardName = cardName.replace(/\s*\(Custom\s+\d{2}:\d{2}\)/, '').trim();
    console.log('Extracted card name:', cardName);
    
    // Get the checkbox value and row
    const isChecked = e.range.getValue();
    const row = e.range.getRow();
    console.log('Checkbox checked:', isChecked, 'Row:', row);
    
    // Only process when checkbox is checked (not unchecked)
    if (isChecked !== true) {
      console.log('Checkbox not checked, ignoring');
      return;
    }
    
    // Get the month number from column G
    const monthCell = sheet.getRange(row, 7); // Column G
    const month = monthCell.getValue();
    console.log('Month from G' + row + ':', month, typeof month);
    
    // Get the payment amount from column H
    const paymentCell = sheet.getRange(row, 8); // Column H
    const paymentAmount = paymentCell.getValue();
    console.log('Payment from H' + row + ':', paymentAmount, typeof paymentAmount);
    
    // Skip if no valid month or payment amount
    if (!month || !paymentAmount || month <= 0 || typeof month !== 'number') {
      console.log('Invalid month or payment amount, skipping');
      return;
    }
    
    // Update the Credit Cards sheet
    console.log('About to update Credit Cards sheet for:', cardName, 'month:', month, 'amount:', paymentAmount);
    updateCreditCardPayment(cardName, month, paymentAmount);
    
    // Provide user feedback
    SpreadsheetApp.getUi().alert(
      'Payment Validated',
      `Payment of $${paymentAmount.toFixed(2)} for month ${month} has been recorded for ${cardName} in the Credit Cards sheet.`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (error) {
    console.error('onValidationChange error:', error);
    SpreadsheetApp.getUi().alert('Error', 'Validation error: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Update the Credit Cards sheet with validated payment
 */
function updateCreditCardPayment(cardName, month, paymentAmount) {
  try {
    console.log('updateCreditCardPayment called with:', {cardName, month, paymentAmount});
    
    const creditCardsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Credit Cards');
    
    if (!creditCardsSheet) {
      console.error('Credit Cards sheet not found');
      SpreadsheetApp.getUi().alert('Error', 'Credit Cards sheet not found.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // Find the row for this card - only search in data rows (skip header)
    const dataRange = creditCardsSheet.getDataRange();
    const data = dataRange.getValues();
    console.log('Credit Cards sheet data rows:', data.length);
    let cardRow = -1;
    
    // Start from row 1 (index 1) to skip header row
    for (let i = 1; i < data.length; i++) {
      const cellValue = data[i][0];
      console.log(`Row ${i+1}: Card name in sheet: "${cellValue}", looking for: "${cardName}"`);
      if (cellValue && cellValue.toString().trim() === cardName.trim()) { // Column A contains card names
        cardRow = i + 1; // Convert to 1-based row number
        console.log('Found matching card at row:', cardRow);
        break;
      }
    }
    
    if (cardRow === -1) {
      const availableCards = data.slice(1).map(row => `"${row[0]}"`).filter(name => name !== '""').join(', ');
      console.error('Card not found. Available cards:', availableCards);
      SpreadsheetApp.getUi().alert('Error', `Card "${cardName}" not found in Credit Cards sheet. Available cards: ${availableCards}`, SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // Validate month is within range
    if (month < 1 || month > 12) {
      console.error('Month out of range:', month);
      SpreadsheetApp.getUi().alert('Error', `Month ${month} is beyond the supported range (1-12).`, SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // Calculate the column for this month (Month 1 = Column I = 9)
    const monthColumn = 8 + month; // I=9, J=10, K=11, etc.
    console.log('Updating cell at row:', cardRow, 'column:', monthColumn, '(Month', month + ')');
    
    // Get the current value before updating
    const targetCell = creditCardsSheet.getRange(cardRow, monthColumn);
    const currentValue = targetCell.getValue();
    console.log('Current cell value:', currentValue);
    
    // Update ONLY the specific cell for this card and month
    targetCell.setValue(paymentAmount);
    console.log('Set cell value to:', paymentAmount);
    
    // Add formatting to highlight the validated payment
    targetCell
      .setBackground('#C8E6C9')  // Light green background
      .setFontColor('#2E7D32')   // Dark green text
      .setFontWeight('bold');
      
    // Verify the update worked
    const newValue = targetCell.getValue();
    console.log('Cell value after update:', newValue);
    
    // Final confirmation
    console.log(`Successfully updated ${cardName} row ${cardRow}, month ${month} (column ${monthColumn}) with payment $${paymentAmount}`);
    
  } catch (error) {
    console.error('updateCreditCardPayment error:', error);
    SpreadsheetApp.getUi().alert('Error', 'Update error: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Generate simulation mortgage amortization table (exact copy of main function)
 * Outputs to columns I-N starting at row 10
 */
function generateSimulationMortgageAmortization(originDate, originalLoanAmount, annualRate, maturityDate, currentPaymentDate, extraPayment = 0, extraPaymentStartMonth = null, threePaycheckAmount = 0, threePaycheckStartMonth = null) {
  try {
    console.log('=== Simulation Calculation Function Debug ===');
    console.log('Received parameters:');
    console.log('  - originDate:', originDate, 'type:', typeof originDate);
    console.log('  - originalLoanAmount:', originalLoanAmount, 'type:', typeof originalLoanAmount);
    console.log('  - annualRate:', annualRate, 'type:', typeof annualRate);
    console.log('  - maturityDate:', maturityDate, 'type:', typeof maturityDate);
    console.log('  - currentPaymentDate:', currentPaymentDate, 'type:', typeof currentPaymentDate);
    
    // Validate parameters before using split()
    if (!originDate || typeof originDate !== 'string') {
      throw new Error(`Invalid originDate: ${originDate} (type: ${typeof originDate})`);
    }
    if (!maturityDate || typeof maturityDate !== 'string') {
      throw new Error(`Invalid maturityDate: ${maturityDate} (type: ${typeof maturityDate})`);
    }
    if (!currentPaymentDate || typeof currentPaymentDate !== 'string') {
      throw new Error(`Invalid currentPaymentDate: ${currentPaymentDate} (type: ${typeof currentPaymentDate})`);
    }
    if (!originalLoanAmount || isNaN(originalLoanAmount)) {
      throw new Error(`Invalid originalLoanAmount: ${originalLoanAmount} (type: ${typeof originalLoanAmount})`);
    }
    
    console.log('All parameters validated successfully');
    
    // Parse dates
    const [startMonth, startYear] = originDate.split('/');
    const [endMonth, endYear] = maturityDate.split('/');
    const [currentMonth, currentYear] = currentPaymentDate.split('/');
    
    const startDateObj = new Date(parseInt(startYear), parseInt(startMonth) - 1, 1);
    const endDateObj = new Date(parseInt(endYear), parseInt(endMonth) - 1, 1);
    const currentDateObj = new Date(parseInt(currentYear), parseInt(currentMonth) - 1, 1);
    
    // Calculate months
    const totalMonths = (endDateObj.getFullYear() - startDateObj.getFullYear()) * 12 + 
                       (endDateObj.getMonth() - startDateObj.getMonth());
    const monthsPaid = (currentDateObj.getFullYear() - startDateObj.getFullYear()) * 12 + 
                      (currentDateObj.getMonth() - startDateObj.getMonth()) + 1;
    
    // Calculate monthly payment from original loan amount
    const monthlyRate = annualRate / 100 / 12;
    const monthlyPayment = calculateMortgagePayment(originalLoanAmount, annualRate, totalMonths);
    
    console.log('=== Calculation Setup ===');
    console.log('totalMonths:', totalMonths);
    console.log('monthsPaid:', monthsPaid);
    console.log('annualRate as percentage:', annualRate, '(should be ~2.75)');
    console.log('monthlyRate calculation:', annualRate, '/ 100 / 12 =', monthlyRate);
    console.log('monthlyRate:', monthlyRate, '(should be ~0.00229)');
    console.log('monthlyPayment:', monthlyPayment, '(should be ~579.48)');
    console.log('Starting balance (originalLoanAmount):', originalLoanAmount);
    
    if (monthlyPayment < 500 || monthlyPayment > 700) {
      console.log('⚠️ WARNING: Monthly payment looks incorrect! Expected ~$579.48');
    }
    if (monthlyRate < 0.002 || monthlyRate > 0.003) {
      console.log('⚠️ WARNING: Monthly rate looks incorrect! Expected ~0.00229');
    }
    
    // Get existing sheet (don't clear it)
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Mortgage Amortization');
    
    // Clear simulation area (columns I through N, starting from row 9)
    const simStartCol = 9; // Column I
    const simStartRow = 10;
    sheet.getRange(9, simStartCol, 500, 6).clearContent().clearFormat();
    
    // Create simulation title (at row 9)
    const titleRange = sheet.getRange(9, simStartCol, 1, 6);
    titleRange.setValues([['📊 Simulation Table (Exact Copy)', '', '', '', '', '']]);
    titleRange.setBackground('#FFE0B2').setFontColor('#FF9800').setFontWeight('bold').setFontSize(12);
    titleRange.merge();
    
    // Simulation table headers (at row 10) - include Extra column
    const headers = ['Payment #', 'Date', 'Payment', 'Principal', 'Interest', 'Balance', 'Extra'];
    const headerRange = sheet.getRange(simStartRow, simStartCol, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.setBackground('#1976D2').setFontColor('#FFFFFF').setFontWeight('bold');
    
    // Add extra payment and 3-paycheck payment logging
    console.log('=== Extra Payment Parameters ===');
    console.log('extraPayment:', extraPayment);
    console.log('extraPaymentStartMonth:', extraPaymentStartMonth);
    console.log('threePaycheckAmount:', threePaycheckAmount);
    console.log('threePaycheckStartMonth:', threePaycheckStartMonth);
    
    // Parse extra payment start month if provided
    let extraPaymentStartDate = null;
    if (extraPaymentStartMonth && extraPayment > 0) {
      const [extraMonth, extraYear] = extraPaymentStartMonth.split('/');
      extraPaymentStartDate = new Date(parseInt(extraYear), parseInt(extraMonth) - 1, 1);
      console.log('Parsed extraPaymentStartDate:', extraPaymentStartDate);
    }
    
    // Parse 3-paycheck start month and get all 3-paycheck months
    let threePaycheckStartDate = null;
    let threePaycheckMonths = [];
    if (threePaycheckStartMonth && threePaycheckAmount > 0) {
      const [threeMonth, threeYear] = threePaycheckStartMonth.split('/');
      threePaycheckStartDate = new Date(parseInt(threeYear), parseInt(threeMonth) - 1, 1);
      
      // Get all 3-paycheck months for the next 40 years
      const allThreePaycheckMonths = calculateThreePaycheckMonths(parseInt(threeYear), 40);
      threePaycheckMonths = allThreePaycheckMonths
        .filter(month => {
          const monthDate = new Date(month.year, month.month - 1, 1);
          return monthDate >= threePaycheckStartDate;
        })
        .map(month => month.formatted);
      
      console.log('Parsed threePaycheckStartDate:', threePaycheckStartDate);
      console.log('First 10 eligible 3-paycheck months:', threePaycheckMonths.slice(0, 10));
    }

    // Generate amortization schedule with extra payments
    const scheduleData = [];
    let balance = originalLoanAmount;
    let currentBalance = 0; // Balance after current payment date
    
    console.log('=== Starting Amortization Loop ===');
    console.log('Initial balance:', balance);
    
    for (let i = 1; i <= totalMonths; i++) {
      const paymentDate = new Date(startDateObj.getFullYear(), startDateObj.getMonth() + i - 1, 1);
      let interestPayment = balance * monthlyRate;
      
      // Calculate base principal payment
      let principalPayment = monthlyPayment - interestPayment;
      let totalPaymentAmount = monthlyPayment;
      
      // Add extra payment if we're at or past the start date
      let extraPaymentThisMonth = 0;
      if (extraPaymentStartDate && paymentDate >= extraPaymentStartDate && balance > 0) {
        extraPaymentThisMonth = Math.min(extraPayment, balance - principalPayment);
        if (extraPaymentThisMonth > 0) {
          principalPayment += extraPaymentThisMonth;
          totalPaymentAmount += extraPaymentThisMonth;
        }
      }
      
      // Add 3-paycheck payment if this is a 3-paycheck month
      let threePaycheckPaymentThisMonth = 0;
      const currentMonthFormatted = Utilities.formatDate(paymentDate, Session.getScriptTimeZone(), 'MM/yyyy');
      if (threePaycheckMonths.includes(currentMonthFormatted) && balance > 0) {
        threePaycheckPaymentThisMonth = Math.min(threePaycheckAmount, balance - principalPayment);
        if (threePaycheckPaymentThisMonth > 0) {
          principalPayment += threePaycheckPaymentThisMonth;
          totalPaymentAmount += threePaycheckPaymentThisMonth;
        }
      }
      
      // Log first few payments for debugging
      if (i <= 3) {
        console.log(`=== Payment ${i} Calculation ===`);
        console.log('  Payment date:', Utilities.formatDate(paymentDate, Session.getScriptTimeZone(), 'MM/yyyy'));
        console.log('  Balance before payment:', balance);
        console.log('  Interest calculation:', balance, '*', monthlyRate, '=', interestPayment);
        console.log('  Base principal calculation:', monthlyPayment, '-', interestPayment, '=', monthlyPayment - interestPayment);
        console.log('  Extra payment this month:', extraPaymentThisMonth);
        console.log('  3-paycheck payment this month:', threePaycheckPaymentThisMonth);
        console.log('  Total principal payment:', principalPayment);
        console.log('  Total payment amount:', totalPaymentAmount);
        
        // TODO: Add more detailed console.log() for next session:
        // - Log exact precision of floating point calculations
        // - Compare each calculation step with main table equivalent
        // - Add validation that calculations are mathematically correct
        // - Log intermediate rounding effects and precision loss
        // - Add boundary condition checks (balance >= 0, payments > 0)
        // - Log payment date formatting and validation
        // - Add detailed status determination logic
        // - Track cumulative errors over multiple payments
        // - Log memory allocation for large payment arrays  
        // - Add performance timing for each calculation loop iteration
      }
      
      // Round to cents for precision
      interestPayment = Math.round(interestPayment * 100) / 100;
      principalPayment = Math.round(principalPayment * 100) / 100;
      
      // Update balance
      balance = Math.max(0, balance - principalPayment);
      balance = Math.round(balance * 100) / 100;
      
      // Log first few payments after rounding
      if (i <= 3) {
        console.log('  After rounding:');
        console.log('    Interest:', interestPayment);
        console.log('    Principal:', principalPayment);
        console.log('    New balance:', balance);
      }
      
      // Capture balance after the current payment date
      if (i === monthsPaid) {
        currentBalance = balance;
      }
      
      // Determine status based on current payment date (for organization)
      let status = '';
      if (i <= monthsPaid) { // Through current payment date
        status = 'PAID';
      } else if (i === monthsPaid + 1) { // Next month after current payment
        status = 'CURRENT';
      } else {
        status = 'REMAINING';
      }
      
      // Check if loan is paid off early due to extra payments
      if (balance <= 0) {
        console.log(`Loan paid off early at payment ${i}!`);
        // Add final payment entry and break
        const extraInfo = [];
        if (extraPaymentThisMonth > 0) extraInfo.push(`+$${extraPaymentThisMonth}`);
        if (threePaycheckPaymentThisMonth > 0) extraInfo.push(`3-pay: +$${threePaycheckPaymentThisMonth}`);
        
        scheduleData.push([
          i,
          Utilities.formatDate(paymentDate, Session.getScriptTimeZone(), 'MMM yyyy'),
          totalPaymentAmount,
          principalPayment,
          interestPayment,
          0,
          extraInfo.join(' ')
        ]);
        break;
      }
      
      // Format extra payment info
      const extraInfo = [];
      if (extraPaymentThisMonth > 0) extraInfo.push(`+$${extraPaymentThisMonth}`);
      if (threePaycheckPaymentThisMonth > 0) extraInfo.push(`3-pay: +$${threePaycheckPaymentThisMonth}`);
      
      scheduleData.push([
        i,
        Utilities.formatDate(paymentDate, Session.getScriptTimeZone(), 'MMM yyyy'),
        totalPaymentAmount,
        principalPayment,
        interestPayment,
        balance,
        extraInfo.join(' ')
      ]);
    }
    
    console.log('=== Amortization Loop Complete ===');
    console.log('Total schedule entries:', scheduleData.length);
    if (scheduleData.length > 0) {
      console.log('First payment entry:', scheduleData[0]);
      console.log('Second payment entry:', scheduleData[1]);
      console.log('Third payment entry:', scheduleData[2]);
    }
    
    // Add data to simulation area with old payments moved to bottom (EXACT same logic)
    const threeMonthsAgo = new Date(currentDateObj.getFullYear(), currentDateObj.getMonth() - 3, 1);
    
    // TODO: Add more detailed console.log() for next session:
    // - Log the three months ago cutoff date calculation
    // - Track how many payments fall into recent vs old categories
    // - Validate date comparisons are working correctly
    // - Log array sizes before and after organization
    // - Add detailed payment sorting and filtering logic
    // - Validate no payments are lost during organization
    // - Log final organized data structure integrity
    // - Add performance metrics for array processing operations
    
    // Separate recent payments (last 3 months + current + future) from old payments
    const recentPayments = [];
    const oldPayments = [];
    
    scheduleData.forEach(payment => {
      const paymentNum = payment[0];
      const paymentDate = new Date(startDateObj.getFullYear(), startDateObj.getMonth() + paymentNum - 1, 1);
      
      if (paymentDate >= threeMonthsAgo || payment[6] === 'CURRENT' || payment[6] === 'REMAINING') {
        recentPayments.push(payment);
      } else {
        oldPayments.push(payment);
      }
    });
    
    // Combine recent payments first, then old payments
    const organizedData = [...recentPayments, ...oldPayments];
    
    // Keep all 7 columns for simulation data (including Extra column)
    const simulationData = organizedData.map(row => row.slice(0, 7));
    
    // Add data to simulation area
    const dataRange = sheet.getRange(simStartRow + 1, simStartCol, simulationData.length, headers.length);
    dataRange.setValues(simulationData);
    
    console.log('=== Data Written to Sheet ===');
    console.log('Simulation data rows:', simulationData.length);
    console.log('Starting at row:', simStartRow + 1, 'column:', simStartCol);
    console.log('Headers:', headers);
    if (simulationData.length > 0) {
      console.log('First row data written:', simulationData[0]);
    }
    
    // Format currency columns (same as main table) - columns 3,4,5,6 (Payment, Principal, Interest, Balance)
    const currencyColumns = [3, 4, 5, 6]; // Payment, Principal, Interest, Balance (Extra column is text)
    currencyColumns.forEach(col => {
      sheet.getRange(simStartRow + 1, simStartCol + col - 1, simulationData.length, 1).setNumberFormat('"$"#,##0.00');
    });
    
    // Add visual separator between recent and old payments
    if (oldPayments.length > 0) {
      const separatorRowIndex = recentPayments.length;
      const separatorRow = simStartRow + 1 + separatorRowIndex;
      sheet.getRange(separatorRow, simStartCol, 1, headers.length)
        .setBorder(true, false, false, false, false, false, '#FF9800', SpreadsheetApp.BorderStyle.SOLID_THICK);
    }
    
    // Color coding and highlighting (same as main table)
    for (let i = 0; i < simulationData.length; i++) {
      const paymentNum = simulationData[i][0];
      const paymentDate = new Date(startDateObj.getFullYear(), startDateObj.getMonth() + paymentNum - 1, 1);
      const rowRange = sheet.getRange(simStartRow + 1 + i, simStartCol, 1, headers.length);
      
      // Highlight current payment row
      if (paymentNum === monthsPaid + 1) {
        rowRange.setBorder(true, true, true, true, false, false, '#FF9800', SpreadsheetApp.BorderStyle.SOLID_THICK);
      }
      
      // Add subtle background for old payments
      if (paymentDate < threeMonthsAgo && paymentNum <= monthsPaid) {
        rowRange.setBackground('#F5F5F5'); // Light grey background for old payments
      }
    }
    
    // Auto-resize simulation columns
    for (let col = simStartCol; col < simStartCol + headers.length; col++) {
      sheet.autoResizeColumn(col);
    }
    
    SpreadsheetApp.getUi().alert(
      'Simulation Table Created!',
      'Simulation table generated in columns I-N.\nThis uses the exact same calculation logic as your main amortization table.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    console.error('Simulation error:', error);
    SpreadsheetApp.getUi().alert('Error', 'Failed to generate simulation: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Wrapper function to call simulation with existing mortgage data
 */
/**
 * Calculate all months with 3 paychecks for bi-weekly pay schedule
 * Based on the user's pattern: Nov 2026, Jun 2027, Nov 2027, May 2028, etc.
 */
function calculateThreePaycheckMonths(startYear = 2026, yearsAhead = 40) {
  const threePaycheckMonths = [];
  
  // Starting from November 2026, the pattern roughly follows:
  // Nov 2026 -> Jun 2027 (+7 months) -> Nov 2027 (+5 months) -> May 2028 (+6 months) -> Nov 2028 (+6 months) -> Apr 2029 (+5 months), etc.
  // The pattern alternates between 5-7 month gaps due to bi-weekly pay cycle
  
  let currentMonth = 11; // November
  let currentYear = 2026;
  const endYear = startYear + yearsAhead;
  
  // The gap pattern for bi-weekly pay with 3-paycheck months
  const gapPattern = [7, 5, 6, 6, 5, 6, 6, 5, 6, 6, 5]; // Months between 3-paycheck months
  let gapIndex = 0;
  
  while (currentYear <= endYear) {
    threePaycheckMonths.push({
      month: currentMonth,
      year: currentYear,
      formatted: `${String(currentMonth).padStart(2, '0')}/${currentYear}`
    });
    
    // Add the next gap to get to the next 3-paycheck month
    const monthsToAdd = gapPattern[gapIndex % gapPattern.length];
    currentMonth += monthsToAdd;
    
    // Handle year rollover
    while (currentMonth > 12) {
      currentMonth -= 12;
      currentYear++;
    }
    
    gapIndex++;
  }
  
  // Filter to only include months from startYear onward
  const filteredMonths = threePaycheckMonths
    .filter(item => item.year >= startYear)
    .sort((a, b) => (a.year * 12 + a.month) - (b.year * 12 + b.month));
  
  console.log('=== Calculated 3-Paycheck Months ===');
  console.log('Total months found:', filteredMonths.length);
  console.log('First 20:', filteredMonths.slice(0, 20).map(m => m.formatted));
  
  return filteredMonths;
}

function generateExtraPaymentSimulation() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Mortgage Amortization');
    if (!sheet) {
      SpreadsheetApp.getUi().alert('Error', 'No mortgage amortization sheet found. Please generate one first.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    const ui = SpreadsheetApp.getUi();
    
    // Prompt for extra payment amount
    const extraPaymentResult = ui.prompt(
      'Extra Payment Simulation',
      'Enter the extra monthly payment amount (e.g., 100, 200, 500):',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (extraPaymentResult.getSelectedButton() !== ui.Button.OK) {
      return;
    }
    
    const extraPaymentInput = extraPaymentResult.getResponseText().trim();
    const extraPayment = parseFloat(extraPaymentInput.replace(/[$,]/g, ''));
    
    if (!extraPayment || isNaN(extraPayment) || extraPayment <= 0) {
      ui.alert('Error', `Invalid extra payment amount: "${extraPaymentInput}". Please enter a positive number.`, ui.ButtonSet.OK);
      return;
    }
    
    // Prompt for starting month
    const startMonthResult = ui.prompt(
      'Starting Month',
      'Enter the month to start extra payments (MM/YYYY format):\n\nExamples: 06/2026, 12/2026, 01/2027',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (startMonthResult.getSelectedButton() !== ui.Button.OK) {
      return;
    }
    
    const startMonthInput = startMonthResult.getResponseText().trim();
    
    // Validate MM/YYYY format
    const monthYearMatch = startMonthInput.match(/^(\d{1,2})\/(\d{4})$/);
    if (!monthYearMatch) {
      ui.alert('Error', `Invalid date format: "${startMonthInput}". Please use MM/YYYY format (e.g., 06/2026).`, ui.ButtonSet.OK);
      return;
    }
    
    const startMonth = parseInt(monthYearMatch[1]);
    const startYear = parseInt(monthYearMatch[2]);
    
    if (startMonth < 1 || startMonth > 12) {
      ui.alert('Error', `Invalid month: "${startMonth}". Please enter a month between 1-12.`, ui.ButtonSet.OK);
      return;
    }
    
    const startMonthFormatted = `${String(startMonth).padStart(2, '0')}/${startYear}`;
    
    console.log('=== Extra Payment Simulation Input ===');
    console.log('Extra payment amount:', extraPayment);
    console.log('Starting month input:', startMonthInput);
    console.log('Formatted starting month:', startMonthFormatted);
    
    // Prompt for 3-paycheck month large payments
    const threePaycheckResult = ui.prompt(
      'Three-Paycheck Months',
      'Do you want to make large principal payments on months with 3 paychecks?\n\n(You get 3 paychecks roughly every 6 months)\n\nEnter "yes" or "no":',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (threePaycheckResult.getSelectedButton() !== ui.Button.OK) {
      return;
    }
    
    let threePaycheckAmount = 0;
    let threePaycheckStartMonth = null;
    const enableThreePaycheck = threePaycheckResult.getResponseText().trim().toLowerCase() === 'yes';
    
    if (enableThreePaycheck) {
      // Prompt for 3-paycheck payment amount
      const threePaycheckAmountResult = ui.prompt(
        'Three-Paycheck Payment Amount',
        'Enter the large payment amount for 3-paycheck months (e.g., 1000, 1500, 2000):',
        ui.ButtonSet.OK_CANCEL
      );
      
      if (threePaycheckAmountResult.getSelectedButton() !== ui.Button.OK) {
        return;
      }
      
      const threePaycheckAmountInput = threePaycheckAmountResult.getResponseText().trim();
      threePaycheckAmount = parseFloat(threePaycheckAmountInput.replace(/[$,]/g, ''));
      
      if (!threePaycheckAmount || isNaN(threePaycheckAmount) || threePaycheckAmount <= 0) {
        ui.alert('Error', `Invalid 3-paycheck payment amount: "${threePaycheckAmountInput}". Please enter a positive number.`, ui.ButtonSet.OK);
        return;
      }
      
      // Prompt for 3-paycheck starting month
      const threePaycheckStartResult = ui.prompt(
        'Three-Paycheck Starting Month',
        'Enter the month to start 3-paycheck payments (MM/YYYY format):\n\nNext 3-paycheck months: 11/2026, 06/2027, 11/2027, 05/2028...',
        ui.ButtonSet.OK_CANCEL
      );
      
      if (threePaycheckStartResult.getSelectedButton() !== ui.Button.OK) {
        return;
      }
      
      const threePaycheckStartInput = threePaycheckStartResult.getResponseText().trim();
      
      // Validate MM/YYYY format
      const threePaycheckMonthMatch = threePaycheckStartInput.match(/^(\d{1,2})\/(\d{4})$/);
      if (!threePaycheckMonthMatch) {
        ui.alert('Error', `Invalid date format: "${threePaycheckStartInput}". Please use MM/YYYY format (e.g., 11/2026).`, ui.ButtonSet.OK);
        return;
      }
      
      const threePaycheckStartMonthNum = parseInt(threePaycheckMonthMatch[1]);
      const threePaycheckStartYear = parseInt(threePaycheckMonthMatch[2]);
      
      if (threePaycheckStartMonthNum < 1 || threePaycheckStartMonthNum > 12) {
        ui.alert('Error', `Invalid month: "${threePaycheckStartMonthNum}". Please enter a month between 1-12.`, ui.ButtonSet.OK);
        return;
      }
      
      threePaycheckStartMonth = `${String(threePaycheckStartMonthNum).padStart(2, '0')}/${threePaycheckStartYear}`;
    }
    
    console.log('=== Three-Paycheck Payment Input ===');
    console.log('Enable 3-paycheck payments:', enableThreePaycheck);
    console.log('3-paycheck payment amount:', threePaycheckAmount);
    console.log('3-paycheck starting month:', threePaycheckStartMonth);
    
    // Read parameters from existing sheet headers (same as main function uses)
    const originDate = sheet.getRange('B6').getDisplayValue();
    const originalLoanAmount = parseFloat(String(sheet.getRange('B2').getValue()).replace(/[$,]/g, ''));
    
    // Handle annualRate - cell B4 might contain decimal (0.0275) or percentage (2.75%)
    const rateValue = sheet.getRange('B4').getValue();
    let annualRate;
    if (typeof rateValue === 'string' && rateValue.includes('%')) {
      // If it's a string like "2.75%", parse and keep as percentage
      annualRate = parseFloat(rateValue.replace('%', ''));
    } else {
      // If it's already a decimal like 0.0275, convert to percentage
      const rateNum = parseFloat(String(rateValue).replace('%', ''));
      annualRate = rateNum < 1 ? rateNum * 100 : rateNum; // Convert 0.0275 to 2.75
    }
    
    const maturityDate = sheet.getRange('B7').getDisplayValue();
    
    console.log('=== Simulation Wrapper Function Debug ===');
    console.log('Reading originalLoanAmount from cell B2:', sheet.getRange('B2').getValue());
    console.log('Parsed originalLoanAmount:', originalLoanAmount);
    console.log('Reading rate from cell B4:', sheet.getRange('B4').getValue());
    console.log('Converted annualRate:', annualRate, '(should be ~2.75)');
    console.log('originDate:', originDate);
    console.log('maturityDate:', maturityDate);
    
    // Validate parameters before proceeding
    if (!originDate || typeof originDate !== 'string' || !originDate.includes('/')) {
      throw new Error(`Invalid originDate from cell B6: "${originDate}" (type: ${typeof originDate})`);
    }
    if (!maturityDate || typeof maturityDate !== 'string' || !maturityDate.includes('/')) {
      throw new Error(`Invalid maturityDate from cell B7: "${maturityDate}" (type: ${typeof maturityDate})`);
    }
    if (!originalLoanAmount || isNaN(originalLoanAmount) || originalLoanAmount <= 0) {
      throw new Error(`Invalid originalLoanAmount from cell B2: "${originalLoanAmount}" (type: ${typeof originalLoanAmount})`);
    }
    if (!annualRate || isNaN(annualRate) || annualRate <= 0) {
      throw new Error(`Invalid annualRate from cell B4: "${annualRate}" (type: ${typeof annualRate})`);
    }
    
    console.log('All sheet parameters validated successfully');
    
    // TODO: Add more detailed console.log() for next session:
    // - Validate cell B2 contains expected numeric data and log data type
    // - Check for any parsing errors in date formats with try/catch
    // - Log intermediate calculation steps for payment date derivation
    // - Add validation checks for all input parameters (null/undefined/NaN)
    // - Log the exact cell values before any string manipulation
    // - Add error handling with detailed logging for edge cases
    // - Compare input values with main table's actual calculated values
    // - Add timestamps and execution timing measurements
    // - Log memory usage and performance metrics
    // - Add detailed mathematical validation of each calculation step
    // - Log array lengths and data structures throughout processing
    // - Add checkpoints to verify data integrity at each transformation
    
    // Extract current payment date from "Payments Made / Total" format
    const currentPaymentStr = sheet.getRange('B8').getDisplayValue();
    
    console.log('currentPaymentStr from B8:', currentPaymentStr, 'type:', typeof currentPaymentStr);
    
    // Parse the current payment date from the payments made count
    let currentPaymentDate = '03/2026'; // fallback
    const paymentsMatch = String(currentPaymentStr).match(/(\d+)\s*\/\s*(\d+)/);
    if (paymentsMatch) {
      const paymentsMade = parseInt(paymentsMatch[1]);
      console.log('Payments made from B8:', paymentsMade);
      
      // Calculate current payment date from origin date + payments made
      const [originMonth, originYear] = originDate.split('/');
      const originDateObj = new Date(parseInt(originYear), parseInt(originMonth) - 1, 1);
      const currentDateObj = new Date(originDateObj.getFullYear(), originDateObj.getMonth() + paymentsMade - 1, 1);
      const currentMonth = String(currentDateObj.getMonth() + 1).padStart(2, '0');
      const currentYear = currentDateObj.getFullYear();
      currentPaymentDate = `${currentMonth}/${currentYear}`;
    }
    
    console.log('Calculated currentPaymentDate:', currentPaymentDate);
    
    // Final validation before calling simulation function
    if (!currentPaymentDate || typeof currentPaymentDate !== 'string' || !currentPaymentDate.includes('/')) {
      throw new Error(`Invalid currentPaymentDate calculated: "${currentPaymentDate}"`);
    }
    console.log('About to call generateSimulationMortgageAmortization with:');
    console.log('  - originDate:', originDate);
    console.log('  - originalLoanAmount:', originalLoanAmount);
    console.log('  - annualRate:', annualRate);
    console.log('  - maturityDate:', maturityDate);
    console.log('  - currentPaymentDate:', currentPaymentDate);
    console.log('  - extraPayment:', extraPayment);
    console.log('  - extraPaymentStartMonth:', startMonthFormatted);
    console.log('  - threePaycheckAmount:', threePaycheckAmount);
    console.log('  - threePaycheckStartMonth:', threePaycheckStartMonth);
    
    // Call the full simulation function with extra payment parameters
    generateSimulationMortgageAmortization(originDate, originalLoanAmount, annualRate, maturityDate, currentPaymentDate, 
                                          extraPayment, startMonthFormatted, threePaycheckAmount, threePaycheckStartMonth);
    
  } catch (error) {
    console.error('Wrapper error:', error);
    SpreadsheetApp.getUi().alert('Error', 'Failed to generate simulation: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}
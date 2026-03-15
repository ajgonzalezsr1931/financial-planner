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
    .addItem('Clear All Data', 'clearAllData')
    .addToUi();
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
    
    timeline.push({
      month: month,
      payment: monthlyPayment,
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
 * Add a new credit card to the spreadsheet
 */
function addCreditCard(cardData) {
  const sheet = getOrCreateSheet('Credit Cards');
  
  // Add headers if this is the first card
  if (sheet.getLastRow() === 0) {
    const headers = ['Card Name', 'Current Balance', 'APR (%)', 'Minimum Payment', 'Target Months', 'Required Payment', 'Total Interest', 'Total Paid'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  
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
  
  sheet.appendRow(row);
  
  // Auto-resize columns
  const headers = ['Card Name', 'Current Balance', 'APR (%)', 'Minimum Payment', 'Target Months', 'Required Payment', 'Total Interest', 'Total Paid'];
  sheet.autoResizeColumns(1, headers.length);
  
  return {
    success: true,
    requiredPayment: requiredPayment,
    totalInterest: timeline.totalInterest,
    totalPaid: timeline.totalPaid
  };
}

/**
 * Generate detailed payoff schedule for a specific card
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
  addInteractiveTimeline(sheet, balance, apr, monthlyPayment, timeline.totalMonths, false); // false for regular schedule
  
  sheet.autoResizeColumns(1, headers.length);
  sheet.autoResizeColumns(7, 5); // Auto-resize timeline columns G-K
  
  return timeline;
}

/**
 * Add interactive timeline with interconnected formulas starting at G8
 */
function addInteractiveTimeline(sheet, originalBalance, apr, defaultPayment, maxMonths, isCustomSchedule = false) {
  const timelineStartCol = 7; // Column G
  const timelineStartRow = 8;
  
  // Timeline headers in G8:K8 - different label for custom schedules
  const timelineLabel = isCustomSchedule ? 'Actual Custom' : 'Actual Payoff';
  const timelineHeaders = [timelineLabel, 'Payment', 'Interest', 'Principal', 'Remaining Balance'];
  const headerRange = sheet.getRange(timelineStartRow, timelineStartCol, 1, 5);
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
      // First month: MIN(Default Payment, Balance + Interest)
      paymentFormula = `=MIN(${defaultPayment},B2+(B2*B3/100/12))`;
    } else {
      // Subsequent months: MIN(Default Payment, Previous Balance + Interest)
      const prevBalanceCell = sheet.getRange(currentRow - 1, timelineStartCol + 4).getA1Notation();
      paymentFormula = `=MIN(${defaultPayment},${prevBalanceCell}+(${prevBalanceCell}*B3/100/12))`;
    }
    paymentCell.setFormula(paymentFormula);
    paymentCell.setBackground('#E8F5E8')  // Light green to indicate it's dynamic
      .setBorder(true, true, true, true, false, false, '#4CAF50', SpreadsheetApp.BorderStyle.SOLID)
      .setNumberFormat('"$"#,##0.00');
    
    // Column I: Interest calculation formula
    let interestFormula;
    if (month === 1) {
      // First month: Interest = Starting Balance * Monthly Rate
      interestFormula = `=IF(B2<=0,0,B2*B3/100/12)`;
    } else {
      // Subsequent months: Interest = Previous Balance * Monthly Rate
      const prevBalanceCell = sheet.getRange(currentRow - 1, timelineStartCol + 4).getA1Notation();
      interestFormula = `=IF(${prevBalanceCell}<=0,0,${prevBalanceCell}*B3/100/12)`;
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
    
    // Apply alternating row colors
    const rowRange = sheet.getRange(currentRow, timelineStartCol, 1, 5);
    if (month % 2 === 0) {
      rowRange.setBackground('#FFF3E0');  // Light orange
    } else {
      rowRange.setBackground('#F3E5F5');  // Light purple
    }
    
    // Add borders for definition
    rowRange.setBorder(false, false, true, false, false, false, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);
    
    // Highlight milestone months (every 6 months) with bold formatting
    if (month % 6 === 0) {
      rowRange.setFontWeight('bold')
        .setBackground('#FFE082');  // Golden yellow for milestones
    }
  }
  
  // Add dynamic row handling for remaining balance
  addDynamicBalanceHandling(sheet, timelineStartCol, dataStartRow, monthsToAdd, originalBalance, apr);
  
  // Add conditional formatting for remaining balance to show progress
  const balanceRange = sheet.getRange(dataStartRow, timelineStartCol + 4, monthsToAdd, 1);
  
  // Add a summary section below the timeline
  const summaryRow = dataStartRow + monthsToAdd + 2;
  sheet.getRange(summaryRow, timelineStartCol, 1, 5).setValues([['Summary:', 'Total Payments', 'Total Interest', 'Final Balance', 'Payoff Month']]);
  sheet.getRange(summaryRow, timelineStartCol, 1, 5)
    .setBackground('#9C27B0')  // Purple header
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');
  
  // Summary formulas
  const paymentsRange = `${sheet.getRange(dataStartRow, timelineStartCol + 1).getA1Notation()}:${sheet.getRange(dataStartRow + monthsToAdd - 1, timelineStartCol + 1).getA1Notation()}`;
  const interestRange = `${sheet.getRange(dataStartRow, timelineStartCol + 2).getA1Notation()}:${sheet.getRange(dataStartRow + monthsToAdd - 1, timelineStartCol + 2).getA1Notation()}`;
  const balancesRange = `${sheet.getRange(dataStartRow, timelineStartCol + 4).getA1Notation()}:${sheet.getRange(dataStartRow + monthsToAdd - 1, timelineStartCol + 4).getA1Notation()}`;
  
  sheet.getRange(summaryRow + 1, timelineStartCol + 1).setFormula(`=SUM(${paymentsRange})`).setNumberFormat('"$"#,##0.00');
  sheet.getRange(summaryRow + 1, timelineStartCol + 2).setFormula(`=SUM(${interestRange})`).setNumberFormat('"$"#,##0.00');
  sheet.getRange(summaryRow + 1, timelineStartCol + 3).setFormula(`=MIN(${balancesRange})`).setNumberFormat('"$"#,##0.00');
  sheet.getRange(summaryRow + 1, timelineStartCol + 4).setFormula(`=IFERROR(MATCH(TRUE,${balancesRange}=0,0),"Not paid off")`);
  
  // Add instructions in a note with troubleshooting info
  const scheduleType = isCustomSchedule ? "CUSTOM" : "REGULAR";
  const instructionText = 
    `INTERACTIVE TIMELINE INSTRUCTIONS (${scheduleType} SCHEDULE):\\n\\n` +
    "• Payment amounts automatically adjust to prevent overpayments\\n" +
    "• Final payments are capped at remaining balance + interest\\n" +
    "• Interest, Principal, and Remaining Balance automatically recalculate\\n" +
    "• Payment column shows optimal payment amounts (auto-calculated)\\n" +
    "• All calculations update automatically when balance or APR changes\\n" +
    "• Monthly interest rate: " + (apr/12).toFixed(4) + "%\\n" +
    "• Balance reference: B2 (" + originalBalance + ")\\n" +
    "• APR reference: B3 (" + apr + "%)\\n" +
    "• Default payment: $" + defaultPayment.toFixed(2) + " (adjusted as needed)\\n" +
    (isCustomSchedule ? "• This is a CUSTOM schedule - original schedules remain unchanged\\n" : "") +
    "• Payments automatically optimize to prevent wasteful overpayments!\\n\\n" +
    "TROUBLESHOOTING: If calculations seem wrong, verify B2 contains balance and B3 contains APR percentage.";
  
  sheet.getRange(timelineStartRow, timelineStartCol).setNote(instructionText);
}

/**
 * Add dynamic row handling to extend timeline if balance remains
 */
function addDynamicBalanceHandling(sheet, timelineStartCol, dataStartRow, initialMonths, originalBalance, apr) {
  // Add a formula in a hidden area to check if additional months are needed
  const checkRow = dataStartRow + initialMonths + 10; // Place check formulas well below data
  
  // Formula to check if last balance > 0
  const lastBalanceCell = sheet.getRange(dataStartRow + initialMonths - 1, timelineStartCol + 4).getA1Notation();
  sheet.getRange(checkRow, timelineStartCol).setFormula(`=IF(${lastBalanceCell}>0.01,"EXTEND","COMPLETE")`);
  
  // Add up to 24 additional months that will only show if needed
  const maxExtraMonths = 24;
  
  for (let extraMonth = 1; extraMonth <= maxExtraMonths; extraMonth++) {
    const currentRow = dataStartRow + initialMonths + extraMonth - 1;
    const month = initialMonths + extraMonth;
    
    // Only show this row if previous balance > 0.01
    const prevBalanceCell = sheet.getRange(currentRow - 1, timelineStartCol + 4).getA1Notation();
    
    // Column G: Month number (conditional) - show month if balance remains
    sheet.getRange(currentRow, timelineStartCol).setFormula(`=IF(${prevBalanceCell}>0.01,${month},"")`);
    
    // Column H: Payment (automatic adjustment for extension months)
    const paymentCell = sheet.getRange(currentRow, timelineStartCol + 1);
    // Simplified formula to avoid nesting issues
    paymentCell.setFormula(`=IF(${prevBalanceCell}>0.01,MIN(${defaultPayment},${prevBalanceCell}+(${prevBalanceCell}*B3/100/12)),"")`);
    paymentCell.setBackground('#FFE0E0')  // Light red to indicate extra month
      .setBorder(true, true, true, true, false, false, '#FF6B6B', SpreadsheetApp.BorderStyle.SOLID)
      .setNumberFormat('"$"#,##0.00');
    
    // Column I: Interest calculation (conditional)
    const interestCell = sheet.getRange(currentRow, timelineStartCol + 2);
    interestCell.setFormula(`=IF(${prevBalanceCell}>0.01,${prevBalanceCell}*B3/100/12,"")`);
    interestCell.setNumberFormat('"$"#,##0.00');
    
    // Column J: Principal calculation (conditional)
    const paymentCellRef = paymentCell.getA1Notation();
    const interestCellRef = interestCell.getA1Notation();
    const principalCell = sheet.getRange(currentRow, timelineStartCol + 3);
    principalCell.setFormula(`=IF(${prevBalanceCell}>0.01,MIN(MAX(0,${paymentCellRef}-${interestCellRef}),${prevBalanceCell}),"")`);
    principalCell.setNumberFormat('"$"#,##0.00');
    
    // Column K: Remaining balance (conditional)
    const principalCellRef = principalCell.getA1Notation();
    const balanceFormula = `=IF(${prevBalanceCell}>0.01,MAX(0,${prevBalanceCell}-${principalCellRef}),"")`;
    const balanceCell = sheet.getRange(currentRow, timelineStartCol + 4);
    balanceCell.setFormula(balanceFormula);
    balanceCell.setNumberFormat('"$"#,##0.00');
    
    // Apply special styling for extension rows
    const rowRange = sheet.getRange(currentRow, timelineStartCol, 1, 5);
    rowRange.setBackground('#FFE8E8');  // Light red background for extension rows
    rowRange.setBorder(false, false, true, false, false, false, '#FF9999', SpreadsheetApp.BorderStyle.SOLID);
    
    // Add conditional formatting rule to hide entire row if not needed
    const hideRowRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=${prevBalanceCell}<=0.01`)
      .setFontColor('#FFFFFF')
      .setBackground('#FFFFFF')
      .build();
    
    const rules = sheet.getConditionalFormatRules();
    rules.push(hideRowRule);
    sheet.setConditionalFormatRules(rules);
  }
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
  addInteractiveTimeline(sheet, balance, apr, customPayment, timeline.totalMonths, true); // true for custom schedule
  
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
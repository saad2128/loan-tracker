/**
 * Enhanced Personal Loan Tracker for Google Sheets - Pakistan Edition
 * Supports unlimited installments and dashboard stats in sheets
 * Production-grade UI with comprehensive forms and analytics
 */

// Configuration constants
const MAIN_SHEET_NAME = 'Loan Tracker';
const PAYMENTS_SHEET_NAME = 'Payment History';
const DASHBOARD_SHEET_NAME = 'Dashboard';
const HEADER_ROW = 1;
const DATA_START_ROW = 2;

// Main sheet column indices (0-based)
const MAIN_COLS = {
  LOAN_ID: 0,          // A: Unique Loan ID
  LENDER: 1,           // B: Lender Name
  TOTAL_AMOUNT: 2,     // C: Total Loan Amount
  TOTAL_PAID: 3,       // D: Total Amount Paid
  REMAINING: 4,        // E: Remaining Amount
  PAYMENTS_COUNT: 5,   // F: Number of Payments
  LAST_PAYMENT_DATE: 6,// G: Last Payment Date
  LAST_PAYMENT_AMOUNT: 7, // H: Last Payment Amount
  EST_COMPLETION: 8,   // I: Estimated Completion
  LOAN_PURPOSE: 9,     // J: Loan Purpose
  STATUS: 10,          // K: Status (Active/Paid/Overdue)
  NOTES: 11            // L: Notes
};

// Payment history column indices
const PAYMENT_COLS = {
  PAYMENT_ID: 0,       // A: Payment ID
  LOAN_ID: 1,          // B: Loan ID (Reference)
  LENDER: 2,           // C: Lender Name
  PAYMENT_DATE: 3,     // D: Payment Date
  PAYMENT_AMOUNT: 4,   // E: Payment Amount
  REMAINING_AFTER: 5,  // F: Remaining After Payment
  PAYMENT_METHOD: 6,   // G: Payment Method
  NOTES: 7             // H: Payment Notes
};

/**
 * Creates custom menu when spreadsheet opens
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('💰 Enhanced Loan Tracker')
    .addItem('🏗️ Initialize Loan Tracker', 'initializeEnhancedLoanTracker')
    .addSeparator()
    .addItem('➕ Add New Loan', 'showAddLoanForm')
    .addItem('💸 Record Payment', 'showPaymentForm')
    .addSeparator()
    .addItem('📊 Refresh Dashboard', 'refreshDashboard')
    .addItem('📈 View Payment History', 'showPaymentHistory')
    .addSeparator()
    .addItem('🔄 Refresh All Calculations', 'refreshAllCalculations')
    .addItem('❓ Help & Format Guide', 'showHelpDialog')
    .addToUi();
}

/**
 * Initialize the enhanced loan tracker with multiple sheets
 */
function initializeEnhancedLoanTracker() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Initialize main loan tracker sheet
  initializeMainSheet(ss);
  
  // Initialize payment history sheet
  initializePaymentHistorySheet(ss);
  
  // Initialize dashboard sheet
  initializeDashboardSheet(ss);
  
  // Create initial dashboard
  refreshDashboard();
  
  SpreadsheetApp.getUi().alert('✅ Success', 'Enhanced Loan Tracker initialized successfully!\n\n• Main tracker with unlimited payments\n• Payment history log\n• Live dashboard with stats\n\nUse the menu options to get started.', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Initialize main loan tracker sheet
 */
function initializeMainSheet(ss) {
  let sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(MAIN_SHEET_NAME);
  }
  
  sheet.clear();
  
  const headers = [
    'Loan ID',
    'Lender Name',
    'Total Loan Amount (PKR)',
    'Total Paid',
    'Remaining Amount',
    'Payments Count',
    'Last Payment Date',
    'Last Payment Amount',
    'Estimated Completion',
    'Loan Purpose',
    'Status',
    'Notes'
  ];
  
  sheet.getRange(HEADER_ROW, 1, 1, headers.length).setValues([headers]);
  
  // Format headers
  const headerRange = sheet.getRange(HEADER_ROW, 1, 1, headers.length);
  headerRange.setBackground('#1a73e8')
           .setFontColor('white')
           .setFontWeight('bold')
           .setFontSize(11)
           .setHorizontalAlignment('center')
           .setVerticalAlignment('middle');
  
  // Set column widths
  const columnWidths = [80, 150, 180, 150, 150, 100, 130, 150, 180, 120, 100, 200];
  columnWidths.forEach((width, index) => {
    sheet.setColumnWidth(index + 1, width);
  });
  
  sheet.setFrozenRows(1);
}

/**
 * Initialize payment history sheet
 */
function initializePaymentHistorySheet(ss) {
  let sheet = ss.getSheetByName(PAYMENTS_SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(PAYMENTS_SHEET_NAME);
  }
  
  sheet.clear();
  
  const headers = [
    'Payment ID',
    'Loan ID',
    'Lender Name',
    'Payment Date',
    'Payment Amount (PKR)',
    'Remaining After Payment',
    'Payment Method',
    'Notes'
  ];
  
  sheet.getRange(HEADER_ROW, 1, 1, headers.length).setValues([headers]);
  
  // Format headers
  const headerRange = sheet.getRange(HEADER_ROW, 1, 1, headers.length);
  headerRange.setBackground('#34a853')
           .setFontColor('white')
           .setFontWeight('bold')
           .setFontSize(11)
           .setHorizontalAlignment('center');
  
  // Set column widths
  const columnWidths = [100, 80, 150, 120, 150, 150, 120, 200];
  columnWidths.forEach((width, index) => {
    sheet.setColumnWidth(index + 1, width);
  });
  
  sheet.setFrozenRows(1);
}

/**
 * Initialize dashboard sheet
 */
function initializeDashboardSheet(ss) {
  let sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(DASHBOARD_SHEET_NAME);
  }
  
  sheet.clear();
  
  // Dashboard will be populated by refreshDashboard function
}

/**
 * Show enhanced add loan form
 */
function showAddLoanForm() {
  const htmlOutput = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body {
            font-family: 'Google Sans', -apple-system, BlinkMacSystemFont, sans-serif;
            margin: 0;
            padding: 20px;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
          }
          .container {
            max-width: 500px;
            margin: 0 auto;
            background: white;
            border-radius: 16px;
            box-shadow: 0 10px 40px rgba(0,0,0,0.1);
            overflow: hidden;
          }
          .header {
            background: linear-gradient(135deg, #1a73e8, #4285f4);
            color: white;
            padding: 24px;
            text-align: center;
          }
          .header h2 {
            margin: 0;
            font-size: 24px;
            font-weight: 500;
          }
          .form-content {
            padding: 32px;
          }
          .form-group {
            margin-bottom: 24px;
          }
          label {
            display: block;
            margin-bottom: 8px;
            font-weight: 500;
            color: #333;
            font-size: 14px;
          }
          input, textarea, select {
            width: 100%;
            padding: 12px 16px;
            border: 2px solid #e8eaed;
            border-radius: 8px;
            font-size: 14px;
            transition: border-color 0.2s;
            box-sizing: border-box;
          }
          input:focus, textarea:focus, select:focus {
            outline: none;
            border-color: #1a73e8;
            box-shadow: 0 0 0 3px rgba(26, 115, 232, 0.1);
          }
          .amount-help {
            background: #f8f9fa;
            border: 1px solid #e8eaed;
            border-radius: 8px;
            padding: 12px;
            margin-top: 8px;
            font-size: 12px;
            color: #5f6368;
          }
          .amount-help strong {
            color: #1a73e8;
          }
          .button-group {
            display: flex;
            gap: 12px;
            margin-top: 32px;
          }
          button {
            flex: 1;
            padding: 12px 24px;
            border: none;
            border-radius: 8px;
            font-size: 14px;
            font-weight: 500;
            cursor: pointer;
            transition: all 0.2s;
          }
          .btn-primary {
            background: #1a73e8;
            color: white;
          }
          .btn-primary:hover {
            background: #1557b0;
            transform: translateY(-1px);
            box-shadow: 0 4px 12px rgba(26, 115, 232, 0.3);
          }
          .btn-secondary {
            background: #f8f9fa;
            color: #5f6368;
            border: 1px solid #dadce0;
          }
          .btn-secondary:hover {
            background: #e8eaed;
          }
          .error {
            color: #d93025;
            font-size: 12px;
            margin-top: 4px;
            display: none;
          }
        </style>
      </head>
      <body>
        <div class="container">
          <div class="header">
            <h2>➕ Add New Loan</h2>
          </div>
          <div class="form-content">
            <form id="loanForm">
              <div class="form-group">
                <label for="lenderName">Lender Name *</label>
                <input type="text" id="lenderName" name="lenderName" required placeholder="e.g., Ahmed Ali, Bank Alfalah, Family Friend">
                <div class="error" id="lenderError">Please enter a valid lender name</div>
              </div>
              
              <div class="form-group">
                <label for="loanAmount">Total Loan Amount *</label>
                <input type="text" id="loanAmount" name="loanAmount" required placeholder="e.g., 5 lac, 50k, 2.5 lac, 150000">
                <div class="amount-help">
                  <strong>Supported formats:</strong><br>
                  • <strong>5 lac</strong> = ₨5,00,000<br>
                  • <strong>50k</strong> = ₨50,000<br>
                  • <strong>2.5 lac</strong> = ₨2,50,000<br>
                  • <strong>150000</strong> = ₨1,50,000
                </div>
                <div class="error" id="amountError">Please enter a valid amount</div>
              </div>
              
              <div class="form-group">
                <label for="loanPurpose">Loan Purpose</label>
                <select id="loanPurpose" name="loanPurpose">
                  <option value="">Select purpose...</option>
                  <option value="Business">Business Investment</option>
                  <option value="Education">Education</option>
                  <option value="Medical">Medical Emergency</option>
                  <option value="Personal">Personal Use</option>
                  <option value="Property">Property/Real Estate</option>
                  <option value="Vehicle">Vehicle Purchase</option>
                  <option value="Wedding">Wedding</option>
                  <option value="Travel">Travel</option>
                  <option value="Emergency">Emergency</option>
                  <option value="Other">Other</option>
                </select>
              </div>
              
              <div class="form-group">
                <label for="expectedMonthlyPayment">Expected Monthly Payment (Optional)</label>
                <input type="text" id="expectedMonthlyPayment" name="expectedMonthlyPayment" placeholder="e.g., 25k, 50000">
                <div class="amount-help">
                  This helps estimate repayment timeline. Same format as loan amount.
                </div>
              </div>
              
              <div class="form-group">
                <label for="notes">Additional Notes (Optional)</label>
                <textarea id="notes" name="notes" rows="3" placeholder="Any additional information about the loan..."></textarea>
              </div>
              
              <div class="button-group">
                <button type="button" class="btn-secondary" onclick="google.script.host.close()">Cancel</button>
                <button type="submit" class="btn-primary">Add Loan</button>
              </div>
            </form>
          </div>
        </div>
        
        <script>
          document.getElementById('loanForm').addEventListener('submit', function(e) {
            e.preventDefault();
            
            const formData = {
              lenderName: document.getElementById('lenderName').value.trim(),
              loanAmount: document.getElementById('loanAmount').value.trim(),
              loanPurpose: document.getElementById('loanPurpose').value,
              expectedMonthlyPayment: document.getElementById('expectedMonthlyPayment').value.trim(),
              notes: document.getElementById('notes').value.trim()
            };
            
            // Validate required fields
            let isValid = true;
            
            if (!formData.lenderName) {
              document.getElementById('lenderError').style.display = 'block';
              isValid = false;
            } else {
              document.getElementById('lenderError').style.display = 'none';
            }
            
            if (!formData.loanAmount) {
              document.getElementById('amountError').style.display = 'block';
              isValid = false;
            } else {
              document.getElementById('amountError').style.display = 'none';
            }
            
            if (isValid) {
              google.script.run
                .withSuccessHandler(function(result) {
                  if (result.success) {
                    alert('✅ Loan added successfully!\\n\\nLender: ' + result.lenderName + '\\nAmount: ₨' + result.formattedAmount + '\\nLoan ID: ' + result.loanId);
                    google.script.host.close();
                  } else {
                    alert('❌ Error: ' + result.error);
                  }
                })
                .withFailureHandler(function(error) {
                  alert('❌ Error: ' + error.message);
                })
                .processAddLoanForm(formData);
            }
          });
        </script>
      </body>
    </html>
  `).setWidth(600).setHeight(700);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Add New Loan');
}

/**
 * Show enhanced payment form with loan selection
 */
function showPaymentForm() {
  const mainSheet = getOrCreateMainSheet();
  const loans = getActiveLoans(mainSheet);
  
  if (loans.length === 0) {
    SpreadsheetApp.getUi().alert('❌ No Active Loans', 'No active loans found. Please add a loan first!', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  const loanOptions = loans.map((loan, index) => 
    `<option value="${loan.loanId}">${loan.lender} - Remaining: ₨${formatPKR(loan.remaining)} (${loan.paymentsCount} payments)</option>`
  ).join('');
  
  const htmlOutput = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body {
            font-family: 'Google Sans', -apple-system, BlinkMacSystemFont, sans-serif;
            margin: 0;
            padding: 20px;
            background: linear-gradient(135deg, #ff6b6b 0%, #ee5a24 100%);
            min-height: 100vh;
          }
          .container {
            max-width: 500px;
            margin: 0 auto;
            background: white;
            border-radius: 16px;
            box-shadow: 0 10px 40px rgba(0,0,0,0.1);
            overflow: hidden;
          }
          .header {
            background: linear-gradient(135deg, #ff6b6b, #ee5a24);
            color: white;
            padding: 24px;
            text-align: center;
          }
          .header h2 {
            margin: 0;
            font-size: 24px;
            font-weight: 500;
          }
          .form-content {
            padding: 32px;
          }
          .form-group {
            margin-bottom: 24px;
          }
          label {
            display: block;
            margin-bottom: 8px;
            font-weight: 500;
            color: #333;
            font-size: 14px;
          }
          select, input, textarea {
            width: 100%;
            padding: 12px 16px;
            border: 2px solid #e8eaed;
            border-radius: 8px;
            font-size: 14px;
            transition: border-color 0.2s;
            box-sizing: border-box;
          }
          select:focus, input:focus, textarea:focus {
            outline: none;
            border-color: #ff6b6b;
            box-shadow: 0 0 0 3px rgba(255, 107, 107, 0.1);
          }
          .loan-info {
            background: #f8f9fa;
            border-left: 4px solid #ff6b6b;
            padding: 16px;
            margin: 16px 0;
            border-radius: 0 8px 8px 0;
          }
          .amount-help {
            background: #f8f9fa;
            border: 1px solid #e8eaed;
            border-radius: 8px;
            padding: 12px;
            margin-top: 8px;
            font-size: 12px;
            color: #5f6368;
          }
          .amount-help strong {
            color: #ff6b6b;
          }
          .button-group {
            display: flex;
            gap: 12px;
            margin-top: 32px;
          }
          button {
            flex: 1;
            padding: 12px 24px;
            border: none;
            border-radius: 8px;
            font-size: 14px;
            font-weight: 500;
            cursor: pointer;
            transition: all 0.2s;
          }
          .btn-primary {
            background: #ff6b6b;
            color: white;
          }
          .btn-primary:hover {
            background: #ee5a24;
            transform: translateY(-1px);
            box-shadow: 0 4px 12px rgba(255, 107, 107, 0.3);
          }
          .btn-secondary {
            background: #f8f9fa;
            color: #5f6368;
            border: 1px solid #dadce0;
          }
          .btn-secondary:hover {
            background: #e8eaed;
          }
          .error {
            color: #d93025;
            font-size: 12px;
            margin-top: 4px;
            display: none;
          }
        </style>
      </head>
      <body>
        <div class="container">
          <div class="header">
            <h2>💸 Record Payment</h2>
          </div>
          <div class="form-content">
            <form id="paymentForm">
              <div class="form-group">
                <label for="loanSelect">Select Loan *</label>
                <select id="loanSelect" name="loanSelect" required onchange="updateLoanInfo()">
                  <option value="">Choose a loan...</option>
                  ${loanOptions}
                </select>
                <div class="error" id="loanError">Please select a loan</div>
              </div>
              
              <div id="loanInfo" class="loan-info" style="display: none;">
                <div id="loanDetails"></div>
              </div>
              
              <div class="form-group">
                <label for="paymentAmount">Payment Amount *</label>
                <input type="text" id="paymentAmount" name="paymentAmount" required placeholder="e.g., 25k, 50000, 1 lac">
                <div class="amount-help">
                  <strong>Supported formats:</strong> 25k, 1 lac, 2.5 lac, 50000
                </div>
                <div class="error" id="amountError">Please enter a valid payment amount</div>
              </div>
              
              <div class="form-group">
                <label for="paymentDate">Payment Date</label>
                <input type="date" id="paymentDate" name="paymentDate">
              </div>
              
              <div class="form-group">
                <label for="paymentMethod">Payment Method</label>
                <select id="paymentMethod" name="paymentMethod">
                  <option value="">Select method...</option>
                  <option value="Cash">Cash</option>
                  <option value="Bank Transfer">Bank Transfer</option>
                  <option value="Mobile Banking">Mobile Banking (JazzCash/Easypaisa)</option>
                  <option value="Cheque">Cheque</option>
                  <option value="Online Banking">Online Banking</option>
                  <option value="Other">Other</option>
                </select>
              </div>
              
              <div class="form-group">
                <label for="paymentNotes">Payment Notes (Optional)</label>
                <textarea id="paymentNotes" name="paymentNotes" rows="3" placeholder="Any notes about this payment..."></textarea>
              </div>
              
              <div class="button-group">
                <button type="button" class="btn-secondary" onclick="google.script.host.close()">Cancel</button>
                <button type="submit" class="btn-primary">Record Payment</button>
              </div>
            </form>
          </div>
        </div>
        
        <script>
          const loans = ${JSON.stringify(loans)};
          
          // Set today's date as default
          document.getElementById('paymentDate').valueAsDate = new Date();
          
          function updateLoanInfo() {
            const loanId = document.getElementById('loanSelect').value;
            const loanInfo = document.getElementById('loanInfo');
            const loanDetails = document.getElementById('loanDetails');
            
            if (loanId !== '') {
              const loan = loans.find(l => l.loanId === loanId);
              if (loan) {
                loanDetails.innerHTML = \`
                  <strong>\${loan.lender}</strong><br>
                  Total Loan: ₨\${formatPKR(loan.totalAmount)}<br>
                  Total Paid: ₨\${formatPKR(loan.totalPaid)}<br>
                  Remaining: ₨\${formatPKR(loan.remaining)}<br>
                  Payments Made: \${loan.paymentsCount}
                \`;
                loanInfo.style.display = 'block';
              }
            } else {
              loanInfo.style.display = 'none';
            }
          }
          
          function formatPKR(amount) {
            if (amount >= 10000000) {
              return (amount / 10000000).toFixed(1).replace('.0', '') + ' Cr';
            } else if (amount >= 100000) {
              return (amount / 100000).toFixed(1).replace('.0', '') + ' Lac';
            } else if (amount >= 1000) {
              return (amount / 1000).toFixed(0) + 'k';
            } else {
              return new Intl.NumberFormat('en-PK').format(amount);
            }
          }
          
          document.getElementById('paymentForm').addEventListener('submit', function(e) {
            e.preventDefault();
            
            const formData = {
              loanId: document.getElementById('loanSelect').value,
              paymentAmount: document.getElementById('paymentAmount').value.trim(),
              paymentDate: document.getElementById('paymentDate').value,
              paymentMethod: document.getElementById('paymentMethod').value,
              paymentNotes: document.getElementById('paymentNotes').value.trim()
            };
            
            // Validate required fields
            let isValid = true;
            
            if (!formData.loanId) {
              document.getElementById('loanError').style.display = 'block';
              isValid = false;
            } else {
              document.getElementById('loanError').style.display = 'none';
            }
            
            if (!formData.paymentAmount) {
              document.getElementById('amountError').style.display = 'block';
              isValid = false;
            } else {
              document.getElementById('amountError').style.display = 'none';
            }
            
            if (isValid) {
              google.script.run
                .withSuccessHandler(function(result) {
                  if (result.success) {
                    alert('✅ Payment recorded successfully!\\n\\nAmount: ₨' + result.formattedAmount + '\\nRemaining: ₨' + result.newRemaining + '\\nPayment #' + result.paymentNumber);
                    google.script.host.close();
                  } else {
                    alert('❌ Error: ' + result.error);
                  }
                })
                .withFailureHandler(function(error) {
                  alert('❌ Error: ' + error.message);
                })
                .processPaymentForm(formData);
            }
          });
        </script>
      </body>
    </html>
  `).setWidth(600).setHeight(750);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Record Payment');
}

/**
 * Process add loan form submission
 */
function processAddLoanForm(formData) {
  try {
    const mainSheet = getOrCreateMainSheet();
    
    // Parse amount with flexible format
    const parsedAmount = parseFlexibleAmount(formData.loanAmount);
    if (parsedAmount === null) {
      return { success: false, error: 'Invalid amount format. Use formats like: 5 lac, 50k, 150000' };
    }
    
    // Generate unique loan ID
    const loanId = generateLoanId();
    
    const lastRow = mainSheet.getLastRow();
    const newRow = lastRow + 1;
    
    // Add new loan entry
    const rowData = [
      loanId,                           // Loan ID
      formData.lenderName,              // Lender Name
      parsedAmount,                     // Total Amount
      0,                                // Total Paid
      parsedAmount,                     // Remaining
      0,                                // Payments Count
      '',                               // Last Payment Date
      '',                               // Last Payment Amount
      formData.expectedMonthlyPayment ? 
        calculateEstimatedCompletion(parsedAmount, parseFlexibleAmount(formData.expectedMonthlyPayment)) : '', // Est Completion
      formData.loanPurpose || '',       // Purpose
      'Active',                         // Status
      formData.notes || ''              // Notes
    ];
    
    mainSheet.getRange(newRow, 1, 1, rowData.length).setValues([rowData]);
    
    // Format the row
    formatMainSheetRow(mainSheet, newRow);
    
    // Refresh dashboard
    refreshDashboard();
    
    return {
      success: true,
      lenderName: formData.lenderName,
      formattedAmount: formatPKR(parsedAmount),
      loanId: loanId
    };
    
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

/**
 * Process payment form submission
 */
function processPaymentForm(formData) {
  try {
    const mainSheet = getOrCreateMainSheet();
    const paymentSheet = getOrCreatePaymentHistorySheet();
    
    // Parse payment amount
    const paymentAmount = parseFlexibleAmount(formData.paymentAmount);
    if (paymentAmount === null) {
      return { success: false, error: 'Invalid payment amount format' };
    }
    
    // Find the loan row
    const loanRow = findLoanRow(mainSheet, formData.loanId);
    if (!loanRow) {
      return { success: false, error: 'Loan not found' };
    }
    
    // Get current loan data
    const loanData = mainSheet.getRange(loanRow, 1, 1, 12).getValues()[0];
    const currentRemaining = parseFloat(loanData[MAIN_COLS.REMAINING]) || 0;
    const currentTotalPaid = parseFloat(loanData[MAIN_COLS.TOTAL_PAID]) || 0;
    const currentPaymentsCount = parseInt(loanData[MAIN_COLS.PAYMENTS_COUNT]) || 0;
    
    // Calculate new values
    const newTotalPaid = currentTotalPaid + paymentAmount;
    const newRemaining = Math.max(0, currentRemaining - paymentAmount);
    const newPaymentsCount = currentPaymentsCount + 1;
    const paymentDate = formData.paymentDate ? new Date(formData.paymentDate) : new Date();
    
    // Update main sheet
    mainSheet.getRange(loanRow, MAIN_COLS.TOTAL_PAID + 1).setValue(newTotalPaid);
    mainSheet.getRange(loanRow, MAIN_COLS.REMAINING + 1).setValue(newRemaining);
    mainSheet.getRange(loanRow, MAIN_COLS.PAYMENTS_COUNT + 1).setValue(newPaymentsCount);
    mainSheet.getRange(loanRow, MAIN_COLS.LAST_PAYMENT_DATE + 1).setValue(paymentDate);
    mainSheet.getRange(loanRow, MAIN_COLS.LAST_PAYMENT_AMOUNT + 1).setValue(paymentAmount);
    
    // Update status
    const status = newRemaining <= 0 ? 'Paid' : 'Active';
    mainSheet.getRange(loanRow, MAIN_COLS.STATUS + 1).setValue(status);
    
    // Add payment to history sheet
    const paymentId = generatePaymentId();
    const lastPaymentRow = paymentSheet.getLastRow();
    const newPaymentRow = lastPaymentRow + 1;
    
    const paymentData = [
      paymentId,                        // Payment ID
      formData.loanId,                  // Loan ID
      loanData[MAIN_COLS.LENDER],       // Lender Name
      paymentDate,                      // Payment Date
      paymentAmount,                    // Payment Amount
      newRemaining,                     // Remaining After Payment
      formData.paymentMethod || '',     // Payment Method
      formData.paymentNotes || ''       // Notes
    ];
    
    paymentSheet.getRange(newPaymentRow, 1, 1, paymentData.length).setValues([paymentData]);
    
    // Format rows
    formatMainSheetRow(mainSheet, loanRow);
    formatPaymentHistoryRow(paymentSheet, newPaymentRow);
    
    // Refresh dashboard
    refreshDashboard();
    
    return {
      success: true,
      formattedAmount: formatPKR(paymentAmount),
      newRemaining: formatPKR(newRemaining),
      paymentNumber: newPaymentsCount
    };
    
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

/**
 * Refresh the dashboard with current statistics
 */
function refreshDashboard() {
  try {
    const dashboardSheet = getOrCreateDashboardSheet();
    const mainSheet = getOrCreateMainSheet();
    const paymentSheet = getOrCreatePaymentHistorySheet();
    
    dashboardSheet.clear();
    
    // Get loan data
    const loans = getAllLoans(mainSheet);
    const payments = getAllPayments(paymentSheet);
    
    // Calculate statistics
    const stats = calculateDashboardStats(loans, payments);
    
    // Create dashboard layout
    createDashboardLayout(dashboardSheet, stats, loans, payments);
    
    return { success: true };
    
  } catch (error) {
    console.log('Dashboard refresh error:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * Calculate comprehensive dashboard statistics
 */
function calculateDashboardStats(loans, payments) {
  const totalLoans = loans.length;
  const activeLoans = loans.filter(loan => loan.status === 'Active').length;
  const paidLoans = loans.filter(loan => loan.status === 'Paid').length;
  
  const totalBorrowed = loans.reduce((sum, loan) => sum + loan.totalAmount, 0);
  const totalPaid = loans.reduce((sum, loan) => sum + loan.totalPaid, 0);
  const totalRemaining = loans.reduce((sum, loan) => sum + loan.remaining, 0);
  
  const totalPayments = payments.length;
  const avgPaymentAmount = totalPayments > 0 ? totalPaid / totalPayments : 0;
  
  // Payment frequency analysis
  const paymentDates = payments.map(p => new Date(p.paymentDate)).filter(d => d instanceof Date && !isNaN(d));
  const recentPayments = paymentDates.filter(date => {
    const thirtyDaysAgo = new Date();
    thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);
    return date >= thirtyDaysAgo;
  }).length;
  
  // Progress percentage
  const progressPercentage = totalBorrowed > 0 ? (totalPaid / totalBorrowed) * 100 : 0;
  
  // Loan purpose breakdown
  const purposeBreakdown = {};
  loans.forEach(loan => {
    const purpose = loan.purpose || 'Not Specified';
    purposeBreakdown[purpose] = (purposeBreakdown[purpose] || 0) + loan.totalAmount;
  });
  
  return {
    totalLoans,
    activeLoans,
    paidLoans,
    totalBorrowed,
    totalPaid,
    totalRemaining,
    totalPayments,
    avgPaymentAmount,
    recentPayments,
    progressPercentage,
    purposeBreakdown
  };
}

/**
 * Create comprehensive dashboard layout
 */
function createDashboardLayout(sheet, stats, loans, payments) {
  let currentRow = 1;
  
  // Dashboard Header
  sheet.getRange(currentRow, 1, 1, 6).merge();
  sheet.getRange(currentRow, 1).setValue('📊 LOAN TRACKER DASHBOARD');
  sheet.getRange(currentRow, 1).setBackground('#1a73e8').setFontColor('white').setFontWeight('bold').setFontSize(16).setHorizontalAlignment('center');
  currentRow += 2;
  
  // Date and time
  sheet.getRange(currentRow, 1, 1, 6).merge();
  sheet.getRange(currentRow, 1).setValue(`Last Updated: ${new Date().toLocaleString('en-PK')}`);
  sheet.getRange(currentRow, 1).setFontStyle('italic').setHorizontalAlignment('center');
  currentRow += 2;
  
  // Key Statistics Section
  sheet.getRange(currentRow, 1).setValue('🔢 KEY STATISTICS');
  sheet.getRange(currentRow, 1).setBackground('#34a853').setFontColor('white').setFontWeight('bold').setFontSize(14);
  currentRow++;
  
  // Stats grid
  const statsData = [
    ['Total Loans', stats.totalLoans, 'Active Loans', stats.activeLoans],
    ['Paid Loans', stats.paidLoans, 'Total Payments', stats.totalPayments],
    ['Total Borrowed', `₨${formatPKR(stats.totalBorrowed)}`, 'Total Paid', `₨${formatPKR(stats.totalPaid)}`],
    ['Total Remaining', `₨${formatPKR(stats.totalRemaining)}`, 'Avg Payment', `₨${formatPKR(stats.avgPaymentAmount)}`],
    ['Progress', `${stats.progressPercentage.toFixed(1)}%`, 'Recent Payments (30d)', stats.recentPayments]
  ];
  
  sheet.getRange(currentRow, 1, statsData.length, 4).setValues(statsData);
  
  // Format stats section
  for (let i = 0; i < statsData.length; i++) {
    const row = currentRow + i;
    sheet.getRange(row, 1).setBackground('#f0f0f0').setFontWeight('bold');
    sheet.getRange(row, 3).setBackground('#f0f0f0').setFontWeight('bold');
    sheet.getRange(row, 2).setHorizontalAlignment('right');
    sheet.getRange(row, 4).setHorizontalAlignment('right');
  }
  
  currentRow += statsData.length + 2;
  
  // Loan Purpose Breakdown
  sheet.getRange(currentRow, 1).setValue('📈 LOAN PURPOSE BREAKDOWN');
  sheet.getRange(currentRow, 1).setBackground('#ff9800').setFontColor('white').setFontWeight('bold').setFontSize(14);
  currentRow++;
  
  const purposeHeaders = ['Purpose', 'Amount', 'Percentage'];
  sheet.getRange(currentRow, 1, 1, purposeHeaders.length).setValues([purposeHeaders]);
  sheet.getRange(currentRow, 1, 1, purposeHeaders.length).setBackground('#fff3e0').setFontWeight('bold');
  currentRow++;
  
  Object.entries(stats.purposeBreakdown).forEach(([purpose, amount]) => {
    const percentage = ((amount / stats.totalBorrowed) * 100).toFixed(1);
    sheet.getRange(currentRow, 1, 1, 3).setValues([[purpose, `₨${formatPKR(amount)}`, `${percentage}%`]]);
    sheet.getRange(currentRow, 2).setHorizontalAlignment('right');
    sheet.getRange(currentRow, 3).setHorizontalAlignment('right');
    currentRow++;
  });
  
  currentRow += 2;
  
  // Active Loans Summary
  const activeLoans = loans.filter(loan => loan.status === 'Active');
  if (activeLoans.length > 0) {
    sheet.getRange(currentRow, 1).setValue('🏃 ACTIVE LOANS SUMMARY');
    sheet.getRange(currentRow, 1).setBackground('#e91e63').setFontColor('white').setFontWeight('bold').setFontSize(14);
    currentRow++;
    
    const activeLoanHeaders = ['Lender', 'Total Amount', 'Paid', 'Remaining', 'Payments', 'Last Payment'];
    sheet.getRange(currentRow, 1, 1, activeLoanHeaders.length).setValues([activeLoanHeaders]);
    sheet.getRange(currentRow, 1, 1, activeLoanHeaders.length).setBackground('#fce4ec').setFontWeight('bold');
    currentRow++;
    
    activeLoans.forEach(loan => {
      const activeLoanData = [
        loan.lender,
        `₨${formatPKR(loan.totalAmount)}`,
        `₨${formatPKR(loan.totalPaid)}`,
        `₨${formatPKR(loan.remaining)}`,
        loan.paymentsCount,
        loan.lastPaymentDate ? new Date(loan.lastPaymentDate).toLocaleDateString('en-PK') : 'None'
      ];
      sheet.getRange(currentRow, 1, 1, activeLoanData.length).setValues([activeLoanData]);
      
      // Color code by remaining amount
      if (loan.remaining > loan.totalAmount * 0.7) {
        sheet.getRange(currentRow, 1, 1, activeLoanData.length).setBackground('#ffebee'); // High remaining
      } else if (loan.remaining > loan.totalAmount * 0.3) {
        sheet.getRange(currentRow, 1, 1, activeLoanData.length).setBackground('#fff8e1'); // Medium remaining
      } else {
        sheet.getRange(currentRow, 1, 1, activeLoanData.length).setBackground('#e8f5e8'); // Low remaining
      }
      
      currentRow++;
    });
    
    currentRow += 2;
  }
  
  // Recent Payments Summary
  if (payments.length > 0) {
    sheet.getRange(currentRow, 1).setValue('💸 RECENT PAYMENTS (Last 10)');
    sheet.getRange(currentRow, 1).setBackground('#9c27b0').setFontColor('white').setFontWeight('bold').setFontSize(14);
    currentRow++;
    
    const recentPaymentHeaders = ['Date', 'Lender', 'Amount', 'Method', 'Remaining After'];
    sheet.getRange(currentRow, 1, 1, recentPaymentHeaders.length).setValues([recentPaymentHeaders]);
    sheet.getRange(currentRow, 1, 1, recentPaymentHeaders.length).setBackground('#f3e5f5').setFontWeight('bold');
    currentRow++;
    
    // Sort payments by date (most recent first) and take last 10
    const sortedPayments = payments.sort((a, b) => new Date(b.paymentDate) - new Date(a.paymentDate)).slice(0, 10);
    
    sortedPayments.forEach(payment => {
      const paymentData = [
        new Date(payment.paymentDate).toLocaleDateString('en-PK'),
        payment.lenderName,
        `₨${formatPKR(payment.paymentAmount)}`,
        payment.paymentMethod || 'Not specified',
        `₨${formatPKR(payment.remainingAfter)}`
      ];
      sheet.getRange(currentRow, 1, 1, paymentData.length).setValues([paymentData]);
      currentRow++;
    });
  }
  
  // Set column widths for better display
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 150);
  sheet.setColumnWidth(4, 150);
  sheet.setColumnWidth(5, 100);
  sheet.setColumnWidth(6, 150);
  
  // Auto-resize to fit content
  sheet.autoResizeColumns(1, 6);
}

/**
 * Show payment history in a dialog
 */
function showPaymentHistory() {
  const paymentSheet = getOrCreatePaymentHistorySheet();
  const payments = getAllPayments(paymentSheet);
  
  if (payments.length === 0) {
    SpreadsheetApp.getUi().alert('📈 Payment History', 'No payments recorded yet.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  // Sort by date (most recent first)
  const sortedPayments = payments.sort((a, b) => new Date(b.paymentDate) - new Date(a.paymentDate));
  
  const historyHtml = `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body { font-family: 'Google Sans', sans-serif; margin: 0; padding: 20px; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); }
          .container { max-width: 900px; margin: 0 auto; background: white; border-radius: 16px; box-shadow: 0 10px 40px rgba(0,0,0,0.1); overflow: hidden; }
          .header { background: linear-gradient(135deg, #667eea, #764ba2); color: white; padding: 24px; text-align: center; }
          .header h2 { margin: 0; font-size: 24px; font-weight: 500; }
          .content { padding: 32px; max-height: 600px; overflow-y: auto; }
          .payment-item { background: #f8f9fa; border-radius: 8px; padding: 16px; margin-bottom: 12px; border-left: 4px solid #667eea; }
          .payment-header { font-weight: bold; margin-bottom: 8px; color: #333; display: flex; justify-content: space-between; }
          .payment-details { font-size: 14px; color: #666; line-height: 1.4; }
          .amount { color: #1a73e8; font-weight: bold; }
          .close-btn { background: #667eea; color: white; border: none; padding: 12px 24px; border-radius: 8px; font-size: 14px; cursor: pointer; margin: 20px auto; display: block; }
        </style>
      </head>
      <body>
        <div class="container">
          <div class="header">
            <h2>📈 Payment History</h2>
          </div>
          <div class="content">
            ${sortedPayments.map(payment => `
              <div class="payment-item">
                <div class="payment-header">
                  <span>${payment.lenderName}</span>
                  <span class="amount">₨${formatPKR(payment.paymentAmount)}</span>
                </div>
                <div class="payment-details">
                  Date: ${new Date(payment.paymentDate).toLocaleDateString('en-PK')} | 
                  Method: ${payment.paymentMethod || 'Not specified'} | 
                  Remaining After: ₨${formatPKR(payment.remainingAfter)}
                  ${payment.notes ? `<br>Notes: ${payment.notes}` : ''}
                </div>
              </div>
            `).join('')}
          </div>
          <button class="close-btn" onclick="google.script.host.close()">Close</button>
        </div>
      </body>
    </html>
  `;
  
  const htmlOutput = HtmlService.createHtmlOutput(historyHtml).setWidth(950).setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Payment History');
}

/**
 * Helper function to parse flexible amount formats
 */
function parseFlexibleAmount(amountStr) {
  if (!amountStr || typeof amountStr !== 'string') return null;
  
  let cleaned = amountStr.toLowerCase().trim().replace(/,/g, '').replace(/₨/g, '').replace(/rs/g, '');
  
  if (cleaned.includes('lac') || cleaned.includes('lakh')) {
    const number = parseFloat(cleaned.replace(/lac|lakh/g, '').trim());
    if (isNaN(number)) return null;
    return number * 100000;
  }
  
  if (cleaned.includes('cr') || cleaned.includes('crore')) {
    const number = parseFloat(cleaned.replace(/cr|crore/g, '').trim());
    if (isNaN(number)) return null;
    return number * 10000000;
  }
  
  if (cleaned.includes('k')) {
    const number = parseFloat(cleaned.replace(/k/g, '').trim());
    if (isNaN(number)) return null;
    return number * 1000;
  }
  
  const number = parseFloat(cleaned);
  return isNaN(number) ? null : number;
}

/**
 * Format amount in Pakistani style
 */
function formatPKR(amount) {
  if (amount >= 10000000) {
    return (amount / 10000000).toFixed(1).replace('.0', '') + ' Cr';
  } else if (amount >= 100000) {
    return (amount / 100000).toFixed(1).replace('.0', '') + ' Lac';
  } else if (amount >= 1000) {
    return (amount / 1000).toFixed(0) + 'k';
  } else {
    return amount.toLocaleString('en-PK');
  }
}

/**
 * Generate unique loan ID
 */
function generateLoanId() {
  const prefix = 'LOAN';
  const timestamp = new Date().getTime().toString().slice(-6);
  return `${prefix}${timestamp}`;
}

/**
 * Generate unique payment ID
 */
function generatePaymentId() {
  const prefix = 'PAY';
  const timestamp = new Date().getTime().toString().slice(-6);
  return `${prefix}${timestamp}`;
}

/**
 * Calculate estimated completion based on expected monthly payment
 */
function calculateEstimatedCompletion(totalAmount, monthlyPayment) {
  if (!monthlyPayment || monthlyPayment <= 0) return '';
  
  const months = Math.ceil(totalAmount / monthlyPayment);
  const completionDate = new Date();
  completionDate.setMonth(completionDate.getMonth() + months);
  
  return `~${months} months (${completionDate.toLocaleDateString('en-PK')})`;
}

/**
 * Get or create main sheet
 */
function getOrCreateMainSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  
  if (!sheet) {
    initializeEnhancedLoanTracker();
    sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  }
  
  return sheet;
}

/**
 * Get or create payment history sheet
 */
function getOrCreatePaymentHistorySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(PAYMENTS_SHEET_NAME);
  
  if (!sheet) {
    initializeEnhancedLoanTracker();
    sheet = ss.getSheetByName(PAYMENTS_SHEET_NAME);
  }
  
  return sheet;
}

/**
 * Get or create dashboard sheet
 */
function getOrCreateDashboardSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  
  if (!sheet) {
    initializeEnhancedLoanTracker();
    sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  }
  
  return sheet;
}

/**
 * Get all active loans
 */
function getActiveLoans(sheet) {
  const loans = getAllLoans(sheet);
  return loans.filter(loan => loan.status === 'Active');
}

/**
 * Get all loans from main sheet
 */
function getAllLoans(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return [];
  
  const dataRange = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, 12);
  const values = dataRange.getValues();
  
  return values.map((row, index) => ({
    row: DATA_START_ROW + index,
    loanId: row[MAIN_COLS.LOAN_ID] || '',
    lender: row[MAIN_COLS.LENDER] || '',
    totalAmount: parseFloat(row[MAIN_COLS.TOTAL_AMOUNT]) || 0,
    totalPaid: parseFloat(row[MAIN_COLS.TOTAL_PAID]) || 0,
    remaining: parseFloat(row[MAIN_COLS.REMAINING]) || 0,
    paymentsCount: parseInt(row[MAIN_COLS.PAYMENTS_COUNT]) || 0,
    lastPaymentDate: row[MAIN_COLS.LAST_PAYMENT_DATE],
    lastPaymentAmount: parseFloat(row[MAIN_COLS.LAST_PAYMENT_AMOUNT]) || 0,
    estCompletion: row[MAIN_COLS.EST_COMPLETION] || '',
    purpose: row[MAIN_COLS.LOAN_PURPOSE] || '',
    status: row[MAIN_COLS.STATUS] || '',
    notes: row[MAIN_COLS.NOTES] || ''
  })).filter(loan => loan.lender && loan.totalAmount > 0);
}

/**
 * Get all payments from payment history sheet
 */
function getAllPayments(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return [];
  
  const dataRange = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, 8);
  const values = dataRange.getValues();
  
  return values.map(row => ({
    paymentId: row[PAYMENT_COLS.PAYMENT_ID] || '',
    loanId: row[PAYMENT_COLS.LOAN_ID] || '',
    lenderName: row[PAYMENT_COLS.LENDER] || '',
    paymentDate: row[PAYMENT_COLS.PAYMENT_DATE],
    paymentAmount: parseFloat(row[PAYMENT_COLS.PAYMENT_AMOUNT]) || 0,
    remainingAfter: parseFloat(row[PAYMENT_COLS.REMAINING_AFTER]) || 0,
    paymentMethod: row[PAYMENT_COLS.PAYMENT_METHOD] || '',
    notes: row[PAYMENT_COLS.NOTES] || ''
  })).filter(payment => payment.paymentId && payment.paymentAmount > 0);
}

/**
 * Find loan row by loan ID
 */
function findLoanRow(sheet, loanId) {
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return null;
  
  for (let row = DATA_START_ROW; row <= lastRow; row++) {
    const cellValue = sheet.getRange(row, MAIN_COLS.LOAN_ID + 1).getValue();
    if (cellValue === loanId) {
      return row;
    }
  }
  
  return null;
}

/**
 * Format main sheet row
 */
function formatMainSheetRow(sheet, rowIndex) {
  // Format currency columns
  const currencyColumns = [MAIN_COLS.TOTAL_AMOUNT + 1, MAIN_COLS.TOTAL_PAID + 1, MAIN_COLS.REMAINING + 1, MAIN_COLS.LAST_PAYMENT_AMOUNT + 1];
  
  currencyColumns.forEach(col => {
    const cell = sheet.getRange(rowIndex, col);
    const value = cell.getValue();
    if (value !== '' && !isNaN(value)) {
      cell.setNumberFormat('₨#,##0');
    }
  });
  
  // Format date column
  const dateCell = sheet.getRange(rowIndex, MAIN_COLS.LAST_PAYMENT_DATE + 1);
  if (dateCell.getValue() !== '') {
    dateCell.setNumberFormat('dd/mm/yyyy');
  }
  
  // Color code status
  const statusCell = sheet.getRange(rowIndex, MAIN_COLS.STATUS + 1);
  const status = statusCell.getValue();
  if (status === 'Paid') {
    statusCell.setBackground('#d4edda').setFontColor('#155724');
  } else if (status === 'Active') {
    statusCell.setBackground('#fff3cd').setFontColor('#856404');
  }
  
  // Alternate row colors
  if (rowIndex % 2 === 0) {
    sheet.getRange(rowIndex, 1, 1, 12).setBackground('#f8f9fa');
  }
}

/**
 * Format payment history row
 */
function formatPaymentHistoryRow(sheet, rowIndex) {
  // Format currency columns
  const currencyColumns = [PAYMENT_COLS.PAYMENT_AMOUNT + 1, PAYMENT_COLS.REMAINING_AFTER + 1];
  
  currencyColumns.forEach(col => {
    const cell = sheet.getRange(rowIndex, col);
    const value = cell.getValue();
    if (value !== '' && !isNaN(value)) {
      cell.setNumberFormat('₨#,##0');
    }
  });
  
  // Format date column
  const dateCell = sheet.getRange(rowIndex, PAYMENT_COLS.PAYMENT_DATE + 1);
  if (dateCell.getValue() !== '') {
    dateCell.setNumberFormat('dd/mm/yyyy');
  }
  
  // Alternate row colors
  if (rowIndex % 2 === 0) {
    sheet.getRange(rowIndex, 1, 1, 8).setBackground('#f8f9fa');
  }
}

/**
 * Refresh all calculations
 */
function refreshAllCalculations() {
  try {
    refreshDashboard();
    SpreadsheetApp.getUi().alert('✅ Success', 'All calculations and dashboard refreshed successfully!', SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (error) {
    SpreadsheetApp.getUi().alert('❌ Error', 'Error refreshing calculations: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Show enhanced help dialog
 */
function showHelpDialog() {
  const helpHtml = `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body { font-family: 'Google Sans', sans-serif; margin: 0; padding: 20px; background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%); min-height: 100vh; }
          .container { max-width: 700px; margin: 0 auto; background: white; border-radius: 16px; box-shadow: 0 10px 40px rgba(0,0,0,0.1); overflow: hidden; }
          .header { background: linear-gradient(135deg, #11998e, #38ef7d); color: white; padding: 24px; text-align: center; }
          .header h2 { margin: 0; font-size: 24px; font-weight: 500; }
          .content { padding: 32px; }
          .section { background: #f8f9fa; border-radius: 8px; padding: 20px; margin-bottom: 24px; }
          .section h3 { margin-top: 0; color: #333; font-size: 18px; }
          .feature-list { display: grid; gap: 12px; margin-top: 16px; }
          .feature { background: white; padding: 12px 16px; border-radius: 6px; border-left: 4px solid #11998e; }
          .tip { background: #e3f2fd; border-left: 4px solid #2196f3; padding: 16px; border-radius: 0 8px 8px 0; margin: 16px 0; }
          .close-btn { background: #11998e; color: white; border: none; padding: 12px 24px; border-radius: 8px; font-size: 14px; cursor: pointer; margin: 20px auto; display: block; }
        </style>
      </head>
      <body>
        <div class="container">
          <div class="header">
            <h2>❓ Enhanced Loan Tracker Guide</h2>
          </div>
          <div class="content">
            <div class="section">
              <h3>🚀 New Features</h3>
              <div class="feature-list">
                <div class="feature"><strong>Unlimited Payments:</strong> No more 2-payment limit. Record as many payments as needed.</div>
                <div class="feature"><strong>Live Dashboard:</strong> Real-time statistics and analytics in a dedicated sheet.</div>
                <div class="feature"><strong>Payment History:</strong> Complete log of all payments with detailed tracking.</div>
                <div class="feature"><strong>Smart Analytics:</strong> Purpose breakdown, progress tracking, and payment patterns.</div>
                <div class="feature"><strong>Enhanced Forms:</strong> Modern UI with better validation and user experience.</div>
              </div>
            </div>
            
            <div class="section">
              <h3>📊 Dashboard Features</h3>
              <div class="feature-list">
                <div class="feature"><strong>Key Statistics:</strong> Total loans, payments, amounts, and progress percentage.</div>
                <div class="feature"><strong>Purpose Breakdown:</strong> Visual analysis of loan purposes and amounts.</div>
                <div class="feature"><strong>Active Loans Summary:</strong> Current status of all ongoing loans.</div>
                <div class="feature"><strong>Recent Payments:</strong> Last 10 payments with details and trends.</div>
                <div class="feature"><strong>Auto-Refresh:</strong> Dashboard updates automatically when you add loans or payments.</div>
              </div>
            </div>
            
            <div class="section">
              <h3>💰 Amount Format Examples</h3>
              <div class="feature-list">
                <div class="feature"><strong>5 lac</strong> = ₨5,00,000</div>
                <div class="feature"><strong>2.5 lac</strong> = ₨2,50,000</div>
                <div class="feature"><strong>50k</strong> = ₨50,000</div>
                <div class="feature"><strong>1.2 cr</strong> = ₨1,20,00,000</div>
                <div class="feature"><strong>150000</strong> = ₨1,50,000</div>
              </div>
            </div>
            
            <div class="tip">
              <strong>💡 Quick Start:</strong><br>
              1. Initialize the tracker (creates 3 sheets)<br>
              2. Add your loans with the enhanced form<br>
              3. Record payments as you make them<br>
              4. Check the Dashboard sheet for live analytics<br>
              5. Use Payment History to view all transactions
            </div>
            
            <div class="section">
              <h3>🔧 Advanced Features</h3>
              <div class="feature-list">
                <div class="feature"><strong>Unique IDs:</strong> Each loan and payment gets a unique identifier for tracking.</div>
                <div class="feature"><strong>Payment Methods:</strong> Track how payments were made (cash, transfer, etc.).</div>
                <div class="feature"><strong>Status Tracking:</strong> Automatic status updates (Active → Paid).</div>
                <div class="feature"><strong>Color Coding:</strong> Visual indicators for loan status and progress.</div>
                <div class="feature"><strong>Export Ready:</strong> All data is properly formatted for reports.</div>
              </div>
            </div>
            
            <div class="tip">
              <strong>📝 Pro Tips:</strong><br>
              • Use the Dashboard sheet as your main overview<br>
              • Payment History sheet shows complete transaction log<br>
              • All calculations are automatic - no manual updates needed<br>
              • Works with Pakistani number formats (lac, crore, k)<br>
              • Forms prevent common data entry errors
            </div>
            
            <button class="close-btn" onclick="google.script.host.close()">Got It!</button>
          </div>
        </div>
      </body>
    </html>
  `;
  
  const htmlOutput = HtmlService.createHtmlOutput(helpHtml).setWidth(750).setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Enhanced Loan Tracker Guide');
}
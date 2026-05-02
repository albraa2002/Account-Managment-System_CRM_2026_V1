// ============================================================
//  AL ARRAB AGENCY — Account Management System
//  Setup.gs  |  One-click Sheet Builder
//  Run setupAll() once to create all sheets with headers
// ============================================================

function setupAll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupTeamSheet(ss);
  setupClientsSheet(ss);
  setupContractsSheet(ss);
  setupReportsSheet(ss);
  setupFollowUpsSheet(ss);
  setupActivityLogSheet(ss);
  installDailyTrigger();
  SpreadsheetApp.getUi().alert(
    '✅ Setup Complete!\n\n' +
    'All sheets created:\n' +
    '• Team\n• Clients\n• Contracts\n• Reports\n• FollowUps\n• ActivityLog\n\n' +
    'Daily automation trigger installed at 8 AM.\n\n' +
    'Next step: Deploy as Web App (Extensions → Apps Script → Deploy).'
  );
}

// ── HELPERS ────────────────────────────────────────────────

function ensureSheet(ss, name) {
  let sh = ss.getSheetByName(name);
  if (sh) {
    // Sheet exists — don't overwrite, just return
    Logger.log('Sheet already exists: ' + name);
    return null;
  }
  sh = ss.insertSheet(name);
  Logger.log('Created sheet: ' + name);
  return sh;
}

function styleHeader(sh, numCols) {
  const range = sh.getRange(1, 1, 1, numCols);
  range.setBackground('#1a1d2e')
       .setFontColor('#ffffff')
       .setFontWeight('bold')
       .setFontSize(11);
  sh.setFrozenRows(1);
  sh.setRowHeight(1, 32);
}

// ── TEAM ───────────────────────────────────────────────────

function setupTeamSheet(ss) {
  const sh = ensureSheet(ss, 'Team');
  if (!sh) return;

  const headers = [
    'ID', 'Name', 'Role', 'Email', 'Phone',
    'Capacity', 'Active', 'JoinDate'
  ];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader(sh, headers.length);

  // Roles validation
  const roleRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Head of Account','Team Leader','Senior Account Manager','Account Manager'], true)
    .build();
  sh.getRange('C2:C1000').setDataValidation(roleRule);

  // Active validation
  const boolRule = SpreadsheetApp.newDataValidation()
    .requireCheckbox().build();
  sh.getRange('G2:G1000').setDataValidation(boolRule);

  // Column widths
  sh.setColumnWidth(1, 120); // ID
  sh.setColumnWidth(2, 180); // Name
  sh.setColumnWidth(3, 180); // Role
  sh.setColumnWidth(4, 200); // Email
  sh.setColumnWidth(5, 130); // Phone

  // Sample data
  sh.appendRow(['TM0001','Ahmed Hassan','Head of Account','ahmed@alarrab.com','01001234567',0,true,'2022-01-01']);
  sh.appendRow(['TM0002','Sara Ali','Team Leader','sara@alarrab.com','01101234567',6,true,'2022-03-15']);
  sh.appendRow(['TM0003','Omar Khalil','Senior Account Manager','omar@alarrab.com','01201234567',10,true,'2022-06-01']);
  sh.appendRow(['TM0004','Nour Ibrahim','Account Manager','nour@alarrab.com','01501234567',10,true,'2023-01-10']);
  sh.appendRow(['TM0005','Karim Mostafa','Account Manager','karim@alarrab.com','01001234568',10,true,'2023-03-20']);
}

// ── CLIENTS ────────────────────────────────────────────────

function setupClientsSheet(ss) {
  const sh = ensureSheet(ss, 'Clients');
  if (!sh) return;

  const headers = [
    'ID','ClientName','CompanyName','Service','Package','Priority',
    'AccountManagerID','AccountManagerName','ClientRole',
    'StartDate','RenewalType','RenewalDate','Status','MonthlyFee',
    'Phone','LastContactDate','NextFollowUpDate','Notes','CreatedAt'
  ];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader(sh, headers.length);

  // Validations
  sh.getRange('D2:D1000').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['SEO','Social Media','Ads','Website','قوة العراب'], true).build()
  );
  sh.getRange('E2:E1000').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['Basic','Standard','Premium'], true).build()
  );
  sh.getRange('F2:F1000').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['Normal','VIP'], true).build()
  );
  sh.getRange('K2:K1000').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['Monthly','1.5 Month','2 Months','3 Months','6 Months','Yearly'], true).build()
  );
  sh.getRange('M2:M1000').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['Active','Inactive'], true).build()
  );

  // Date columns formatting
  sh.getRange('J2:J1000').setNumberFormat('yyyy-mm-dd');
  sh.getRange('L2:L1000').setNumberFormat('yyyy-mm-dd');
  sh.getRange('P2:P1000').setNumberFormat('yyyy-mm-dd');
  sh.getRange('Q2:Q1000').setNumberFormat('yyyy-mm-dd');

  // Column widths
  sh.setColumnWidth(1, 110);
  sh.setColumnWidth(2, 160);
  sh.setColumnWidth(3, 160);
  sh.setColumnWidth(8, 160);
  sh.setColumnWidth(18, 200);

  // Conditional formatting — VIP rows
  const vipRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$F2="VIP"')
    .setBackground('#faf5ff')
    .setRanges([sh.getRange('A2:S1000')])
    .build();
  sh.setConditionalFormatRules([vipRule]);
}

// ── CONTRACTS ──────────────────────────────────────────────

function setupContractsSheet(ss) {
  const sh = ensureSheet(ss, 'Contracts');
  if (!sh) return;

  const headers = [
    'ID','ClientID','ClientName','Service','AccountManagerID','AccountManagerName',
    'ContractStatus','PaymentType','RenewalDate','LastPaymentDate','NextPaymentDue',
    'PaymentStatus','TotalContractValue','PaidAmount','Remaining',
    'LostReason','Notes','Priority','LastRenewalDate','NextRenewalDate','CreatedAt'
  ];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader(sh, headers.length);

  sh.getRange('G2:G1000').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['Active','On Hold','Cancelled','Lost','Renewed'], true).build()
  );
  sh.getRange('H2:H1000').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['Full Payment','Installments','Monthly Payment','Quarterly','Yearly'], true).build()
  );
  sh.getRange('L2:L1000').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['Paid','Pending','Overdue','Partially Paid'], true).build()
  );

  sh.getRange('I2:I1000').setNumberFormat('yyyy-mm-dd');
  sh.getRange('J2:J1000').setNumberFormat('yyyy-mm-dd');
  sh.getRange('K2:K1000').setNumberFormat('yyyy-mm-dd');
  sh.getRange('M2:O1000').setNumberFormat('#,##0');

  // Conditional formatting
  const overdueRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$L2="Overdue"')
    .setBackground('#fff5f5').setFontColor('#dc2626')
    .setRanges([sh.getRange('A2:U1000')]).build();
  const lostRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$G2="Lost"')
    .setBackground('#fef2f2')
    .setRanges([sh.getRange('A2:U1000')]).build();
  sh.setConditionalFormatRules([overdueRule, lostRule]);

  sh.setColumnWidth(3, 160);
  sh.setColumnWidth(6, 160);
  sh.setColumnWidth(13, 140);
  sh.setColumnWidth(14, 120);
  sh.setColumnWidth(15, 120);
}

// ── REPORTS ────────────────────────────────────────────────

function setupReportsSheet(ss) {
  const sh = ensureSheet(ss, 'Reports');
  if (!sh) return;

  const headers = [
    'ID','ClientID','ClientName','Service','ReportType',
    'DueDate','SubmittedDate','Status','AccountManagerID','AccountManagerName',
    'Notes','Period','CreatedAt'
  ];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader(sh, headers.length);

  sh.getRange('E2:E1000').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['Weekly','Monthly'], true).build()
  );
  sh.getRange('H2:H1000').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['Pending','Submitted','Delayed'], true).build()
  );

  sh.getRange('F2:G1000').setNumberFormat('yyyy-mm-dd');

  const delayedRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$H2="Delayed"')
    .setBackground('#fff5f5').setFontColor('#dc2626')
    .setRanges([sh.getRange('A2:M1000')]).build();
  const submittedRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$H2="Submitted"')
    .setBackground('#f0fdf4')
    .setRanges([sh.getRange('A2:M1000')]).build();
  sh.setConditionalFormatRules([delayedRule, submittedRule]);

  sh.setColumnWidth(3, 160);
  sh.setColumnWidth(10, 160);
}

// ── FOLLOW-UPS ─────────────────────────────────────────────

function setupFollowUpsSheet(ss) {
  const sh = ensureSheet(ss, 'FollowUps');
  if (!sh) return;

  const headers = [
    'ID','ClientID','ClientName','Service','AccountManagerID','AccountManagerName',
    'LastContactDate','NextFollowUpDate','Method','Status','Notes','CreatedAt'
  ];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader(sh, headers.length);

  sh.getRange('I2:I1000').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['Phone','WhatsApp','Email','Meeting'], true).build()
  );
  sh.getRange('J2:J1000').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['Upcoming','Today','Missed','Done'], true).build()
  );

  sh.getRange('G2:H1000').setNumberFormat('yyyy-mm-dd');

  const missedRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$J2="Missed"')
    .setBackground('#fff5f5').setFontColor('#dc2626')
    .setRanges([sh.getRange('A2:L1000')]).build();
  const todayRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$J2="Today"')
    .setBackground('#fffbeb')
    .setRanges([sh.getRange('A2:L1000')]).build();
  sh.setConditionalFormatRules([missedRule, todayRule]);

  sh.setColumnWidth(3, 160);
  sh.setColumnWidth(6, 160);
}

// ── ACTIVITY LOG ───────────────────────────────────────────

function setupActivityLogSheet(ss) {
  const sh = ensureSheet(ss, 'ActivityLog');
  if (!sh) return;

  const headers = ['Timestamp','User','Action','Entity','EntityID','Details'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader(sh, headers.length);
  sh.getRange('A2:A1000').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  sh.setColumnWidth(1, 160);
  sh.setColumnWidth(2, 180);
  sh.setColumnWidth(6, 260);
}

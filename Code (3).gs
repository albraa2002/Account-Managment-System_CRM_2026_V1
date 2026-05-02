// ============================================================
//  AL ARRAB AGENCY — Account Management System
//  Code.gs  |  Backend API + Automation
//  Version: 1.0.0
// ============================================================

// ── SHEET NAMES (single source of truth) ──────────────────
const SHEET = {
  CLIENTS:    'Clients',
  CONTRACTS:  'Contracts',
  REPORTS:    'Reports',
  FOLLOWUPS:  'FollowUps',
  TEAM:       'Team',
  LOG:        'ActivityLog'
};

// ── CORS HEADERS ───────────────────────────────────────────
function corsHeaders() {
  return {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Content-Type': 'application/json'
  };
}

// ── MAIN ROUTER ────────────────────────────────────────────
function doGet(e) {
  try {
    const action = e.parameter.action || '';
    const params = e.parameter;
    let result;

    switch (action) {
      case 'getClients':      result = getClients(params);      break;
      case 'getClient':       result = getClient(params.id);    break;
      case 'getTeam':         result = getTeam();               break;
      case 'getReports':      result = getReports(params);      break;
      case 'getFollowUps':    result = getFollowUps(params);    break;
      case 'getContracts':    result = getContracts(params);    break;
      case 'getDashboard':    result = getDashboardStats();     break;
      case 'getAlerts':       result = getAlerts();             break;
      default:
        result = { error: 'Unknown action: ' + action };
    }

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, data: result }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    const body   = JSON.parse(e.postData.contents);
    const action = body.action || '';
    let result;

    switch (action) {
      case 'addClient':       result = addClient(body.data);              break;
      case 'updateClient':    result = updateClient(body.id, body.data);  break;
      case 'deleteClient':    result = deleteClient(body.id);             break;

      case 'addContract':     result = addContract(body.data);            break;
      case 'updateContract':  result = updateContract(body.id, body.data);break;

      case 'addReport':       result = addReport(body.data);              break;
      case 'updateReport':    result = updateReport(body.id, body.data);  break;

      case 'addFollowUp':     result = addFollowUp(body.data);            break;
      case 'updateFollowUp':  result = updateFollowUp(body.id, body.data);break;

      case 'addTeamMember':   result = addTeamMember(body.data);          break;
      case 'updateTeamMember':result = updateTeamMember(body.id,body.data);break;

      default:
        result = { error: 'Unknown action: ' + action };
    }

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, data: result }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================================
//  HELPERS
// ============================================================

function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error('Sheet not found: ' + name);
  return sh;
}

function sheetToObjects(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map((row, i) => {
    const obj = { _rowIndex: i + 2 }; // 1-based, header is row 1
    headers.forEach((h, j) => { obj[h] = row[j]; });
    return obj;
  });
}

function genId(prefix) {
  return prefix + Date.now().toString(36).toUpperCase();
}

function today() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function fmtDate(d) {
  if (!d) return '';
  if (d instanceof Date) return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(d).slice(0, 10);
}

function logActivity(action, entity, entityId, details) {
  try {
    const sh = getSheet(SHEET.LOG);
    sh.appendRow([
      new Date(),
      Session.getActiveUser().getEmail() || 'system',
      action,
      entity,
      entityId,
      details || ''
    ]);
  } catch (_) {}
}

// ============================================================
//  TEAM
// ============================================================

function getTeam() {
  return sheetToObjects(getSheet(SHEET.TEAM)).map(r => ({
    id:       r['ID'],
    name:     r['Name'],
    role:     r['Role'],
    email:    r['Email'],
    phone:    r['Phone'],
    capacity: Number(r['Capacity']) || 10,
    active:   r['Active'] === true || r['Active'] === 'TRUE'
  }));
}

function addTeamMember(data) {
  const sh = getSheet(SHEET.TEAM);
  const id = genId('TM');
  sh.appendRow([
    id, data.name, data.role, data.email || '',
    data.phone || '', data.capacity || 10, true, today()
  ]);
  logActivity('ADD', 'Team', id, data.name);
  return { id };
}

function updateTeamMember(id, data) {
  const sh   = getSheet(SHEET.TEAM);
  const rows = sheetToObjects(sh);
  const row  = rows.find(r => r['ID'] === id);
  if (!row) throw new Error('Team member not found: ' + id);
  const ri = row._rowIndex;
  sh.getRange(ri, 2).setValue(data.name     ?? row['Name']);
  sh.getRange(ri, 3).setValue(data.role     ?? row['Role']);
  sh.getRange(ri, 4).setValue(data.email    ?? row['Email']);
  sh.getRange(ri, 5).setValue(data.phone    ?? row['Phone']);
  sh.getRange(ri, 6).setValue(data.capacity ?? row['Capacity']);
  sh.getRange(ri, 7).setValue(data.active   ?? row['Active']);
  logActivity('UPDATE', 'Team', id, data.name);
  return { ok: true };
}

// ============================================================
//  CLIENTS
// ============================================================

function getClients(params) {
  let rows = sheetToObjects(getSheet(SHEET.CLIENTS));
  if (params.manager)  rows = rows.filter(r => r['AccountManagerID'] === params.manager);
  if (params.service)  rows = rows.filter(r => r['Service']          === params.service);
  if (params.priority) rows = rows.filter(r => r['Priority']         === params.priority);
  if (params.status)   rows = rows.filter(r => r['Status']           === params.status);
  if (params.search) {
    const q = params.search.toLowerCase();
    rows = rows.filter(r =>
      String(r['ClientName']).toLowerCase().includes(q) ||
      String(r['CompanyName']).toLowerCase().includes(q) ||
      String(r['ID']).toLowerCase().includes(q)
    );
  }
  return rows.map(r => ({
    id:               r['ID'],
    clientName:       r['ClientName'],
    companyName:      r['CompanyName'],
    service:          r['Service'],
    package:          r['Package'],
    priority:         r['Priority'],
    accountManagerId: r['AccountManagerID'],
    accountManager:   r['AccountManagerName'],
    role:             r['ClientRole'],
    startDate:        fmtDate(r['StartDate']),
    renewalType:      r['RenewalType'],
    renewalDate:      fmtDate(r['RenewalDate']),
    status:           r['Status'],
    monthlyFee:       Number(r['MonthlyFee']) || 0,
    lastContact:      fmtDate(r['LastContactDate']),
    nextFollowup:     fmtDate(r['NextFollowUpDate']),
    notes:            r['Notes'] || '',
    phone:            r['Phone'] || '',
    createdAt:        fmtDate(r['CreatedAt'])
  }));
}

function getClient(id) {
  const rows = sheetToObjects(getSheet(SHEET.CLIENTS));
  const r = rows.find(r => r['ID'] === id);
  if (!r) throw new Error('Client not found: ' + id);
  return r;
}

function addClient(data) {
  const sh = getSheet(SHEET.CLIENTS);
  const id = genId('CL');
  const renewalDate = calcRenewalDate(data.startDate, data.renewalType);
  sh.appendRow([
    id,
    data.clientName,
    data.companyName    || '',
    data.service,
    data.package        || 'Standard',
    data.priority       || 'Normal',
    data.accountManagerId,
    data.accountManager || '',
    data.clientRole     || '',
    data.startDate,
    data.renewalType,
    renewalDate,
    'Active',
    Number(data.monthlyFee) || 0,
    data.phone          || '',
    today(),            // lastContact
    data.nextFollowup   || '',
    data.notes          || '',
    today()             // createdAt
  ]);
  logActivity('ADD', 'Client', id, data.clientName);
  return { id, renewalDate };
}

function updateClient(id, data) {
  const sh   = getSheet(SHEET.CLIENTS);
  const rows = sheetToObjects(sh);
  const row  = rows.find(r => r['ID'] === id);
  if (!row) throw new Error('Client not found: ' + id);
  const ri = row._rowIndex;
  const cols = {
    'ClientName':       data.clientName,
    'CompanyName':      data.companyName,
    'Service':          data.service,
    'Package':          data.package,
    'Priority':         data.priority,
    'AccountManagerID': data.accountManagerId,
    'AccountManagerName': data.accountManager,
    'ClientRole':       data.clientRole,
    'StartDate':        data.startDate,
    'RenewalType':      data.renewalType,
    'RenewalDate':      data.renewalDate,
    'Status':           data.status,
    'MonthlyFee':       data.monthlyFee,
    'Phone':            data.phone,
    'LastContactDate':  data.lastContact,
    'NextFollowUpDate': data.nextFollowup,
    'Notes':            data.notes
  };
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  headers.forEach((h, j) => {
    if (cols[h] !== undefined) sh.getRange(ri, j + 1).setValue(cols[h]);
  });
  logActivity('UPDATE', 'Client', id, data.clientName);
  return { ok: true };
}

function deleteClient(id) {
  const sh   = getSheet(SHEET.CLIENTS);
  const rows = sheetToObjects(sh);
  const row  = rows.find(r => r['ID'] === id);
  if (!row) throw new Error('Client not found: ' + id);
  sh.deleteRow(row._rowIndex);
  logActivity('DELETE', 'Client', id, '');
  return { ok: true };
}

function calcRenewalDate(startDate, renewalType) {
  const d = new Date(startDate);
  const map = {
    'Monthly':   30,
    '1.5 Month': 45,
    '2 Months':  60,
    '3 Months':  90,
    '6 Months': 180,
    'Yearly':   365
  };
  d.setDate(d.getDate() + (map[renewalType] || 30));
  return fmtDate(d);
}

// ============================================================
//  CONTRACTS
// ============================================================

function getContracts(params) {
  let rows = sheetToObjects(getSheet(SHEET.CONTRACTS));
  if (params.status)    rows = rows.filter(r => r['ContractStatus'] === params.status);
  if (params.payment)   rows = rows.filter(r => r['PaymentStatus']  === params.payment);
  if (params.manager)   rows = rows.filter(r => r['AccountManagerID'] === params.manager);
  if (params.search) {
    const q = params.search.toLowerCase();
    rows = rows.filter(r =>
      String(r['ClientName']).toLowerCase().includes(q) ||
      String(r['ClientID']).toLowerCase().includes(q)
    );
  }
  return rows.map(r => {
    const total   = Number(r['TotalContractValue']) || 0;
    const paid    = Number(r['PaidAmount'])          || 0;
    const remaining = total - paid;
    return {
      id:               r['ID'],
      clientId:         r['ClientID'],
      clientName:       r['ClientName'],
      service:          r['Service'],
      accountManagerId: r['AccountManagerID'],
      accountManager:   r['AccountManagerName'],
      contractStatus:   r['ContractStatus'],
      paymentType:      r['PaymentType'],
      renewalDate:      fmtDate(r['RenewalDate']),
      lastPayDate:      fmtDate(r['LastPaymentDate']),
      nextPayDate:      fmtDate(r['NextPaymentDue']),
      paymentStatus:    r['PaymentStatus'],
      totalValue:       total,
      paidAmount:       paid,
      remaining:        remaining,
      lostReason:       r['LostReason']    || '',
      notes:            r['Notes']         || '',
      priority:         r['Priority']      || 'Normal',
      lastRenewalDate:  fmtDate(r['LastRenewalDate']),
      nextRenewalDate:  fmtDate(r['NextRenewalDate'])
    };
  });
}

function addContract(data) {
  const sh = getSheet(SHEET.CONTRACTS);
  const id = genId('CN');
  const total   = Number(data.totalValue)  || 0;
  const paid    = Number(data.paidAmount)  || 0;
  const payStatus = paid >= total ? 'Paid' : paid > 0 ? 'Partially Paid' : 'Pending';
  sh.appendRow([
    id,
    data.clientId,
    data.clientName,
    data.service,
    data.accountManagerId  || '',
    data.accountManager    || '',
    data.contractStatus    || 'Active',
    data.paymentType       || 'Full Payment',
    data.renewalDate       || '',
    data.lastPayDate       || '',
    data.nextPayDate       || '',
    payStatus,
    total, paid,
    total - paid,
    data.lostReason        || '',
    data.notes             || '',
    data.priority          || 'Normal',
    data.lastRenewalDate   || '',
    data.nextRenewalDate   || '',
    today()
  ]);
  logActivity('ADD', 'Contract', id, data.clientName);
  return { id };
}

function updateContract(id, data) {
  const sh   = getSheet(SHEET.CONTRACTS);
  const rows = sheetToObjects(sh);
  const row  = rows.find(r => r['ID'] === id);
  if (!row) throw new Error('Contract not found: ' + id);
  const ri      = row._rowIndex;
  const total   = Number(data.totalValue  ?? row['TotalContractValue']) || 0;
  const paid    = Number(data.paidAmount  ?? row['PaidAmount'])          || 0;
  const payStatus = data.paymentStatus ?? (paid >= total ? 'Paid' : paid > 0 ? 'Partially Paid' : 'Pending');

  const cols = {
    'ContractStatus':     data.contractStatus,
    'PaymentType':        data.paymentType,
    'RenewalDate':        data.renewalDate,
    'LastPaymentDate':    data.lastPayDate,
    'NextPaymentDue':     data.nextPayDate,
    'PaymentStatus':      payStatus,
    'TotalContractValue': total,
    'PaidAmount':         paid,
    'Remaining':          total - paid,
    'LostReason':         data.lostReason,
    'Notes':              data.notes,
    'LastRenewalDate':    data.lastRenewalDate,
    'NextRenewalDate':    data.nextRenewalDate
  };
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  headers.forEach((h, j) => {
    if (cols[h] !== undefined) sh.getRange(ri, j + 1).setValue(cols[h]);
  });
  logActivity('UPDATE', 'Contract', id, data.clientName || '');
  return { ok: true };
}

// ============================================================
//  REPORTS
// ============================================================

function getReports(params) {
  let rows = sheetToObjects(getSheet(SHEET.REPORTS));
  if (params.status)  rows = rows.filter(r => r['Status']            === params.status);
  if (params.type)    rows = rows.filter(r => r['ReportType']        === params.type);
  if (params.manager) rows = rows.filter(r => r['AccountManagerID']  === params.manager);
  if (params.client)  rows = rows.filter(r => r['ClientID']          === params.client);
  return rows.map(r => ({
    id:           r['ID'],
    clientId:     r['ClientID'],
    clientName:   r['ClientName'],
    service:      r['Service'],
    reportType:   r['ReportType'],
    dueDate:      fmtDate(r['DueDate']),
    submittedDate:fmtDate(r['SubmittedDate']),
    status:       r['Status'],
    accountManagerId: r['AccountManagerID'],
    accountManager:   r['AccountManagerName'],
    notes:        r['Notes'] || '',
    period:       r['Period'] || ''
  }));
}

function addReport(data) {
  const sh = getSheet(SHEET.REPORTS);
  const id = genId('RP');
  sh.appendRow([
    id,
    data.clientId,
    data.clientName,
    data.service,
    data.reportType,    // Weekly / Monthly
    data.dueDate,
    '',                 // submittedDate
    'Pending',
    data.accountManagerId || '',
    data.accountManager   || '',
    data.notes            || '',
    data.period           || '',
    today()
  ]);
  logActivity('ADD', 'Report', id, data.clientName + ' — ' + data.reportType);
  return { id };
}

function updateReport(id, data) {
  const sh   = getSheet(SHEET.REPORTS);
  const rows = sheetToObjects(sh);
  const row  = rows.find(r => r['ID'] === id);
  if (!row) throw new Error('Report not found: ' + id);
  const ri = row._rowIndex;
  const cols = {
    'DueDate':       data.dueDate,
    'SubmittedDate': data.status === 'Submitted' ? today() : (data.submittedDate ?? row['SubmittedDate']),
    'Status':        data.status,
    'Notes':         data.notes
  };
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  headers.forEach((h, j) => {
    if (cols[h] !== undefined) sh.getRange(ri, j + 1).setValue(cols[h]);
  });
  logActivity('UPDATE', 'Report', id, data.status || '');
  return { ok: true };
}

// ============================================================
//  FOLLOW-UPS
// ============================================================

function getFollowUps(params) {
  let rows = sheetToObjects(getSheet(SHEET.FOLLOWUPS));
  if (params.status)  rows = rows.filter(r => r['Status']           === params.status);
  if (params.manager) rows = rows.filter(r => r['AccountManagerID'] === params.manager);
  if (params.client)  rows = rows.filter(r => r['ClientID']         === params.client);
  const todayStr = today();
  return rows.map(r => {
    const nf    = fmtDate(r['NextFollowUpDate']);
    let status  = r['Status'];
    if (!status || status === 'Pending') {
      if (nf === todayStr)     status = 'Today';
      else if (nf < todayStr)  status = 'Missed';
      else                     status = 'Upcoming';
    }
    return {
      id:           r['ID'],
      clientId:     r['ClientID'],
      clientName:   r['ClientName'],
      service:      r['Service'],
      accountManagerId: r['AccountManagerID'],
      accountManager:   r['AccountManagerName'],
      lastContact:  fmtDate(r['LastContactDate']),
      nextFollowup: nf,
      method:       r['Method'],
      status,
      notes:        r['Notes'] || ''
    };
  });
}

function addFollowUp(data) {
  const sh = getSheet(SHEET.FOLLOWUPS);
  const id = genId('FU');
  sh.appendRow([
    id,
    data.clientId,
    data.clientName,
    data.service,
    data.accountManagerId || '',
    data.accountManager   || '',
    data.lastContact      || today(),
    data.nextFollowup,
    data.method           || 'Phone',
    'Upcoming',
    data.notes            || '',
    today()
  ]);
  // Update client's NextFollowUpDate
  try { updateClientFollowup(data.clientId, data.nextFollowup); } catch(_){}
  logActivity('ADD', 'FollowUp', id, data.clientName);
  return { id };
}

function updateFollowUp(id, data) {
  const sh   = getSheet(SHEET.FOLLOWUPS);
  const rows = sheetToObjects(sh);
  const row  = rows.find(r => r['ID'] === id);
  if (!row) throw new Error('Follow-up not found: ' + id);
  const ri = row._rowIndex;
  const cols = {
    'LastContactDate':  data.lastContact,
    'NextFollowUpDate': data.nextFollowup,
    'Method':           data.method,
    'Status':           data.status,
    'Notes':            data.notes
  };
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  headers.forEach((h, j) => {
    if (cols[h] !== undefined) sh.getRange(ri, j + 1).setValue(cols[h]);
  });
  if (data.nextFollowup) {
    try { updateClientFollowup(row['ClientID'], data.nextFollowup); } catch(_){}
  }
  logActivity('UPDATE', 'FollowUp', id, data.status || '');
  return { ok: true };
}

function updateClientFollowup(clientId, nextDate) {
  const sh   = getSheet(SHEET.CLIENTS);
  const rows = sheetToObjects(sh);
  const row  = rows.find(r => r['ID'] === clientId);
  if (!row) return;
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const col = headers.indexOf('NextFollowUpDate') + 1;
  if (col > 0) sh.getRange(row._rowIndex, col).setValue(nextDate);
}

// ============================================================
//  DASHBOARD STATS
// ============================================================

function getDashboardStats() {
  const clients   = sheetToObjects(getSheet(SHEET.CLIENTS));
  const contracts = sheetToObjects(getSheet(SHEET.CONTRACTS));
  const reports   = sheetToObjects(getSheet(SHEET.REPORTS));
  const followups = sheetToObjects(getSheet(SHEET.FOLLOWUPS));
  const team      = sheetToObjects(getSheet(SHEET.TEAM));
  const todayStr  = today();

  // Clients
  const activeClients    = clients.filter(c => c['Status'] === 'Active');
  const vipClients       = clients.filter(c => c['Priority'] === 'VIP');
  const nearRenewal      = clients.filter(c => {
    const d = fmtDate(c['RenewalDate']);
    const diff = (new Date(d) - new Date()) / 86400000;
    return diff >= 0 && diff <= 30;
  });

  // Reports
  const overdueReports   = reports.filter(r => r['Status'] === 'Delayed');
  const pendingReports   = reports.filter(r => r['Status'] === 'Pending');
  const submittedReports = reports.filter(r => r['Status'] === 'Submitted');

  // Follow-ups
  const todayFU          = followups.filter(f => fmtDate(f['NextFollowUpDate']) === todayStr);
  const missedFU         = followups.filter(f => {
    const nf = fmtDate(f['NextFollowUpDate']);
    return nf < todayStr && f['Status'] !== 'Done';
  });

  // Contracts / Revenue
  const activeContracts  = contracts.filter(c => c['ContractStatus'] === 'Active');
  const lostContracts    = contracts.filter(c => c['ContractStatus'] === 'Lost');
  const renewedContracts = contracts.filter(c => c['ContractStatus'] === 'Renewed');
  const overduePayments  = contracts.filter(c => c['PaymentStatus']  === 'Overdue');
  const totalRevenue     = contracts.reduce((s, c) => s + (Number(c['PaidAmount']) || 0), 0);
  const monthlyRevenue   = activeContracts.reduce((s, c) => s + (Number(c['TotalContractValue']) || 0), 0) / 12;
  const totalRemaining   = contracts.filter(c => c['PaymentStatus'] !== 'Paid')
                                    .reduce((s, c) => s + (Number(c['Remaining']) || 0), 0);

  // Per-manager breakdown
  const managerStats = team.filter(t => t['Active'] === true || t['Active'] === 'TRUE').map(m => {
    const mId      = m['ID'];
    const mClients = clients.filter(c => c['AccountManagerID'] === mId);
    const mReports = reports.filter(r => r['AccountManagerID'] === mId);
    return {
      id:       mId,
      name:     m['Name'],
      role:     m['Role'],
      clients:  mClients.length,
      vip:      mClients.filter(c => c['Priority'] === 'VIP').length,
      capacity: Number(m['Capacity']) || 10,
      submitted:mReports.filter(r => r['Status'] === 'Submitted').length,
      delayed:  mReports.filter(r => r['Status'] === 'Delayed').length,
      pending:  mReports.filter(r => r['Status'] === 'Pending').length
    };
  });

  return {
    totalClients:       clients.length,
    activeClients:      activeClients.length,
    vipClients:         vipClients.length,
    nearRenewal:        nearRenewal.length,
    overdueReports:     overdueReports.length,
    pendingReports:     pendingReports.length,
    submittedReports:   submittedReports.length,
    todayFollowUps:     todayFU.length,
    missedFollowUps:    missedFU.length,
    totalTeam:          team.filter(t => t['Active'] === true || t['Active'] === 'TRUE').length,
    activeContracts:    activeContracts.length,
    lostClients:        lostContracts.length,
    renewedClients:     renewedContracts.length,
    overduePayments:    overduePayments.length,
    totalRevenue:       Math.round(totalRevenue),
    monthlyRevenue:     Math.round(monthlyRevenue),
    totalRemaining:     Math.round(totalRemaining),
    managerStats,
    serviceBreakdown: {
      SEO:         clients.filter(c => c['Service'] === 'SEO').length,
      'Social Media': clients.filter(c => c['Service'] === 'Social Media').length,
      Ads:         clients.filter(c => c['Service'] === 'Ads').length,
      Website:     clients.filter(c => c['Service'] === 'Website').length,
      'قوة العراب': clients.filter(c => c['Service'] === 'قوة العراب').length
    }
  };
}

// ============================================================
//  ALERTS
// ============================================================

function getAlerts() {
  const alerts    = [];
  const todayStr  = today();
  const clients   = sheetToObjects(getSheet(SHEET.CLIENTS));
  const contracts = sheetToObjects(getSheet(SHEET.CONTRACTS));
  const reports   = sheetToObjects(getSheet(SHEET.REPORTS));
  const followups = sheetToObjects(getSheet(SHEET.FOLLOWUPS));

  // Overdue payments
  contracts.filter(c => c['PaymentStatus'] === 'Overdue').forEach(c => {
    alerts.push({ type: 'danger', category: 'Payment', message: `Overdue payment: ${c['ClientName']} — EGP ${Number(c['Remaining']).toLocaleString()}`, clientId: c['ClientID'] });
  });

  // Renewals in 7 days
  clients.filter(c => {
    const diff = (new Date(fmtDate(c['RenewalDate'])) - new Date()) / 86400000;
    return diff >= 0 && diff <= 7;
  }).forEach(c => {
    const diff = Math.round((new Date(fmtDate(c['RenewalDate'])) - new Date()) / 86400000);
    alerts.push({ type: 'warn', category: 'Renewal', message: `Renewal due in ${diff}d: ${c['ClientName']} (${c['Service']})`, clientId: c['ID'] });
  });

  // Overdue reports
  reports.filter(r => r['Status'] === 'Delayed').forEach(r => {
    alerts.push({ type: 'danger', category: 'Report', message: `Delayed ${r['ReportType']} report: ${r['ClientName']}`, clientId: r['ClientID'] });
  });

  // Missed follow-ups
  followups.filter(f => fmtDate(f['NextFollowUpDate']) < todayStr && f['Status'] !== 'Done').forEach(f => {
    alerts.push({ type: 'warn', category: 'FollowUp', message: `Missed follow-up: ${f['ClientName']} (${f['Method']})`, clientId: f['ClientID'] });
  });

  // No follow-up scheduled
  clients.filter(c => !c['NextFollowUpDate']).forEach(c => {
    alerts.push({ type: 'info', category: 'FollowUp', message: `No follow-up scheduled: ${c['ClientName']}`, clientId: c['ID'] });
  });

  return alerts;
}

// ============================================================
//  DAILY AUTOMATION (set as time-based trigger)
// ============================================================

function dailyAutomation() {
  autoUpdateReportStatuses();
  autoUpdateFollowUpStatuses();
  autoUpdateContractPaymentStatuses();
  autoUpdateRenewalStatuses();
  sendDailySummaryEmail();
}

function autoUpdateReportStatuses() {
  const sh      = getSheet(SHEET.REPORTS);
  const rows    = sheetToObjects(sh);
  const todayStr= today();
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const statusCol = headers.indexOf('Status') + 1;
  rows.forEach(r => {
    if (r['Status'] === 'Submitted') return;
    const due = fmtDate(r['DueDate']);
    if (due && due < todayStr) {
      sh.getRange(r._rowIndex, statusCol).setValue('Delayed');
    }
  });
}

function autoUpdateFollowUpStatuses() {
  const sh      = getSheet(SHEET.FOLLOWUPS);
  const rows    = sheetToObjects(sh);
  const todayStr= today();
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const statusCol = headers.indexOf('Status') + 1;
  rows.forEach(r => {
    if (r['Status'] === 'Done') return;
    const nf = fmtDate(r['NextFollowUpDate']);
    if (!nf) return;
    if (nf === todayStr)     sh.getRange(r._rowIndex, statusCol).setValue('Today');
    else if (nf < todayStr)  sh.getRange(r._rowIndex, statusCol).setValue('Missed');
    else                     sh.getRange(r._rowIndex, statusCol).setValue('Upcoming');
  });
}

function autoUpdateContractPaymentStatuses() {
  const sh      = getSheet(SHEET.CONTRACTS);
  const rows    = sheetToObjects(sh);
  const todayStr= today();
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const payStatusCol = headers.indexOf('PaymentStatus') + 1;
  const remainingCol = headers.indexOf('Remaining') + 1;
  rows.forEach(r => {
    const total = Number(r['TotalContractValue']) || 0;
    const paid  = Number(r['PaidAmount'])          || 0;
    const rem   = total - paid;
    sh.getRange(r._rowIndex, remainingCol).setValue(rem);
    if (r['ContractStatus'] === 'Lost' || r['ContractStatus'] === 'Cancelled') return;
    const nextPay = fmtDate(r['NextPaymentDue']);
    let newStatus;
    if (rem <= 0)                              newStatus = 'Paid';
    else if (nextPay && nextPay < todayStr)    newStatus = 'Overdue';
    else if (paid > 0 && paid < total)         newStatus = 'Partially Paid';
    else                                       newStatus = 'Pending';
    sh.getRange(r._rowIndex, payStatusCol).setValue(newStatus);
  });
}

function autoUpdateRenewalStatuses() {
  const sh      = getSheet(SHEET.CLIENTS);
  const rows    = sheetToObjects(sh);
  const todayStr= today();
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const statusCol = headers.indexOf('Status') + 1;
  rows.forEach(r => {
    if (r['Status'] === 'Inactive') return;
    const rd = fmtDate(r['RenewalDate']);
    if (rd && rd < todayStr) {
      // Only flag — don't auto-set to Inactive, let team handle
      // Could add a "RenewalStatus" column here if needed
    }
  });
}

function sendDailySummaryEmail() {
  try {
    const stats    = getDashboardStats();
    const alerts   = getAlerts();
    const teamSheet= sheetToObjects(getSheet(SHEET.TEAM));
    const heads    = teamSheet.filter(t => t['Role'] === 'Head of Account' && t['Email']);
    if (!heads.length) return;

    const body = `
Al Arrab Account Management — Daily Summary (${today()})

📊 Overview
• Total Clients: ${stats.totalClients} (${stats.activeClients} active)
• VIP Clients: ${stats.vipClients}
• Near Renewal (30d): ${stats.nearRenewal}
• Today's Follow-ups: ${stats.todayFollowUps}

📄 Reports
• Overdue: ${stats.overdueReports}
• Pending: ${stats.pendingReports}
• Submitted: ${stats.submittedReports}

💰 Contracts
• Active: ${stats.activeContracts}
• Overdue Payments: ${stats.overduePayments}
• Revenue Collected: EGP ${stats.totalRevenue.toLocaleString()}

⚠ Alerts (${alerts.length} total)
${alerts.slice(0, 10).map(a => `[${a.type.toUpperCase()}] ${a.message}`).join('\n')}
${alerts.length > 10 ? `... and ${alerts.length - 10} more` : ''}

Powered by Al Arrab AMMS
    `.trim();

    heads.forEach(h => {
      GmailApp.sendEmail(h['Email'], `AMMS Daily Summary — ${today()}`, body);
    });
  } catch(e) {
    Logger.log('Email error: ' + e.message);
  }
}

// ============================================================
//  INSTALL TRIGGER (run once)
// ============================================================

function installDailyTrigger() {
  // Delete existing daily triggers
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'dailyAutomation') {
      ScriptApp.deleteTrigger(t);
    }
  });
  // Create new daily trigger at 8 AM
  ScriptApp.newTrigger('dailyAutomation')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();
  Logger.log('Daily trigger installed at 8 AM.');
}

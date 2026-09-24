/**
 * East Sepik Students Association Presidential Election backend.
 * Deploy this file as a Google Apps Script Web App.
 */

const CONFIG = {
  TAB_NAME: 'Votes',
  SPREADSHEET_ID: '1UXzh6hh2oOZ_2Yd6jSWJDvjMR4tk1l4mz1seZa17X7c',
  TOKEN_EXPIRY_HOURS: 8,
  ADMIN_ACCOUNTS: {
    es_association_admin: {
      password: 'ESAS2027#Admin!Pass',
      name: 'East Sepik Students Association Admin',
      role: 'Election Administrator'
    },
    electoral_chair: {
      password: 'ESAS2027#Chair!Secure',
      name: 'Electoral Committee Chair',
      role: 'Chief Returning Officer'
    },
    system_admin: {
      password: 'IBSU#ESAS2027!Admin',
      name: 'System Administrator',
      role: 'Platform Manager'
    }
  }
};

function doGet(e) {
  const action = e && e.parameter && e.parameter.action ? e.parameter.action : 'getDashboardData';
  const callback = e && e.parameter && e.parameter.callback;
  const output = callback ? jsonpOutput(callback) : jsonOutput();

  try {
    if (action === 'ping') {
      return output({
        status: 'success',
        message: 'East Sepik Students Association Presidential Election backend is live and operational.',
        timestamp: new Date().toISOString()
      });
    }

    if (action === 'login') {
      return handleAdminLogin({
        username: e.parameter.username,
        password: e.parameter.password
      }, output);
    }

    if (action === 'submitVote') {
      return handleVoteSubmission({
        voterId: e.parameter.voterId,
        voterName: e.parameter.voterName,
        voterEmail: e.parameter.voterEmail,
        president: e.parameter.president,
        vicePresident: '',
        secretary: '',
        treasurer: ''
      }, output);
    }

    if (action === 'getDashboardData') {
      const token = e.parameter.token || '';
      if (!isValidSession(token)) {
        return output({ status: 'error', message: 'Unauthorized or expired session.' });
      }
      return output({ status: 'success', data: fetchVotingDataFromSheets() });
    }

    return output({ status: 'error', message: 'Invalid GET action.' });
  } catch (err) {
    return output({ status: 'error', message: 'Server error: ' + err.message });
  }
}

function doPost(e) {
  const output = jsonOutput();
  let payload;

  try {
    payload = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return output({ status: 'error', message: 'Malformed JSON payload.' });
  }

  try {
    if (payload.action === 'login') {
      return handleAdminLogin(payload, output);
    }
    if (payload.action === 'submitVote') {
      return handleVoteSubmission(payload, output);
    }
    return output({ status: 'error', message: 'Unknown POST action.' });
  } catch (err) {
    return output({ status: 'error', message: 'Server error: ' + err.message });
  }
}

function handleAdminLogin(payload, output) {
  const username = String(payload.username || '').trim().toLowerCase();
  const password = String(payload.password || '');
  const account = CONFIG.ADMIN_ACCOUNTS[username];

  if (!account || account.password !== password) {
    return output({ status: 'error', message: 'Invalid admin username or passcode.' });
  }

  const token = Utilities.getUuid();
  const expiresAt = Date.now() + CONFIG.TOKEN_EXPIRY_HOURS * 60 * 60 * 1000;
  PropertiesService.getScriptProperties().setProperty(
    'SESSION_TOKEN_' + token,
    JSON.stringify({ username: username, expiresAt: expiresAt })
  );

  return output({
    status: 'success',
    token: token,
    expiresAt: expiresAt,
    user: {
      username: username,
      name: account.name,
      role: account.role
    }
  });
}

function isValidSession(token) {
  if (!token) return false;

  const value = PropertiesService.getScriptProperties().getProperty('SESSION_TOKEN_' + token);
  if (!value) return false;

  try {
    const session = JSON.parse(value);
    if (Date.now() >= Number(session.expiresAt)) {
      PropertiesService.getScriptProperties().deleteProperty('SESSION_TOKEN_' + token);
      return false;
    }
    return true;
  } catch (err) {
    return false;
  }
}

function handleVoteSubmission(payload, output) {
  const voterId = String(payload.voterId || '').trim();
  const voterName = String(payload.voterName || '').trim();
  const voterEmail = String(payload.voterEmail || '').trim();
  const president = String(payload.president || '').trim();
  const allowedCandidates = new Set(['Jubal Suagu', 'Joesen Weikun']);

  if (!voterId || !voterName || !voterEmail || !president) {
    return output({ status: 'error', message: 'All voter details and one presidential candidate selection are required.' });
  }

  if (!allowedCandidates.has(president)) {
    return output({ status: 'error', message: 'Invalid presidential candidate selected.' });
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return output({ status: 'error', message: 'Server is busy. Please try again.' });
  }

  try {
    const sheet = getVotesSheet();
    const rows = sheet.getDataRange().getValues();
    const normalizedId = voterId.toLowerCase();

    for (let index = 1; index < rows.length; index += 1) {
      if (String(rows[index][1] || '').trim().toLowerCase() === normalizedId) {
        return output({ status: 'error', message: 'This Student ID has already submitted a ballot.' });
      }
    }

    const receipt = 'ESAS-2027-' + Math.floor(100000 + Math.random() * 900000);
    const timestamp = new Date();
    sheet.appendRow([
      timestamp,
      voterId,
      voterName,
      voterEmail,
      president,
      '',
      '',
      '',
      receipt
    ]);

    return output({
      status: 'success',
      message: 'Vote successfully recorded.',
      receipt: receipt,
      timestamp: timestamp.toISOString()
    });
  } finally {
    lock.releaseLock();
  }
}

function getTargetSpreadsheet() {
  const spreadsheetId = String(CONFIG.SPREADSHEET_ID || '').trim();
  if (spreadsheetId && spreadsheetId !== 'PASTE_YOUR_GOOGLE_SHEET_ID_HERE') {
    return SpreadsheetApp.openById(spreadsheetId);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getVotesSheet() {
  const spreadsheet = getTargetSpreadsheet();
  let sheet = spreadsheet.getSheetByName(CONFIG.TAB_NAME);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(CONFIG.TAB_NAME);
    sheet.appendRow([
      'Timestamp',
      'Voter ID',
      'Full Name',
      'Email',
      'President',
      'Vice President',
      'Secretary',
      'Treasurer',
      'Receipt Token'
    ]);
    sheet.getRange('1:1').setFontWeight('bold');
  }

  return sheet;
}

function fetchVotingDataFromSheets() {
  const sheet = getVotesSheet();
  const rows = sheet.getDataRange().getValues().slice(1);
  const candidateCounts = {
    President: {}
  };

  const votes = rows.map(function(row) {
    const vote = {
      timestamp: row[0] ? new Date(row[0]).toISOString() : new Date().toISOString(),
      voterId: String(row[1] || ''),
      voterName: String(row[2] || ''),
      voterEmail: String(row[3] || ''),
      president: String(row[4] || 'Abstain'),
      vicePresident: String(row[5] || 'Abstain'),
      secretary: String(row[6] || 'Abstain'),
      treasurer: String(row[7] || 'Abstain'),
      receipt: String(row[8] || '')
    };

    countCandidateVote(candidateCounts.President, vote.president);
    return vote;
  });

  return {
    votes: votes,
    totalVotes: votes.length,
    candidateCounts: candidateCounts,
    timestamp: new Date().toISOString()
  };
}

function countCandidateVote(positionCounts, candidateName) {
  if (candidateName && candidateName !== 'Abstain') {
    positionCounts[candidateName] = (positionCounts[candidateName] || 0) + 1;
  }
}

function jsonOutput() {
  return function(data) {
    return ContentService
      .createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);
  };
}

function jsonpOutput(callback) {
  const safeCallback = String(callback).match(/^[A-Za-z_$][0-9A-Za-z_$]*$/) ? callback : 'invalidCallback';

  return function(data) {
    return ContentService
      .createTextOutput(safeCallback + '(' + JSON.stringify(data) + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  };
}

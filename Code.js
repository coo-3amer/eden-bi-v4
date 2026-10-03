const SPREADSHEET_ID = '1d7NLDM1Q5YCxFHcCxdifmTBTc7tSmScuVMapI_wnIWo';
const DEALS_GID = 835779772;
const HEADER_ROW = 3;
const DATA_START_ROW = 4;

const BROKER_SPREADSHEET_ID = '1_a4Vwtkj72C2Adbtv8qeEV6yKudPVKbnLTbl6_JMsDA';
const BROKER_SHEET_NAME = 'EGY - All Brokers';
const BROKER_HEADER_ROW = 1;
const BROKER_DATA_START_ROW = 2;

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function doGet(e) {
  // JSON API for the dashboard page hosted on GitHub Pages (see Api.js).
  if (e && e.parameter && e.parameter.api) return apiDoGet_(e);

  // Ensure the lightweight Users sheet exists before the login screen loads.
  // This does not read dashboard data and will not slow down the dashboard.
  getUsersSheet_();

  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('EDEN DEVELOPMENT DASHBOARD')
    .setFaviconUrl('https://raw.githubusercontent.com/coo-3amer/eden-bi-v4/redesign/web/favicon.png')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


// ===== SIMPLE LOGIN SYSTEM =====
// Users are stored in a lightweight sheet named "Users" inside the Transactions spreadsheet.
// Required columns: Username | Password | Name | Role | Status
// Default first user is created automatically if the sheet does not exist:
// Username: admin   Password: admin123

function getUsersSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName('Users');

  if (!sheet) {
    sheet = ss.insertSheet('Users');
    sheet.getRange(1, 1, 1, 5).setValues([['Username', 'Password', 'Name', 'Role', 'Status']]);
    sheet.getRange(2, 1, 1, 5).setValues([['admin', 'admin123', 'System Admin', 'Super Admin', 'Active']]);
    sheet.setFrozenRows(1);
  }

  return sheet;
}


function setupUsersSheet() {
  // Manual setup function. Run once from Apps Script if you want to create/check the Users tab.
  const sheet = getUsersSheet_();
  return {
    success: true,
    sheetName: sheet.getName(),
    message: 'Users sheet is ready inside the Transactions spreadsheet.'
  };
}

function loginUser(username, password) {
  username = String(username || '').trim().toLowerCase();
  password = String(password || '').trim();

  if (!username || !password) {
    return { success: false, message: 'Username and password are required.' };
  }

  const sheet = getUsersSheet_();
  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) {
    return { success: false, message: 'No active users found.' };
  }

  const headers = values[0].map(h => String(h || '').trim());
  const idx = name => headers.indexOf(name);

  const uIdx = idx('Username');
  const pIdx = idx('Password');
  const nIdx = idx('Name');
  const rIdx = idx('Role');
  const sIdx = idx('Status');

  if (uIdx === -1 || pIdx === -1) {
    throw new Error('Users sheet must contain Username and Password columns.');
  }

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const rowUser = String(row[uIdx] || '').trim().toLowerCase();
    const rowPass = String(row[pIdx] || '').trim();
    const status = sIdx > -1 ? String(row[sIdx] || '').trim().toLowerCase() : 'active';

    if (rowUser === username && rowPass === password && status !== 'inactive') {
      const token = Utilities.getUuid();
      const user = {
        username: rowUser,
        name: nIdx > -1 ? row[nIdx] : rowUser,
        role: rIdx > -1 ? row[rIdx] : 'User'
      };

      CacheService.getScriptCache().put('LOGIN_' + token, JSON.stringify(user), 21600); // 6 hours

      return {
        success: true,
        token: token,
        user: user
      };
    }
  }

  return { success: false, message: 'Invalid username or password.' };
}

function validateAuthToken_(token) {
  token = String(token || '').trim();
  if (!token) throw new Error('Access denied. Please login again.');

  const cached = CacheService.getScriptCache().get('LOGIN_' + token);
  if (!cached) throw new Error('Session expired. Please login again.');

  return JSON.parse(cached);
}

function logoutUser(token) {
  token = String(token || '').trim();
  if (token) CacheService.getScriptCache().remove('LOGIN_' + token);
  return { success: true };
}


// ===== USER MANAGEMENT FROM DASHBOARD =====
// Lightweight CRUD for Users tab. Only Super Admin / Admin can manage users.
function assertUserAdmin_(token) {
  const user = validateAuthToken_(token);
  const role = String(user.role || '').toLowerCase();
  if (role !== 'super admin' && role !== 'admin') {
    throw new Error('Access denied. Admin only.');
  }
  return user;
}

function getUsersList(authToken) {
  assertUserAdmin_(authToken);
  const sheet = getUsersSheet_();
  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) return [];

  const headers = values[0].map(h => String(h || '').trim());
  const idx = name => headers.indexOf(name);
  const uIdx = idx('Username');
  const nIdx = idx('Name');
  const rIdx = idx('Role');
  const sIdx = idx('Status');

  return values.slice(1).filter(r => r.join('').trim() !== '').map((r, i) => ({
    rowNumber: i + 2,
    username: uIdx > -1 ? r[uIdx] : '',
    name: nIdx > -1 ? r[nIdx] : '',
    role: rIdx > -1 ? r[rIdx] : '',
    status: sIdx > -1 ? r[sIdx] : ''
  }));
}

function createUser(authToken, userData) {
  assertUserAdmin_(authToken);
  userData = userData || {};
  const username = String(userData.username || '').trim().toLowerCase();
  const password = String(userData.password || '').trim();
  const name = String(userData.name || '').trim();
  const role = String(userData.role || 'Viewer').trim();
  const status = String(userData.status || 'Active').trim();

  if (!username || !password || !name) {
    throw new Error('Username, password and name are required.');
  }

  const sheet = getUsersSheet_();
  const values = sheet.getDataRange().getDisplayValues();
  const headers = values[0].map(h => String(h || '').trim());
  const uIdx = headers.indexOf('Username');
  if (uIdx === -1) throw new Error('Users sheet must contain Username column.');

  const exists = values.slice(1).some(r => String(r[uIdx] || '').trim().toLowerCase() === username);
  if (exists) throw new Error('Username already exists.');

  sheet.appendRow([username, password, name, role, status]);
  return { success: true, message: 'User created successfully.' };
}

function updateUser(authToken, userData) {
  const admin = assertUserAdmin_(authToken);
  userData = userData || {};
  const username = String(userData.username || '').trim().toLowerCase();
  if (!username) throw new Error('Username is required.');

  const sheet = getUsersSheet_();
  const values = sheet.getDataRange().getDisplayValues();
  const headers = values[0].map(h => String(h || '').trim());
  const idx = name => headers.indexOf(name) + 1;
  const uCol = idx('Username');
  if (uCol < 1) throw new Error('Users sheet must contain Username column.');

  let rowNumber = -1;
  for (let i = 2; i <= values.length; i++) {
    if (String(values[i - 1][uCol - 1] || '').trim().toLowerCase() === username) {
      rowNumber = i;
      break;
    }
  }
  if (rowNumber === -1) throw new Error('User not found.');

  const setByHeader = (header, value) => {
    const c = idx(header);
    if (c > 0 && value !== undefined) sheet.getRange(rowNumber, c).setValue(value);
  };

  setByHeader('Name', String(userData.name || '').trim());
  if (String(userData.password || '').trim()) setByHeader('Password', String(userData.password || '').trim());
  setByHeader('Role', String(userData.role || 'Viewer').trim());
  setByHeader('Status', String(userData.status || 'Active').trim());

  return { success: true, message: 'User updated successfully.' };
}

function deleteUser(authToken, username) {
  const admin = assertUserAdmin_(authToken);
  username = String(username || '').trim().toLowerCase();
  if (!username) throw new Error('Username is required.');
  if (username === String(admin.username || '').trim().toLowerCase()) {
    throw new Error('You cannot delete your own user while logged in.');
  }

  const sheet = getUsersSheet_();
  const values = sheet.getDataRange().getDisplayValues();
  const headers = values[0].map(h => String(h || '').trim());
  const uIdx = headers.indexOf('Username');
  if (uIdx === -1) throw new Error('Users sheet must contain Username column.');

  for (let i = 2; i <= values.length; i++) {
    if (String(values[i - 1][uIdx] || '').trim().toLowerCase() === username) {
      sheet.deleteRow(i);
      return { success: true, message: 'User deleted successfully.' };
    }
  }

  throw new Error('User not found.');
}

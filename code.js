const SHEETS = { USERS: 'Users', CATEGORIES: 'Categories', SUPPLIERS: 'Suppliers', PURCHASES: 'Purchases', ITEM_STOCKS: 'Item_Stocks', CUSTOMERS: 'Customers', SALES: 'Sales', SALE_ITEMS: 'Sale_Items', PAYMENTS: 'Payments', EXPENSES: 'Expenses', IMPORT_LOGS: 'Import_Logs', SETTINGS: 'Settings', LOGS: 'Activity_Logs' };

// cols: 0=id, 1=full_name, 2=email, 3=phone, 4=pwd, 5=role, 6=avatar, 7=is_active, 8=created_at, 9=updated_at, 10=otp, 11=otp_expires
const U = { ID: 0, NAME: 1, EMAIL: 2, PHONE: 3, PWD: 4, ROLE: 5, AVATAR: 6, ACTIVE: 7, CREATED: 8, UPDATED: 9, OTP: 10, OTP_EXP: 11 };

// cols: 0=id, 1=name, 2=description, 3=is_active, 4=created_by, 5=created_at
const C = { ID: 0, NAME: 1, DESC: 2, ACTIVE: 3, CREATED_BY: 4, CREATED: 5 };

// cols: 0=id, 1=name, 2=phone, 3=address, 4=is_active, 5=created_by, 6=created_at
const SP = { ID: 0, NAME: 1, PHONE: 2, ADDR: 3, ACTIVE: 4, CREATED_BY: 5, CREATED: 6 };

// cols: 0=id, 1=item_code, 2=item_name, 3=category_id, 4=purchase_id, 5=buy_rate, 6=buy_price, 7=sell_rate, 8=sell_price, 9=status, 10=notes, 11=created_by, 12=created_at, 13=updated_at, 14=qty, 15=image
const WS = { ID: 0, ITEM_CODE: 1, ITEM_NAME: 2, CAT_ID: 3, PUR_ID: 4, BUY_RATE: 5, BUY_PRICE: 6, SELL_RATE: 7, SELL_PRICE: 8, STATUS: 9, NOTES: 10, CREATED_BY: 11, CREATED: 12, UPDATED: 13, QTY: 14, IMG: 15 };

// cols: 0=id, 1=purchase_no, 2=supplier_id, 3=category_id, 4=total_qty, 5=avg_rate, 6=total_amount, 7=paid_amount, 8=due_amount, 9=purchase_date, 10=notes, 11=status, 12=created_by, 13=created_at, 14=updated_at
const PU = { ID: 0, NO: 1, SUPPLIER_ID: 2, CAT_ID: 3, QTY: 4, RATE: 5, TOTAL: 6, PAID: 7, DUE: 8, DATE: 9, NOTES: 10, STATUS: 11, CREATED_BY: 12, CREATED: 13, UPDATED: 14 };

// cols: 0=id, 1=name, 2=phone, 3=address, 4=total_purchase, 5=total_paid, 6=total_due, 7=is_active, 8=created_by, 9=created_at
const CU = { ID: 0, NAME: 1, PHONE: 2, ADDR: 3, TOTAL: 4, PAID: 5, DUE: 6, ACTIVE: 7, CREATED_BY: 8, CREATED: 9 };

// unified payments
const PAY = { ID: 0, SALE_ID: 1, PUR_ID: 2, TYPE: 3, AMT: 4, METHOD: 5, REF: 6, DATE: 7, NOTES: 8, CREATED_BY: 9, CREATED: 10 };

// cols: 0=id, 1=invoice_no, 2=customer_id, 3=total_items, 4=subtotal, 5=discount, 6=grand_total, 7=paid_amount, 8=due_amount, 9=payment_method, 10=sale_date, 11=status, 12=notes, 13=created_by, 14=created_at, 15=updated_at
const SL = { ID: 0, INV_NO: 1, CUST_ID: 2, ITEMS: 3, SUBTOTAL: 4, DISC: 5, TOTAL: 6, PAID: 7, DUE: 8, METHOD: 9, DATE: 10, STATUS: 11, NOTES: 12, CREATED_BY: 13, CREATED: 14, UPDATED: 15 };

// sale items: 0=id, 1=sale_id, 2=item_stock_id, 3=item_code, 4=item_name, 5=qty, 6=rate, 7=line_total
const SI = { ID: 0, SALE_ID: 1, WS_ID: 2, ITEM_CODE: 3, ITEM_NAME: 4, QTY: 5, RATE: 6, TOTAL: 7 };

// expenses
const EX = { ID: 0, TITLE: 1, CAT: 2, AMT: 3, DATE: 4, NOTES: 5, CREATED_BY: 6, CREATED: 7 };

// import_logs
const IL = { ID: 0, FILE: 1, CAT_ID: 2, TOTAL: 3, SUCCESS: 4, FAILED: 5, STATUS: 6, ERRORS: 7, CREATED_BY: 8, CREATED: 9 };

// settings
const ST = { ID: 0, KEY: 1, VAL: 2, UPDATED_BY: 3, UPDATED: 4 };

// logs
const L = { ID: 0, UID: 1, UNAME: 2, ACTION: 3, TABLE: 4, RID: 5, DETAILS: 6, CREATED: 7 };

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('CDS Publication POS')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}

function include(filename) { return HtmlService.createHtmlOutputFromFile(filename).getContent(); }
function getSheet(name) { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name); }
function getSheetData(name) {
  const sh = getSheet(name);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
}
function getNextId(name) {
  const sh = getSheet(name);
  if (!sh || sh.getLastRow() < 2) return 1;
  const ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().flat().filter(id => id !== '');
  return ids.length ? Math.max(...ids) + 1 : 1;
}
function findRowByValue(name, col, val) {
  const data = getSheetData(name);
  for (let i = 0; i < data.length; i++) {
    if (data[i][col] == val) return { row: i + 2, data: data[i] };
  }
  return null;
}
function ts() { return new Date().toISOString(); }
function isActive(val) { return val === 1 || val === true || val === '1'; }
function sanitizeHtml(str) { return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;'); }
function getUsernameById(uid) { const r = findRowByValue(SHEETS.USERS, U.ID, parseInt(uid)); return r ? r.data[U.NAME] : 'Unknown'; }

function buildItemStatsMap() {
  const m = {};
  try {
    const data = getSheetData(SHEETS.ITEM_STOCKS);
    data.forEach(r => {
      const catId = r[WS.CAT_ID];
      const qty = parseInt(r[WS.QTY]) || 0;
      if (!m[catId]) m[catId] = { total: 0, available: 0, sold: 0, stockValue: 0, itemCount: 0 };
      m[catId].total += qty;
      m[catId].itemCount += 1;
      if (r[WS.STATUS] === 'available') { m[catId].available += qty; m[catId].stockValue += (parseFloat(r[WS.SELL_PRICE]) || 0); }
      else if (r[WS.STATUS] === 'sold') m[catId].sold += qty;
    });
  } catch (e) { }
  return m;
}

function getItemStockCount(catId) {
  try { const data = getSheetData(SHEETS.ITEM_STOCKS); return data.filter(r => r[WS.CAT_ID] == catId).length; } catch (e) { return 0; }
}

function getDriveFolder(name) { const it = DriveApp.getFoldersByName(name); return it.hasNext() ? it.next() : DriveApp.createFolder(name); }
function uploadToDrive(b64, fileName, mimeType, folderName) {
  const folder = getDriveFolder(folderName || 'ItemPOS_Images');
  const blob = Utilities.newBlob(Utilities.base64Decode(b64), mimeType, fileName);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getId();
}
function deleteFromDrive(fileId) { try { if (fileId) DriveApp.getFileById(fileId).setTrashed(true); } catch (e) { } }

// Auth
function login(email, password) {
  try {
    if (!email || !password) return { success: false, message: 'Email and password required' };
    const r = findRowByValue(SHEETS.USERS, U.EMAIL, email.trim().toLowerCase());
    if (!r) return { success: false, message: 'Invalid email or password' };
    const u = r.data;
    if (!isActive(u[U.ACTIVE])) return { success: false, message: 'Account inactive. Contact admin.' };
    if (password !== u[U.PWD]) return { success: false, message: 'Invalid email or password' };
    getSheet(SHEETS.USERS).getRange(r.row, U.UPDATED + 1).setValue(ts());
    logActivity(u[U.ID], 'LOGIN', 'Users', u[U.ID], '');
    return { success: true, message: 'Login successful', data: { id: u[U.ID], full_name: u[U.NAME], email: u[U.EMAIL], role: u[U.ROLE], avatar: u[U.AVATAR] || '' } };
  } catch (e) { return { success: false, message: 'Login failed' }; }
}

// Users
function getUsers(userId, role) {
  try {
    if (role !== 'admin') return { success: false, message: 'Access denied' };
    const data = getSheetData(SHEETS.USERS);
    const users = data.map(r => ({
      id: r[U.ID], full_name: r[U.NAME], email: r[U.EMAIL], phone: r[U.PHONE] || '', role: r[U.ROLE], avatar: r[U.AVATAR] || '',
      is_active: isActive(r[U.ACTIVE]) ? 1 : 0, created_at: r[U.CREATED] instanceof Date ? r[U.CREATED].toISOString() : r[U.CREATED],
      updated_at: r[U.UPDATED] instanceof Date ? r[U.UPDATED].toISOString() : r[U.UPDATED] || ''
    }));
    return { success: true, data: users.reverse() };
  } catch (e) { return { success: false, message: 'Failed to load users' }; }
}

function addUser(userData, userId, role) {
  try {
    if (role !== 'admin') return { success: false, message: 'Access denied' };
    const { full_name, email, phone, password, userRole, avatarData } = userData;
    if (!full_name || !email || !password) return { success: false, message: 'Name, email and password required' };
    if (password.length < 6) return { success: false, message: 'Password min 6 characters' };
    if (findRowByValue(SHEETS.USERS, U.EMAIL, email.trim().toLowerCase())) return { success: false, message: 'Email already exists' };

    let avatarId = '';
    if (avatarData && avatarData.data) avatarId = uploadToDrive(avatarData.data, avatarData.name, avatarData.type);

    const sh = getSheet(SHEETS.USERS);
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const newId = getNextId(SHEETS.USERS);
      const now = ts();
      sh.appendRow([newId, full_name.trim(), email.trim().toLowerCase(), phone || '', password, userRole || 'cashier', avatarId, 1, now, now, '', '']);
      logActivity(userId, 'CREATE', 'Users', newId, 'Added: ' + full_name.trim());
      return { success: true, message: 'User added successfully', data: { id: newId } };
    } finally { lock.releaseLock(); }
  } catch (e) { return { success: false, message: 'Failed to add user' }; }
}

function updateUser(userData, userId, role) {
  try {
    if (role !== 'admin') return { success: false, message: 'Access denied' };
    const { id, full_name, email, phone, password, userRole, is_active, avatarData, removeAvatar } = userData;
    const r = findRowByValue(SHEETS.USERS, U.ID, parseInt(id));
    if (!r) return { success: false, message: 'User not found' };

    const ec = findRowByValue(SHEETS.USERS, U.EMAIL, email.trim().toLowerCase());
    if (ec && ec.data[U.ID] != parseInt(id)) return { success: false, message: 'Email already in use' };

    const row = r.data.slice();
    row[U.NAME] = full_name.trim(); row[U.EMAIL] = email.trim().toLowerCase(); row[U.PHONE] = phone || '';
    row[U.ROLE] = userRole || 'cashier'; row[U.ACTIVE] = (is_active == 1 || is_active === true) ? 1 : 0;
    if (password && password.length >= 6) row[U.PWD] = password;

    if (avatarData && avatarData.data) {
      if (row[U.AVATAR]) deleteFromDrive(row[U.AVATAR]);
      row[U.AVATAR] = uploadToDrive(avatarData.data, avatarData.name, avatarData.type);
    } else if (removeAvatar) {
      if (row[U.AVATAR]) deleteFromDrive(row[U.AVATAR]);
      row[U.AVATAR] = '';
    }

    row[U.UPDATED] = ts();
    getSheet(SHEETS.USERS).getRange(r.row, 1, 1, row.length).setValues([row]);
    logActivity(userId, 'UPDATE', 'Users', parseInt(id), 'Updated: ' + full_name.trim());
    return { success: true, message: 'User updated successfully' };
  } catch (e) { return { success: false, message: 'Failed to update user' }; }
}

function deleteUser(id, userId, role) {
  try {
    if (role !== 'admin') return { success: false, message: 'Access denied' };
    if (parseInt(id) === parseInt(userId)) return { success: false, message: 'Cannot delete your own account' };
    const r = findRowByValue(SHEETS.USERS, U.ID, parseInt(id));
    if (!r) return { success: false, message: 'User not found' };
    if (r.data[U.AVATAR]) deleteFromDrive(r.data[U.AVATAR]);
    const name = r.data[U.NAME];
    getSheet(SHEETS.USERS).deleteRow(r.row);
    logActivity(userId, 'DELETE', 'Users', parseInt(id), 'Deleted: ' + name);
    return { success: true, message: 'User deleted successfully' };
  } catch (e) { return { success: false, message: 'Failed to delete user' }; }
}

function toggleUserStatus(id, userId, role) {
  try {
    if (role !== 'admin') return { success: false, message: 'Access denied' };
    if (parseInt(id) === parseInt(userId)) return { success: false, message: 'Cannot toggle your own status' };
    const r = findRowByValue(SHEETS.USERS, U.ID, parseInt(id));
    if (!r) return { success: false, message: 'User not found' };
    const row = r.data.slice();
    const was = isActive(row[U.ACTIVE]);
    row[U.ACTIVE] = was ? 0 : 1;
    row[U.UPDATED] = ts();
    getSheet(SHEETS.USERS).getRange(r.row, 1, 1, row.length).setValues([row]);
    const status = was ? 'Deactivated' : 'Activated';
    logActivity(userId, 'TOGGLE_STATUS', 'Users', parseInt(id), status + ': ' + r.data[U.NAME]);
    return { success: true, message: 'User ' + status.toLowerCase(), data: { is_active: was ? 0 : 1 } };
  } catch (e) { return { success: false, message: 'Failed to toggle status' }; }
}

function getProfile(userId) {
  try {
    const r = findRowByValue(SHEETS.USERS, U.ID, parseInt(userId));
    if (!r) return { success: false, message: 'User not found' };
    const u = r.data;
    return { success: true, data: { id: u[U.ID], full_name: u[U.NAME], email: u[U.EMAIL], phone: u[U.PHONE] || '', role: u[U.ROLE], avatar: u[U.AVATAR] || '', is_active: isActive(u[U.ACTIVE]) ? 1 : 0, created_at: u[U.CREATED] instanceof Date ? u[U.CREATED].toISOString() : u[U.CREATED] } };
  } catch (e) { return { success: false, message: 'Failed to load profile' }; }
}

function updateProfile(profileData, userId) {
  try {
    const { full_name, phone, avatarData, removeAvatar } = profileData;
    const r = findRowByValue(SHEETS.USERS, U.ID, parseInt(userId));
    if (!r) return { success: false, message: 'User not found' };
    const row = r.data.slice();
    row[U.NAME] = full_name.trim(); row[U.PHONE] = phone || '';
    if (avatarData && avatarData.data) {
      if (row[U.AVATAR]) deleteFromDrive(row[U.AVATAR]);
      row[U.AVATAR] = uploadToDrive(avatarData.data, avatarData.name, avatarData.type);
    } else if (removeAvatar) {
      if (row[U.AVATAR]) deleteFromDrive(row[U.AVATAR]);
      row[U.AVATAR] = '';
    }
    row[U.UPDATED] = ts();
    getSheet(SHEETS.USERS).getRange(r.row, 1, 1, row.length).setValues([row]);
    logActivity(userId, 'UPDATE', 'Users', parseInt(userId), 'Updated own profile');
    return { success: true, message: 'Profile updated', data: { full_name: full_name.trim(), phone: phone || '', avatar: row[U.AVATAR] } };
  } catch (e) { return { success: false, message: 'Failed to update profile' }; }
}

function changePassword(userId, currentPassword, newPassword) {
  try {
    if (!newPassword || newPassword.length < 6) return { success: false, message: 'New password min 6 characters' };
    const r = findRowByValue(SHEETS.USERS, U.ID, parseInt(userId));
    if (!r) return { success: false, message: 'User not found' };
    if (currentPassword !== r.data[U.PWD]) return { success: false, message: 'Current password incorrect' };
    getSheet(SHEETS.USERS).getRange(r.row, U.PWD + 1).setValue(newPassword);
    logActivity(userId, 'CHANGE_PWD', 'Users', parseInt(userId), '');
    return { success: true, message: 'Password changed successfully' };
  } catch (e) { return { success: false, message: 'Failed to change password' }; }
}

// Categories (Main Organizer)
function getCategories(userId, role) {
  try {
    const data = getSheetData(SHEETS.CATEGORIES);
    const uData = getSheetData(SHEETS.USERS);
    const uMap = {}; uData.forEach(r => { uMap[r[U.ID]] = r[U.NAME]; });
    const wsMap = buildItemStatsMap();

    const cats = data.map(r => {
      const stats = wsMap[r[C.ID]] || { total: 0, available: 0, sold: 0, stockValue: 0, itemCount: 0 };
      return {
        id: r[C.ID], name: r[C.NAME], description: r[C.DESC] || '',
        is_active: isActive(r[C.ACTIVE]) ? 1 : 0, created_by: r[C.CREATED_BY],
        created_by_name: uMap[r[C.CREATED_BY]] || 'Unknown',
        item_count: stats.itemCount,
        total_pieces: stats.total, available_pieces: stats.available, sold_pieces: stats.sold,
        stock_value: Math.round(stats.stockValue * 100) / 100,
        created_at: r[C.CREATED] instanceof Date ? r[C.CREATED].toISOString() : r[C.CREATED]
      };
    });
    return { success: true, data: cats.reverse() };
  } catch (e) { return { success: false, message: 'Failed to load categories' }; }
}

function addCategory(catData, userId, role) {
  try {
    if (role !== 'admin' && role !== 'manager') return { success: false, message: 'Access denied' };
    const { name, description } = catData;
    if (findRowByValue(SHEETS.CATEGORIES, C.NAME, name.trim())) return { success: false, message: 'Category name already exists' };
    const sh = getSheet(SHEETS.CATEGORIES);
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const newId = getNextId(SHEETS.CATEGORIES);
      sh.appendRow([newId, name.trim(), description || '', 1, userId, ts()]);
      logActivity(userId, 'CREATE', 'Categories', newId, 'Added: ' + name.trim());
      return { success: true, message: 'Category added', data: { id: newId } };
    } finally { lock.releaseLock(); }
  } catch (e) { return { success: false, message: 'Failed to add category' }; }
}

function updateCategory(catData, userId, role) {
  try {
    if (role !== 'admin' && role !== 'manager') return { success: false, message: 'Access denied' };
    const { id, name, description, is_active } = catData;
    const r = findRowByValue(SHEETS.CATEGORIES, C.ID, parseInt(id));
    if (!r) return { success: false, message: 'Category not found' };
    const nc = findRowByValue(SHEETS.CATEGORIES, C.NAME, name.trim());
    if (nc && nc.data[C.ID] != parseInt(id)) return { success: false, message: 'Category name already exists' };
    const row = r.data.slice();
    row[C.NAME] = name.trim(); row[C.DESC] = description || ''; row[C.ACTIVE] = (is_active == 1 || is_active === true) ? 1 : 0;
    getSheet(SHEETS.CATEGORIES).getRange(r.row, 1, 1, row.length).setValues([row]);
    logActivity(userId, 'UPDATE', 'Categories', parseInt(id), 'Updated: ' + name.trim());
    return { success: true, message: 'Category updated' };
  } catch (e) { return { success: false, message: 'Failed to update category' }; }
}

function deleteCategory(id, userId, role) {
  try {
    if (role !== 'admin') return { success: false, message: 'Only admin can delete' };
    const r = findRowByValue(SHEETS.CATEGORIES, C.ID, parseInt(id));
    if (!r) return { success: false, message: 'Category not found' };
    if (getItemStockCount(parseInt(id)) > 0) return { success: false, message: 'Cannot delete — has linked items' };
    const name = r.data[C.NAME];
    getSheet(SHEETS.CATEGORIES).deleteRow(r.row);
    logActivity(userId, 'DELETE', 'Categories', parseInt(id), 'Deleted: ' + name);
    return { success: true, message: 'Category deleted' };
  } catch (e) { return { success: false, message: 'Failed to delete' }; }
}

function toggleCategoryStatus(id, userId, role) {
  try {
    if (role !== 'admin' && role !== 'manager') return { success: false, message: 'Access denied' };
    const r = findRowByValue(SHEETS.CATEGORIES, C.ID, parseInt(id));
    if (!r) return { success: false, message: 'Category not found' };
    const row = r.data.slice();
    const was = isActive(row[C.ACTIVE]);
    row[C.ACTIVE] = was ? 0 : 1;
    getSheet(SHEETS.CATEGORIES).getRange(r.row, 1, 1, row.length).setValues([row]);
    return { success: true, message: 'Status updated' };
  } catch (e) { return { success: false, message: 'Failed to toggle status' }; }
}

function getCategoriesForDropdown() {
  try {
    const data = getSheetData(SHEETS.CATEGORIES);
    return { success: true, data: data.filter(r => isActive(r[C.ACTIVE])).map(r => ({ id: r[C.ID], name: r[C.NAME] })) };
  } catch (e) { return { success: true, data: [] }; }
}
// Suppliers
function buildPurchaseStatsMap() {
  const m = {};
  try {
    const data = getSheetData(SHEETS.PURCHASES);
    data.forEach(r => {
      const sid = r[PU.SUPPLIER_ID];
      if (!m[sid]) m[sid] = { count: 0, totalAmt: 0, paidAmt: 0 };
      m[sid].count++;
      m[sid].totalAmt += (parseFloat(r[PU.TOTAL]) || 0);
      m[sid].paidAmt += (parseFloat(r[PU.PAID]) || 0);
    });
  } catch (e) { }
  return m;
}

function getPurchaseCountForSupplier(suppId) {
  try { const data = getSheetData(SHEETS.PURCHASES); return data.filter(r => r[PU.SUPPLIER_ID] == suppId).length; } catch (e) { return 0; }
}

function getSuppliers(userId, role) {
  try {
    if (role === 'cashier') return { success: false, message: 'Access denied' };
    const data = getSheetData(SHEETS.SUPPLIERS);
    const puMap = buildPurchaseStatsMap();
    const uData = getSheetData(SHEETS.USERS);
    const uMap = {}; uData.forEach(r => { uMap[r[U.ID]] = r[U.NAME]; });

    const suppliers = data.map(r => {
      const pu = puMap[r[SP.ID]] || { count: 0, totalAmt: 0, paidAmt: 0 };
      return {
        id: r[SP.ID], name: r[SP.NAME], phone: r[SP.PHONE] || '', address: r[SP.ADDR] || '',
        is_active: isActive(r[SP.ACTIVE]) ? 1 : 0, created_by_name: uMap[r[SP.CREATED_BY]] || 'Unknown',
        purchase_count: pu.count, total_paid: Math.round(pu.paidAmt * 100) / 100,
        total_due: Math.round((pu.totalAmt - pu.paidAmt) * 100) / 100,
        created_at: r[SP.CREATED] instanceof Date ? r[SP.CREATED].toISOString() : r[SP.CREATED]
      };
    });
    return { success: true, data: suppliers.reverse() };
  } catch (e) { return { success: false, message: 'Failed to load suppliers' }; }
}

function addSupplier(spData, userId, role) {
  try {
    if (role !== 'admin' && role !== 'manager') return { success: false, message: 'Access denied' };
    const { name, phone, address } = spData;
    const sh = getSheet(SHEETS.SUPPLIERS);
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const newId = getNextId(SHEETS.SUPPLIERS);
      sh.appendRow([newId, name.trim(), phone || '', address || '', 1, userId, ts()]);
      logActivity(userId, 'CREATE', 'Suppliers', newId, 'Added: ' + name.trim());
      return { success: true, message: 'Supplier added', data: { id: newId } };
    } finally { lock.releaseLock(); }
  } catch (e) { return { success: false, message: 'Failed' }; }
}

function updateSupplier(spData, userId, role) {
  try {
    if (role !== 'admin' && role !== 'manager') return { success: false, message: 'Access denied' };
    const { id, name, phone, address, is_active } = spData;
    const r = findRowByValue(SHEETS.SUPPLIERS, SP.ID, parseInt(id));
    if (!r) return { success: false, message: 'Supplier not found' };
    const row = r.data.slice();
    row[SP.NAME] = name.trim(); row[SP.PHONE] = phone || ''; row[SP.ADDR] = address || '';
    row[SP.ACTIVE] = (is_active == 1 || is_active === true) ? 1 : 0;
    getSheet(SHEETS.SUPPLIERS).getRange(r.row, 1, 1, row.length).setValues([row]);
    return { success: true, message: 'Supplier updated' };
  } catch (e) { return { success: false, message: 'Failed' }; }
}

function deleteSupplier(id, userId, role) {
  try {
    if (role !== 'admin') return { success: false, message: 'Only admin can delete' };
    const r = findRowByValue(SHEETS.SUPPLIERS, SP.ID, parseInt(id));
    if (!r) return { success: false, message: 'Supplier not found' };
    if (getPurchaseCountForSupplier(parseInt(id)) > 0) return { success: false, message: 'Cannot delete — has linked purchases' };
    getSheet(SHEETS.SUPPLIERS).deleteRow(r.row);
    return { success: true, message: 'Supplier deleted' };
  } catch (e) { return { success: false, message: 'Failed' }; }
}

function toggleSupplierStatus(id, userId, role) {
  try {
    if (role !== 'admin' && role !== 'manager') return { success: false, message: 'Access denied' };
    const r = findRowByValue(SHEETS.SUPPLIERS, SP.ID, parseInt(id));
    if (!r) return { success: false, message: 'Not found' };
    const row = r.data.slice();
    row[SP.ACTIVE] = isActive(row[SP.ACTIVE]) ? 0 : 1;
    getSheet(SHEETS.SUPPLIERS).getRange(r.row, 1, 1, row.length).setValues([row]);
    return { success: true, message: 'Status updated' };
  } catch (e) { return { success: false, message: 'Failed' }; }
}

function getSupplierLedger(supplierId, userId, role) {
  try {
    if (role === 'cashier') return { success: false, message: 'Access denied' };
    const sr = findRowByValue(SHEETS.SUPPLIERS, SP.ID, parseInt(supplierId));
    if (!sr) return { success: false, message: 'Supplier not found' };
    const supplier = { id: sr.data[SP.ID], name: sr.data[SP.NAME], phone: sr.data[SP.PHONE] || '', address: sr.data[SP.ADDR] || '' };
    let ledger = [];
    try {
      const data = getSheetData(SHEETS.PURCHASES);
      let balance = 0;
      ledger = data.filter(r => r[PU.SUPPLIER_ID] == parseInt(supplierId)).map(r => {
        const total = parseFloat(r[PU.TOTAL]) || 0;
        const paid = parseFloat(r[PU.PAID]) || 0;
        balance += (total - paid);
        return {
          id: r[PU.ID], purchase_no: r[PU.NO], date: r[PU.DATE] instanceof Date ? r[PU.DATE].toISOString().split('T')[0] : (r[PU.DATE] || ''),
          total_amount: total, paid_amount: paid, due: Math.round((total - paid) * 100) / 100,
          balance: Math.round(balance * 100) / 100, status: r[PU.STATUS] || 'pending'
        };
      });
    } catch (e) { }
    return { success: true, data: { supplier, ledger } };
  } catch (e) { return { success: false, message: 'Failed to load ledger' }; }
}

function getSuppliersForDropdown() {
  try {
    const data = getSheetData(SHEETS.SUPPLIERS);
    return { success: true, data: data.filter(r => isActive(r[SP.ACTIVE])).map(r => ({ id: r[SP.ID], name: r[SP.NAME] })) };
  } catch (e) { return { success: true, data: [] }; }
}

// Purchases
function genPurchaseNo() {
  const d = new Date();
  const dt = d.getFullYear() + String(d.getMonth() + 1).padStart(2, '0') + String(d.getDate()).padStart(2, '0');
  const data = getSheetData(SHEETS.PURCHASES);
  const today = data.filter(r => String(r[PU.NO]).includes(dt));
  const maxSeq = today.reduce((m, r) => { const parts = String(r[PU.NO]).split('-'); return Math.max(m, parseInt(parts[parts.length - 1]) || 0); }, 0);
  return 'PUR-' + dt + '-' + String(maxSeq + 1).padStart(3, '0');
}

function getPurchases(userId, role) {
  try {
    if (role === 'cashier') return { success: false, message: 'Access denied' };
    const data = getSheetData(SHEETS.PURCHASES);
    const spData = getSheetData(SHEETS.SUPPLIERS);
    const catData = getSheetData(SHEETS.CATEGORIES);
    const spMap = {}; spData.forEach(r => { spMap[r[SP.ID]] = r[SP.NAME]; });
    const catMap = {}; catData.forEach(r => { catMap[r[C.ID]] = r[C.NAME]; });
    const purchases = data.map(r => ({
      id: r[PU.ID], purchase_no: r[PU.NO], supplier_id: r[PU.SUPPLIER_ID], supplier_name: spMap[r[PU.SUPPLIER_ID]] || 'Unknown',
      category_id: r[PU.CAT_ID], category_name: catMap[r[PU.CAT_ID]] || 'Unknown',
      total_qty: parseInt(r[PU.QTY]) || 0, avg_rate: parseFloat(r[PU.RATE]) || 0,
      total_amount: parseFloat(r[PU.TOTAL]) || 0, paid_amount: parseFloat(r[PU.PAID]) || 0, due_amount: parseFloat(r[PU.DUE]) || 0,
      purchase_date: r[PU.DATE] instanceof Date ? r[PU.DATE].toISOString().split('T')[0] : (r[PU.DATE] || ''),
      notes: r[PU.NOTES] || '', status: r[PU.STATUS] || 'pending'
    }));
    return { success: true, data: purchases.reverse() };
  } catch (e) { return { success: false, message: 'Failed to load purchases' }; }
}

function addPurchase(puData, userId, role) {
  try {
    if (role === 'cashier') return { success: false, message: 'Access denied' };
    const { supplier_id, category_id, avg_rate, paid_amount, purchase_date, notes } = puData;
    if (!category_id || !purchase_date) return { success: false, message: 'Category and date required' };

    let totalQty = 0;
    try {
      const wsData = getSheetData(SHEETS.ITEM_STOCKS);
      const linked = wsData.filter(r => r[WS.CAT_ID] == parseInt(category_id));
      totalQty = linked.reduce((s, r) => s + (parseInt(r[WS.QTY]) || 1), 0);
    } catch (e) { }

    const rate = parseFloat(avg_rate) || 0;
    const totalAmt = Math.round(totalQty * rate * 100) / 100;
    const paid = Math.min(parseFloat(paid_amount) || 0, totalAmt);
    const due = Math.round((totalAmt - paid) * 100) / 100;

    const sh = getSheet(SHEETS.PURCHASES);
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const newId = getNextId(SHEETS.PURCHASES);
      const purNo = genPurchaseNo();
      const now = ts();
      sh.appendRow([newId, purNo, parseInt(supplier_id) || '', parseInt(category_id), totalQty, rate, totalAmt, paid, due, purchase_date, notes || '', 'pending', userId, now, now]);
      logActivity(userId, 'CREATE', 'Purchases', newId, 'Added: ' + purNo);
      return { success: true, message: 'Purchase created', data: { id: newId, purchase_no: purNo } };
    } finally { lock.releaseLock(); }
  } catch (e) { return { success: false, message: 'Failed to create purchase' }; }
}

function updatePurchase(puData, userId, role) {
  try {
    if (role === 'cashier') return { success: false, message: 'Access denied' };
    const { id, supplier_id, category_id, avg_rate, paid_amount, purchase_date, notes, status } = puData;
    const r = findRowByValue(SHEETS.PURCHASES, PU.ID, parseInt(id));
    if (!r) return { success: false, message: 'Not found' };

    const row = r.data.slice();
    row[PU.SUPPLIER_ID] = parseInt(supplier_id) || row[PU.SUPPLIER_ID];
    row[PU.CAT_ID] = parseInt(category_id) || row[PU.CAT_ID];
    row[PU.DATE] = purchase_date || row[PU.DATE];
    row[PU.NOTES] = notes !== undefined ? notes : row[PU.NOTES];
    row[PU.STATUS] = status || row[PU.STATUS];

    const rate = parseFloat(avg_rate) || parseFloat(row[PU.RATE]) || 0;
    row[PU.RATE] = rate;
    const qty = parseInt(row[PU.QTY]) || 0;
    row[PU.TOTAL] = Math.round(qty * rate * 100) / 100;
    row[PU.PAID] = Math.min(parseFloat(paid_amount) || parseFloat(row[PU.PAID]) || 0, row[PU.TOTAL]);
    row[PU.DUE] = Math.round((row[PU.TOTAL] - row[PU.PAID]) * 100) / 100;
    row[PU.UPDATED] = ts();

    getSheet(SHEETS.PURCHASES).getRange(r.row, 1, 1, row.length).setValues([row]);
    return { success: true, message: 'Purchase updated' };
  } catch (e) { return { success: false, message: 'Failed' }; }
}

function deletePurchase(id, userId, role) {
  try {
    if (role !== 'admin') return { success: false, message: 'Only admin can delete' };
    const r = findRowByValue(SHEETS.PURCHASES, PU.ID, parseInt(id));
    if (!r) return { success: false, message: 'Not found' };
    getSheet(SHEETS.PURCHASES).deleteRow(r.row);
    return { success: true, message: 'Deleted' };
  } catch (e) { return { success: false, message: 'Failed' }; }
}

function getPurchaseDetail(id, userId, role) {
  try {
    if (role === 'cashier') return { success: false, message: 'Access denied' };
    const r = findRowByValue(SHEETS.PURCHASES, PU.ID, parseInt(id));
    if (!r) return { success: false, message: 'Not found' };
    const p = r.data;
    const spName = p[PU.SUPPLIER_ID] ? (findRowByValue(SHEETS.SUPPLIERS, SP.ID, parseInt(p[PU.SUPPLIER_ID]))?.data[SP.NAME] || 'Unknown') : '—';
    const catName = findRowByValue(SHEETS.CATEGORIES, C.ID, parseInt(p[PU.CAT_ID]))?.data[C.NAME] || 'Unknown';

    let stocks = [];
    try {
      const wsData = getSheetData(SHEETS.ITEM_STOCKS);
      stocks = wsData.filter(w => w[WS.PUR_ID] == parseInt(id)).map(w => ({
        id: w[WS.ID], item_code: w[WS.ITEM_CODE] || '', item_name: w[WS.ITEM_NAME] || '', qty: parseInt(w[WS.QTY]) || 1,
        status: w[WS.STATUS] || '', buy_price: parseFloat(w[WS.BUY_PRICE]) || 0, sell_price: parseFloat(w[WS.SELL_PRICE]) || 0
      }));
    } catch (e) { }

    let payments = [];
    try {
      const pyData = getSheetData(SHEETS.PAYMENTS);
      payments = pyData.filter(py => py[PAY.PUR_ID] == parseInt(id) && py[PAY.TYPE] === 'supplier_payment').map(py => ({
        id: py[PAY.ID], amount: parseFloat(py[PAY.AMT]) || 0, method: py[PAY.METHOD] || '', reference: py[PAY.REF] || '',
        date: py[PAY.DATE] instanceof Date ? py[PAY.DATE].toISOString().split('T')[0] : (py[PAY.DATE] || ''), notes: py[PAY.NOTES] || ''
      }));
    } catch (e) { }

    return {
      success: true,
      data: {
        id: p[PU.ID], purchase_no: p[PU.NO], supplier_name: spName, category_name: catName,
        supplier_id: p[PU.SUPPLIER_ID], category_id: p[PU.CAT_ID],
        total_qty: parseInt(p[PU.QTY]) || 0, avg_rate: parseFloat(p[PU.RATE]) || 0,
        total_amount: parseFloat(p[PU.TOTAL]) || 0, paid_amount: parseFloat(p[PU.PAID]) || 0, due_amount: parseFloat(p[PU.DUE]) || 0,
        purchase_date: p[PU.DATE] instanceof Date ? p[PU.DATE].toISOString().split('T')[0] : (p[PU.DATE] || ''),
        notes: p[PU.NOTES] || '', status: p[PU.STATUS] || 'pending', stocks, payments
      }
    };
  } catch (e) { return { success: false, message: 'Failed' }; }
}

function recalcPurchase(purchaseId) {
  try {
    const r = findRowByValue(SHEETS.PURCHASES, PU.ID, parseInt(purchaseId));
    if (!r) return;
    const wsData = getSheetData(SHEETS.ITEM_STOCKS);
    const linked = wsData.filter(w => w[WS.PUR_ID] == parseInt(purchaseId));
    const row = r.data.slice();

    const getOrigQty = (w) => {
      if (w[WS.STATUS] === 'sold') return parseInt(w[WS.QTY]) || 1;
      const bRate = parseFloat(w[WS.BUY_RATE]) || 0;
      const bPrice = parseFloat(w[WS.BUY_PRICE]) || 0;
      if (bRate) return Math.round(bPrice / bRate) || (parseInt(w[WS.QTY]) || 1);
      return parseInt(w[WS.QTY]) || 1;
    };

    row[PU.QTY] = linked.reduce((s, w) => s + getOrigQty(w), 0);
    const newTotal = Math.round(linked.reduce((s, w) => s + ((parseFloat(w[WS.BUY_RATE]) || 0) * getOrigQty(w)), 0) * 100) / 100;
    row[PU.TOTAL] = newTotal;
    row[PU.RATE] = row[PU.QTY] > 0 ? Math.round((newTotal / row[PU.QTY]) * 100) / 100 : 0;
    row[PU.DUE] = Math.round((row[PU.TOTAL] - (parseFloat(row[PU.PAID]) || 0)) * 100) / 100;
    row[PU.UPDATED] = ts();
    getSheet(SHEETS.PURCHASES).getRange(r.row, 1, 1, row.length).setValues([row]);
  } catch (e) { }
}

function getPurchasesForDropdown(catId) {
  try {
    const data = getSheetData(SHEETS.PURCHASES);
    const filtered = catId ? data.filter(r => r[PU.CAT_ID] == parseInt(catId)) : data;
    return { success: true, data: filtered.map(r => ({ id: r[PU.ID], name: r[PU.NO] })) };
  } catch (e) { return { success: true, data: [] }; }
}

// Items Stock (Core Logic for Category-based deduplication and stock increment)
function getItemStocks(userId, role) {
  try {
    const data = getSheetData(SHEETS.ITEM_STOCKS);
    const catData = getSheetData(SHEETS.CATEGORIES);
    const catMap = {};
    catData.forEach(r => { catMap[r[C.ID]] = r[C.NAME]; });

    const stocks = data.map(r => {
      const qty = parseInt(r[WS.QTY]) || 1;
      return {
        id: r[WS.ID],
        item_code: r[WS.ITEM_CODE] || '',
        item_name: r[WS.ITEM_NAME] || '',
        category_name: catMap[r[WS.CAT_ID]] || 'General',
        category_id: r[WS.CAT_ID],
        qty: qty,
        buy_rate: parseFloat(r[WS.BUY_RATE]) || 0,
        buy_price: parseFloat(r[WS.BUY_PRICE]) || 0,
        sell_rate: parseFloat(r[WS.SELL_RATE]) || 0,
        sell_price: parseFloat(r[WS.SELL_PRICE]) || 0,
        status: r[WS.STATUS] || 'available',
        image: r[WS.IMG] || ''
      };
    });
    return { success: true, data: stocks.reverse() };
  } catch (e) { return { success: false, message: 'Failed to load item stocks' }; }
}

function addItemStock(wsData, userId, role) {
  try {
    if (role !== 'admin' && role !== 'manager' && role !== 'warehouse_staff') return { success: false, message: 'Access denied' };
    const { item_code, item_name, category_id, purchase_id, buy_rate, sell_rate, notes, qty, imageData } = wsData;
    if (!item_name || !category_id) return { success: false, message: 'Item name and category required' };

    const q = parseInt(qty) || 1;
    const bRate = parseFloat(buy_rate) || 0;
    const sRate = parseFloat(sell_rate) || 0;

    let imgId = '';
    if (imageData && imageData.data) imgId = uploadToDrive(imageData.data, imageData.name, imageData.type, 'ItemPOS_Images');

    const sh = getSheet(SHEETS.ITEM_STOCKS);
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      const data = getSheetData(SHEETS.ITEM_STOCKS);
      let existingRowIndex = -1;
      let existingData = null;

      // Stock Update Logic (No Duplicates): Check if item with same name exists under the same category
      for (let i = 0; i < data.length; i++) {
        if (String(data[i][WS.ITEM_NAME]).trim().toLowerCase() === String(item_name).trim().toLowerCase() &&
          parseInt(data[i][WS.CAT_ID]) === parseInt(category_id)) {
          existingRowIndex = i + 2;
          existingData = data[i];
          break;
        }
      }

      const now = ts();

      if (existingRowIndex > -1) {
        // Update/Increment existing item
        const row = existingData.slice();
        const currentQty = parseInt(row[WS.QTY]) || 0;
        const newQty = currentQty + q;
        row[WS.QTY] = newQty;
        // Optionally update rates to the latest
        row[WS.BUY_RATE] = bRate > 0 ? bRate : row[WS.BUY_RATE];
        row[WS.SELL_RATE] = sRate > 0 ? sRate : row[WS.SELL_RATE];
        row[WS.BUY_PRICE] = Math.round(newQty * parseFloat(row[WS.BUY_RATE]) * 100) / 100;
        row[WS.SELL_PRICE] = Math.round(newQty * parseFloat(row[WS.SELL_RATE]) * 100) / 100;

        // Always reset status to available if we are adding new stock
        if (row[WS.STATUS] === 'sold' || row[WS.STATUS] === 'damaged') {
          row[WS.STATUS] = 'available';
        }

        row[WS.UPDATED] = now;
        if (item_code) row[WS.ITEM_CODE] = item_code.trim(); // Update code if provided
        if (purchase_id) row[WS.PUR_ID] = parseInt(purchase_id);
        if (notes) row[WS.NOTES] = (row[WS.NOTES] || '') + ' | ' + notes;

        if (imgId) {
          if (row[WS.IMG]) deleteFromDrive(row[WS.IMG]);
          row[WS.IMG] = imgId;
        }

        sh.getRange(existingRowIndex, 1, 1, row.length).setValues([row]);
        if (purchase_id) recalcPurchase(parseInt(purchase_id));
        logActivity(userId, 'UPDATE', 'Item_Stocks', row[WS.ID], `Incremented stock for ${item_name} by ${q}`);
        return { success: true, message: `Stock incremented for existing item: ${item_name}`, data: { id: row[WS.ID] } };
      } else {
        // Create new item entry
        const newId = getNextId(SHEETS.ITEM_STOCKS);
        const bPrice = Math.round(q * bRate * 100) / 100;
        const sPrice = Math.round(q * sRate * 100) / 100;
        sh.appendRow([newId, (item_code || '').trim(), item_name.trim(), parseInt(category_id), purchase_id ? parseInt(purchase_id) : '', bRate, bPrice, sRate, sPrice, 'available', notes || '', userId, now, now, q, imgId]);
        if (purchase_id) recalcPurchase(parseInt(purchase_id));
        logActivity(userId, 'CREATE', 'Item_Stocks', newId, `Added new item: ${item_name}`);
        return { success: true, message: 'Item added successfully', data: { id: newId } };
      }
    } finally { lock.releaseLock(); }
  } catch (e) { return { success: false, message: 'Failed to add item' }; }
}

function updateItemStock(wsData, userId, role) {
  try {
    if (role !== 'admin' && role !== 'manager') return { success: false, message: 'Access denied' };
    const { id, item_code, item_name, category_id, purchase_id, buy_rate, sell_rate, status, notes, qty, imageData, removeImage } = wsData;

    const r = findRowByValue(SHEETS.ITEM_STOCKS, WS.ID, parseInt(id));
    if (!r) return { success: false, message: 'Item not found' };

    const row = r.data.slice();
    if (item_code !== undefined) row[WS.ITEM_CODE] = item_code.trim();
    if (item_name) row[WS.ITEM_NAME] = item_name.trim();
    if (category_id) row[WS.CAT_ID] = parseInt(category_id);
    if (purchase_id !== undefined) row[WS.PUR_ID] = purchase_id ? parseInt(purchase_id) : '';

    const q = qty !== undefined ? (parseInt(qty) || 0) : (parseInt(row[WS.QTY]) || 0);
    row[WS.QTY] = q;
    row[WS.BUY_RATE] = parseFloat(buy_rate) || parseFloat(row[WS.BUY_RATE]) || 0;
    row[WS.BUY_PRICE] = Math.round(q * row[WS.BUY_RATE] * 100) / 100;
    row[WS.SELL_RATE] = parseFloat(sell_rate) || parseFloat(row[WS.SELL_RATE]) || 0;
    row[WS.SELL_PRICE] = Math.round(q * row[WS.SELL_RATE] * 100) / 100;
    if (status) row[WS.STATUS] = status;
    if (notes !== undefined) row[WS.NOTES] = notes;

    if (imageData && imageData.data) {
      if (row[WS.IMG]) deleteFromDrive(row[WS.IMG]);
      row[WS.IMG] = uploadToDrive(imageData.data, imageData.name, imageData.type, 'ItemPOS_Images');
    } else if (removeImage) {
      if (row[WS.IMG]) deleteFromDrive(row[WS.IMG]);
      row[WS.IMG] = '';
    }

    row[WS.UPDATED] = ts();
    getSheet(SHEETS.ITEM_STOCKS).getRange(r.row, 1, 1, row.length).setValues([row]);
    return { success: true, message: 'Item updated' };
  } catch (e) { return { success: false, message: 'Failed to update' }; }
}

function updateItemStatus(id, status, userId, role) {
  try {
    const r = findRowByValue(SHEETS.ITEM_STOCKS, WS.ID, parseInt(id));
    if (!r) return { success: false, message: 'Not found' };
    const row = r.data.slice();
    row[WS.STATUS] = status; row[WS.UPDATED] = ts();
    getSheet(SHEETS.ITEM_STOCKS).getRange(r.row, 1, 1, row.length).setValues([row]);
    return { success: true, message: 'Status updated' };
  } catch (e) { return { success: false, message: 'Failed' }; }
}

function deleteItemStock(id, userId, role) {
  try {
    const r = findRowByValue(SHEETS.ITEM_STOCKS, WS.ID, parseInt(id));
    if (!r) return { success: false, message: 'Not found' };
    if (r.data[WS.STATUS] === 'sold') return { success: false, message: 'Cannot delete sold item' };
    const purId = r.data[WS.PUR_ID];
    if (r.data[WS.IMG]) deleteFromDrive(r.data[WS.IMG]);
    getSheet(SHEETS.ITEM_STOCKS).deleteRow(r.row);
    if (purId) recalcPurchase(parseInt(purId));
    return { success: true, message: 'Item deleted' };
  } catch (e) { return { success: false, message: 'Failed' }; }
}

function bulkImportItem(items, catId, purchaseId, buyRate, sellRate, userId, role) {
  try {
    const sh = getSheet(SHEETS.ITEM_STOCKS);
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);

    let success = 0, failed = 0, errors = [];
    try {
      const existing = getSheetData(SHEETS.ITEM_STOCKS);
      // Map existing items in this category by normalized name
      const catExistingMap = {};
      existing.forEach((r, idx) => {
        if (parseInt(r[WS.CAT_ID]) === parseInt(catId)) {
          const nameKey = String(r[WS.ITEM_NAME]).trim().toLowerCase();
          if (nameKey) catExistingMap[nameKey] = { rowIdx: idx + 2, data: r };
        }
      });

      let nextId = getNextId(SHEETS.ITEM_STOCKS);
      const now = ts();
      const newRows = [];
      const updateOperations = []; // To collect rows that just need a quantity bump

      items.forEach((item, idx) => {
        const item_name = String(item.item_name || '').trim();
        const item_code = String(item.item_code || '').trim();
        const q = parseInt(item.qty) || 1;
        const bRate = parseFloat(item.buy_rate) || parseFloat(buyRate) || 0;
        const sRate = parseFloat(item.sell_rate) || parseFloat(sellRate) || 0;

        if (!item_name) { failed++; errors.push({ row: idx + 1, item_name, reason: 'Empty Item Name' }); return; }

        const nameKey = item_name.toLowerCase();

        if (catExistingMap[nameKey]) {
          // Increment existing
          const existingItem = catExistingMap[nameKey];
          const row = existingItem.data.slice();
          const currentQty = parseInt(row[WS.QTY]) || 0;
          const newQty = currentQty + q;

          row[WS.QTY] = newQty;
          row[WS.BUY_RATE] = bRate > 0 ? bRate : row[WS.BUY_RATE];
          row[WS.SELL_RATE] = sRate > 0 ? sRate : row[WS.SELL_RATE];
          row[WS.BUY_PRICE] = Math.round(newQty * parseFloat(row[WS.BUY_RATE]) * 100) / 100;
          row[WS.SELL_PRICE] = Math.round(newQty * parseFloat(row[WS.SELL_RATE]) * 100) / 100;
          if (row[WS.STATUS] === 'sold' || row[WS.STATUS] === 'damaged') row[WS.STATUS] = 'available';
          row[WS.UPDATED] = now;
          if (item_code) row[WS.ITEM_CODE] = item_code;
          if (purchaseId) row[WS.PUR_ID] = parseInt(purchaseId);

          updateOperations.push({ rowIdx: existingItem.rowIdx, data: row });
          // Update the map to reflect new state for subsequent rows in the same batch
          existingItem.data = row;
          success++;
        } else {
          // Add new
          newRows.push([nextId++, item_code, item_name, parseInt(catId), purchaseId ? parseInt(purchaseId) : '', bRate, Math.round(q * bRate * 100) / 100, sRate, Math.round(q * sRate * 100) / 100, 'available', '', userId, now, now, q, '']);
          // Add to map so duplicates in the same CSV batch increment correctly
          catExistingMap[nameKey] = { rowIdx: -1, data: newRows[newRows.length - 1], isNew: true };
          success++;
        }
      });

      // Write updates
      updateOperations.forEach(op => {
        sh.getRange(op.rowIdx, 1, 1, op.data.length).setValues([op.data]);
      });

      // Write new rows (filter out those that were merged into themselves in the map)
      const finalNewRows = Object.values(catExistingMap).filter(v => v.isNew).map(v => v.data);
      if (finalNewRows.length > 0) sh.getRange(sh.getLastRow() + 1, 1, finalNewRows.length, 16).setValues(finalNewRows);

      if (purchaseId) recalcPurchase(parseInt(purchaseId));

      try {
        const ilSh = getSheet(SHEETS.IMPORT_LOGS);
        ilSh.appendRow([getNextId(SHEETS.IMPORT_LOGS), 'bulk_import.csv', parseInt(catId), items.length, success, failed, failed > 0 ? 'completed' : 'completed', JSON.stringify(errors.slice(0, 50)), userId, ts()]);
      } catch (e) { }
    } finally { lock.releaseLock(); }

    return { success: true, message: success + ' imported (added/updated), ' + failed + ' failed', data: { success, failed, errors } };
  } catch (e) { return { success: false, message: 'Import failed' }; }
}

function getItemStockDetail(id) {
  try {
    const r = findRowByValue(SHEETS.ITEM_STOCKS, WS.ID, parseInt(id));
    if (!r) return { success: false, message: 'Not found' };
    const w = r.data;
    const catName = findRowByValue(SHEETS.CATEGORIES, C.ID, parseInt(w[WS.CAT_ID]))?.data[C.NAME] || 'Unknown';
    return {
      success: true,
      data: {
        id: w[WS.ID], item_code: w[WS.ITEM_CODE], item_name: w[WS.ITEM_NAME], category_name: catName, category_id: w[WS.CAT_ID],
        purchase_id: w[WS.PUR_ID] || '', qty: parseInt(w[WS.QTY]) || 1,
        buy_rate: parseFloat(w[WS.BUY_RATE]), buy_price: parseFloat(w[WS.BUY_PRICE]),
        sell_rate: parseFloat(w[WS.SELL_RATE]), sell_price: parseFloat(w[WS.SELL_PRICE]),
        status: w[WS.STATUS], notes: w[WS.NOTES] || '', image: w[WS.IMG] || '',
        created_at: w[WS.CREATED] instanceof Date ? w[WS.CREATED].toISOString() : w[WS.CREATED]
      }
    };
  } catch (e) { return { success: false, message: 'Failed' }; }
}

// Customers
function getCustomers(userId, role) {
  try {
    const data = getSheetData(SHEETS.CUSTOMERS);
    const uData = getSheetData(SHEETS.USERS);
    const uMap = {}; uData.forEach(r => { uMap[r[U.ID]] = r[U.NAME]; });
    const custs = data.map(r => ({
      id: r[CU.ID], name: r[CU.NAME], phone: r[CU.PHONE] || '', address: r[CU.ADDR] || '',
      total_purchase: parseFloat(r[CU.TOTAL]) || 0, total_paid: parseFloat(r[CU.PAID]) || 0, total_due: parseFloat(r[CU.DUE]) || 0,
      is_active: isActive(r[CU.ACTIVE]) ? 1 : 0, created_by_name: uMap[r[CU.CREATED_BY]] || 'Unknown'
    }));
    return { success: true, data: custs.reverse() };
  } catch (e) { return { success: false, message: 'Failed' }; }
}

function addCustomer(cuData, userId, role) {
  try {
    const sh = getSheet(SHEETS.CUSTOMERS);
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const newId = getNextId(SHEETS.CUSTOMERS);
      sh.appendRow([newId, cuData.name.trim(), cuData.phone || '', cuData.address || '', 0, 0, 0, 1, userId, ts()]);
      return { success: true, message: 'Added', data: { id: newId, name: cuData.name.trim() } };
    } finally { lock.releaseLock(); }
  } catch (e) { return { success: false, message: 'Failed' }; }
}

function updateCustomer(cuData, userId, role) {
  try {
    const r = findRowByValue(SHEETS.CUSTOMERS, CU.ID, parseInt(cuData.id));
    const row = r.data.slice();
    row[CU.NAME] = cuData.name.trim(); row[CU.PHONE] = cuData.phone || ''; row[CU.ADDR] = cuData.address || '';
    row[CU.ACTIVE] = isActive(cuData.is_active) ? 1 : 0;
    getSheet(SHEETS.CUSTOMERS).getRange(r.row, 1, 1, row.length).setValues([row]);
    return { success: true, message: 'Updated' };
  } catch (e) { return { success: false, message: 'Failed' }; }
}

function deleteCustomer(id, userId, role) {
  try {
    const r = findRowByValue(SHEETS.CUSTOMERS, CU.ID, parseInt(id));
    getSheet(SHEETS.CUSTOMERS).deleteRow(r.row);
    return { success: true, message: 'Deleted' };
  } catch (e) { return { success: false, message: 'Failed' }; }
}

function toggleCustomerStatus(id, userId, role) {
  try {
    const r = findRowByValue(SHEETS.CUSTOMERS, CU.ID, parseInt(id));
    const row = r.data.slice();
    row[CU.ACTIVE] = isActive(row[CU.ACTIVE]) ? 0 : 1;
    getSheet(SHEETS.CUSTOMERS).getRange(r.row, 1, 1, row.length).setValues([row]);
    return { success: true, message: 'Updated' };
  } catch (e) { return { success: false, message: 'Failed' }; }
}

function getCustomerLedger(custId, userId, role) {
  try {
    const cr = findRowByValue(SHEETS.CUSTOMERS, CU.ID, parseInt(custId));
    if (!cr) return { success: false, message: 'Not found' };
    const customer = { id: cr.data[CU.ID], name: cr.data[CU.NAME], phone: cr.data[CU.PHONE] || '', address: cr.data[CU.ADDR] || '', total_purchase: parseFloat(cr.data[CU.TOTAL]) || 0, total_paid: parseFloat(cr.data[CU.PAID]) || 0, total_due: parseFloat(cr.data[CU.DUE]) || 0 };

    let invoices = [];
    try {
      const slData = getSheetData(SHEETS.SALES);
      invoices = slData.filter(s => s[SL.CUST_ID] == parseInt(custId)).map(s => ({
        id: s[SL.ID], invoice_no: s[SL.INV_NO], date: s[SL.DATE] instanceof Date ? s[SL.DATE].toISOString().split('T')[0] : (s[SL.DATE] || ''),
        total: parseFloat(s[SL.TOTAL]) || 0, paid: parseFloat(s[SL.PAID]) || 0, due: parseFloat(s[SL.DUE]) || 0, status: s[SL.STATUS] || 'pending'
      })).reverse();
    } catch (e) { }

    let payments = [];
    try {
      const pyData = getSheetData(SHEETS.PAYMENTS);
      const uData = getSheetData(SHEETS.USERS);
      const uMap = {}; uData.forEach(u => { uMap[u[U.ID]] = u[U.NAME]; });
      const custSaleIds = new Set(invoices.map(i => i.id));
      payments = pyData.filter(p => p[PAY.TYPE] === 'customer_payment' && (custSaleIds.has(p[PAY.SALE_ID]) || p[PAY.PUR_ID] == parseInt(custId))).map(p => ({
        id: p[PAY.ID], sale_id: p[PAY.SALE_ID] || '', amount: parseFloat(p[PAY.AMT]) || 0, method: p[PAY.METHOD] || '', reference: p[PAY.REF] || '',
        date: p[PAY.DATE] instanceof Date ? p[PAY.DATE].toISOString().split('T')[0] : (p[PAY.DATE] || ''), notes: p[PAY.NOTES] || '', created_by_name: uMap[p[PAY.CREATED_BY]] || 'Unknown'
      })).reverse();
    } catch (e) { }

    return { success: true, data: { customer, invoices, payments } };
  } catch (e) { return { success: false, message: 'Failed' }; }
}

function getCustomersForDropdown() {
  try {
    const data = getSheetData(SHEETS.CUSTOMERS);
    return { success: true, data: data.filter(r => isActive(r[CU.ACTIVE])).map(r => ({ id: r[CU.ID], name: r[CU.NAME] })) };
  } catch (e) { return { success: true, data: [] }; }
}
// Payments
function addPayment(payData, userId, role) {
  try {
    if (role === 'cashier') return { success: false, message: 'Access denied' };
    const { purchase_id, amount, payment_method, reference_no, payment_date, notes } = payData;
    const pr = findRowByValue(SHEETS.PURCHASES, PU.ID, parseInt(purchase_id));
    if (!pr) return { success: false, message: 'Purchase not found' };

    const sh = getSheet(SHEETS.PAYMENTS);
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const newId = getNextId(SHEETS.PAYMENTS);
      sh.appendRow([newId, '', parseInt(purchase_id), 'supplier_payment', parseFloat(amount), payment_method || 'cash', reference_no || '', payment_date || ts().split('T')[0], notes || '', userId, ts()]);

      const pRow = pr.data.slice();
      pRow[PU.PAID] = (parseFloat(pRow[PU.PAID]) || 0) + parseFloat(amount);
      pRow[PU.DUE] = Math.round(((parseFloat(pRow[PU.TOTAL]) || 0) - pRow[PU.PAID]) * 100) / 100;
      if (pRow[PU.DUE] <= 0) { pRow[PU.DUE] = 0; pRow[PU.STATUS] = 'completed'; }
      pRow[PU.UPDATED] = ts();
      getSheet(SHEETS.PURCHASES).getRange(pr.row, 1, 1, pRow.length).setValues([pRow]);
      return { success: true, message: 'Payment recorded' };
    } finally { lock.releaseLock(); }
  } catch (e) { return { success: false, message: 'Failed' }; }
}

function addCustomerPayment(payData, userId, role) {
  try {
    const { customer_id, sale_id, amount, method, reference, payment_date, notes } = payData;
    const cr = findRowByValue(SHEETS.CUSTOMERS, CU.ID, parseInt(customer_id));
    const sh = getSheet(SHEETS.PAYMENTS);
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const newId = getNextId(SHEETS.PAYMENTS);
      sh.appendRow([newId, sale_id ? parseInt(sale_id) : '', parseInt(customer_id), 'customer_payment', parseFloat(amount), method || 'cash', reference || '', payment_date || ts().split('T')[0], notes || '', userId, ts()]);

      const crow = cr.data.slice();
      crow[CU.PAID] = (parseFloat(crow[CU.PAID]) || 0) + parseFloat(amount);
      crow[CU.DUE] = Math.round(((parseFloat(crow[CU.TOTAL]) || 0) - crow[CU.PAID]) * 100) / 100;
      if (crow[CU.DUE] < 0) crow[CU.DUE] = 0;
      getSheet(SHEETS.CUSTOMERS).getRange(cr.row, 1, 1, crow.length).setValues([crow]);

      if (sale_id) {
        const sr = findRowByValue(SHEETS.SALES, SL.ID, parseInt(sale_id));
        if (sr) {
          const srow = sr.data.slice();
          srow[SL.PAID] = (parseFloat(srow[SL.PAID]) || 0) + parseFloat(amount);
          srow[SL.DUE] = Math.round(((parseFloat(srow[SL.TOTAL]) || 0) - srow[SL.PAID]) * 100) / 100;
          if (srow[SL.DUE] <= 0) { srow[SL.DUE] = 0; srow[SL.STATUS] = 'completed'; }
          getSheet(SHEETS.SALES).getRange(sr.row, 1, 1, srow.length).setValues([srow]);
        }
      }
      return { success: true, message: 'Payment recorded' };
    } finally { lock.releaseLock(); }
  } catch (e) { return { success: false, message: 'Failed' }; }
}

function getPayments(userId, role) {
  try {
    const data = getSheetData(SHEETS.PAYMENTS);
    const cuData = getSheetData(SHEETS.CUSTOMERS);
    const cuMap = {}; cuData.forEach(r => { cuMap[r[CU.ID]] = r[CU.NAME]; });
    const spData = getSheetData(SHEETS.SUPPLIERS);
    const spMap = {}; spData.forEach(r => { spMap[r[SP.ID]] = r[SP.NAME]; });
    const uData = getSheetData(SHEETS.USERS);
    const uMap = {}; uData.forEach(r => { uMap[r[U.ID]] = r[U.NAME]; });

    const slData = getSheetData(SHEETS.SALES);
    const saleCustMap = {}; slData.forEach(r => { saleCustMap[r[SL.ID]] = { inv: r[SL.INV_NO], cust: r[SL.CUST_ID] ? (cuMap[r[SL.CUST_ID]] || 'Walk-in') : 'Walk-in' }; });
    const puData = getSheetData(SHEETS.PURCHASES);
    const purSuppMap = {}; puData.forEach(r => { purSuppMap[r[PU.ID]] = { no: r[PU.NO], supp: spMap[r[PU.SUPPLIER_ID]] || 'Unknown' }; });

    const payments = data.map(r => {
      const isCust = r[PAY.TYPE] === 'customer_payment';
      const saleInfo = r[PAY.SALE_ID] ? saleCustMap[r[PAY.SALE_ID]] : null;
      const purInfo = !isCust && r[PAY.PUR_ID] ? purSuppMap[r[PAY.PUR_ID]] : null;
      return {
        id: r[PAY.ID], payment_type: r[PAY.TYPE], ref_no: isCust ? (saleInfo?.inv || '') : (purInfo?.no || ''),
        party_name: isCust ? (saleInfo?.cust || cuMap[r[PAY.PUR_ID]] || 'Walk-in') : (purInfo?.supp || 'Unknown'),
        amount: parseFloat(r[PAY.AMT]) || 0, method: r[PAY.METHOD] || '', reference: r[PAY.REF] || '',
        date: r[PAY.DATE] instanceof Date ? r[PAY.DATE].toISOString().split('T')[0] : (r[PAY.DATE] || ''),
        notes: r[PAY.NOTES] || '', created_by_name: uMap[r[PAY.CREATED_BY]] || 'Unknown'
      };
    });
    return { success: true, data: payments.reverse() };
  } catch (e) { return { success: false, message: 'Failed' }; }
}

// Sales / POS
function genInvoiceNo() {
  const d = new Date();
  const dt = d.getFullYear() + String(d.getMonth()+1).padStart(2,'0') + String(d.getDate()).padStart(2,'0');
  const data = getSheetData(SHEETS.SALES);
  const today = data.filter(r => String(r[SL.INV_NO]).includes(dt));
  const maxSeq = today.reduce((m, r) => { const parts = String(r[SL.INV_NO]).split('-'); return Math.max(m, parseInt(parts[parts.length-1]) || 0); }, 0);
  return 'INV-' + dt + '-' + String(maxSeq + 1).padStart(3, '0');
}

function getAvailableItem(catId) {
  try {
    const data = getSheetData(SHEETS.ITEM_STOCKS);
    let results = data.filter(r => r[WS.STATUS] === 'available');
    if (catId) results = results.filter(r => r[WS.CAT_ID] == parseInt(catId));
    const catData = getSheetData(SHEETS.CATEGORIES);
    const catMap = {}; catData.forEach(r => { catMap[r[C.ID]] = r[C.NAME]; });
    return {
      success: true,
      data: results.map(r => ({
        id: r[WS.ID], item_code: r[WS.ITEM_CODE] || '', item_name: r[WS.ITEM_NAME] || '', category_name: catMap[r[WS.CAT_ID]] || '',
        category_id: r[WS.CAT_ID], qty: parseInt(r[WS.QTY]) || 1,
        sell_rate: parseFloat(r[WS.SELL_RATE]), sell_price: parseFloat(r[WS.SELL_PRICE]), image: r[WS.IMG] || ''
      }))
    };
  } catch(e) { return { success: true, data: [] }; }
}

function searchItemBySerial(query, catId) {
  try {
    const data = getSheetData(SHEETS.ITEM_STOCKS);
    const q = String(query).trim().toLowerCase();
    let results = data.filter(r => r[WS.STATUS] === 'available' && (String(r[WS.ITEM_CODE]).toLowerCase().includes(q) || String(r[WS.ITEM_NAME]).toLowerCase().includes(q)));
    if (catId) results = results.filter(r => r[WS.CAT_ID] == parseInt(catId));
    const catData = getSheetData(SHEETS.CATEGORIES);
    const catMap = {}; catData.forEach(r => { catMap[r[C.ID]] = r[C.NAME]; });
    return {
      success: true,
      data: results.slice(0, 20).map(r => ({
        id: r[WS.ID], item_code: r[WS.ITEM_CODE] || '', item_name: r[WS.ITEM_NAME] || '', category_name: catMap[r[WS.CAT_ID]] || '',
        category_id: r[WS.CAT_ID], qty: parseInt(r[WS.QTY]) || 1,
        sell_rate: parseFloat(r[WS.SELL_RATE]), sell_price: parseFloat(r[WS.SELL_PRICE]), image: r[WS.IMG] || ''
      }))
    };
  } catch(e) { return { success: true, data: [] }; }
}

function completeSale(saleData, userId, role) {
  try {
    const { customer_id, items, discount, paid_amount, payment_method, payment_reference, notes, status } = saleData;
    if (!items || !items.length) return { success: false, message: 'Add items to cart' };

    const totalItems = items.reduce((s, i) => s + (parseInt(i.qty) || 1), 0);
    const subtotal = items.reduce((s, i) => s + (parseFloat(i.line_total) || 0), 0);
    const disc = parseFloat(discount) || 0;
    const grandTotal = Math.round((subtotal - disc) * 100) / 100;
    const paid = Math.min(parseFloat(paid_amount) || 0, grandTotal);
    const due = Math.round((grandTotal - paid) * 100) / 100;
    const saleStatus = status || (due <= 0 ? 'completed' : 'pending');

    const sh = getSheet(SHEETS.SALES);
    const siSh = getSheet(SHEETS.SALE_ITEMS);
    const wsSh = getSheet(SHEETS.ITEM_STOCKS);
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);

    try {
      const saleId = getNextId(SHEETS.SALES);
      const invNo = genInvoiceNo();
      const now = ts();

      sh.appendRow([saleId, invNo, customer_id ? parseInt(customer_id) : '', totalItems, Math.round(subtotal*100)/100, disc, grandTotal, paid, due, payment_method || 'cash', now, saleStatus, notes || '', userId, now, now]);

      let siId = getNextId(SHEETS.SALE_ITEMS);
      const siRows = [];
      const wsUpdates = [];

      items.forEach(item => {
        const soldQty = parseInt(item.qty) || 1;
        // SALE_ITEMS cols: 0=id, 1=sale_id, 2=item_stock_id, 3=item_code, 4=item_name, 5=qty, 6=rate, 7=line_total
        siRows.push([siId++, saleId, parseInt(item.item_stock_id), item.item_code || '', item.item_name || '', soldQty, parseFloat(item.rate), parseFloat(item.line_total)]);

        const wr = findRowByValue(SHEETS.ITEM_STOCKS, WS.ID, parseInt(item.item_stock_id));
        if (wr) {
          const wrow = wr.data.slice();
          const stockQty = parseInt(wrow[WS.QTY]) || 1;
          if (soldQty >= stockQty) {
            wrow[WS.QTY] = 0;
            wrow[WS.STATUS] = 'sold';
            wrow[WS.BUY_PRICE] = 0;
            wrow[WS.SELL_PRICE] = 0;
          } else {
            const remain = stockQty - soldQty;
            wrow[WS.QTY] = remain;
            wrow[WS.BUY_PRICE] = Math.round(remain * (parseFloat(wrow[WS.BUY_RATE]) || 0) * 100) / 100;
            wrow[WS.SELL_PRICE] = Math.round(remain * (parseFloat(wrow[WS.SELL_RATE]) || 0) * 100) / 100;
          }
          wrow[WS.UPDATED] = now;
          wsUpdates.push({ row: wr.row, data: wrow });
        }
      });

      if (siRows.length) siSh.getRange(siSh.getLastRow() + 1, 1, siRows.length, 8).setValues(siRows);
      wsUpdates.forEach(u => wsSh.getRange(u.row, 1, 1, u.data.length).setValues([u.data]));

      if (customer_id) {
        const cr = findRowByValue(SHEETS.CUSTOMERS, CU.ID, parseInt(customer_id));
        if (cr) {
          const crow = cr.data.slice();
          crow[CU.TOTAL] = (parseFloat(crow[CU.TOTAL]) || 0) + grandTotal;
          crow[CU.PAID] = (parseFloat(crow[CU.PAID]) || 0) + paid;
          crow[CU.DUE] = Math.round(((parseFloat(crow[CU.TOTAL])) - (parseFloat(crow[CU.PAID]))) * 100) / 100;
          getSheet(SHEETS.CUSTOMERS).getRange(cr.row, 1, 1, crow.length).setValues([crow]);
        }
      }

      if (paid > 0) {
        const pyId = getNextId(SHEETS.PAYMENTS);
        getSheet(SHEETS.PAYMENTS).appendRow([pyId, saleId, customer_id ? parseInt(customer_id) : '', 'customer_payment', paid, payment_method || 'cash', payment_reference || '', now.split('T')[0], '', userId, now]);
      }

      logActivity(userId, 'CREATE', 'Sales', saleId, invNo + ' ৳' + grandTotal);
      return { success: true, message: 'Sale completed', data: { id: saleId, invoice_no: invNo } };
    } finally { lock.releaseLock(); }
  } catch (e) { return { success: false, message: 'Failed' }; }
}

function getSales(userId, role) {
  try {
    const data = getSheetData(SHEETS.SALES);
    const cuData = getSheetData(SHEETS.CUSTOMERS);
    const uData = getSheetData(SHEETS.USERS);
    const cuMap = {}; cuData.forEach(r => { cuMap[r[CU.ID]] = r[CU.NAME]; });
    const uMap = {}; uData.forEach(r => { uMap[r[U.ID]] = r[U.NAME]; });

    let sales = data.map(r => ({
      id: r[SL.ID], invoice_no: r[SL.INV_NO], customer_id: r[SL.CUST_ID], customer_name: r[SL.CUST_ID] ? (cuMap[r[SL.CUST_ID]] || 'Unknown') : 'Walk-in',
      total_items: parseInt(r[SL.ITEMS]) || 0, grand_total: parseFloat(r[SL.TOTAL]) || 0, paid_amount: parseFloat(r[SL.PAID]) || 0,
      due_amount: parseFloat(r[SL.DUE]) || 0, payment_method: r[SL.METHOD] || 'cash',
      sale_date: r[SL.DATE] instanceof Date ? r[SL.DATE].toISOString() : r[SL.DATE],
      status: r[SL.STATUS] || 'completed', created_by: r[SL.CREATED_BY], cashier_name: uMap[r[SL.CREATED_BY]] || 'Unknown'
    }));

    if (role === 'cashier') sales = sales.filter(s => s.created_by == userId);
    return { success: true, data: sales.reverse() };
  } catch (e) { return { success: false, message: 'Failed' }; }
}

function getSaleDetail(id, userId, role) {
  try {
    const r = findRowByValue(SHEETS.SALES, SL.ID, parseInt(id));
    if (!r) return { success: false, message: 'Not found' };
    const s = r.data;
    const custName = s[SL.CUST_ID] ? (findRowByValue(SHEETS.CUSTOMERS, CU.ID, parseInt(s[SL.CUST_ID]))?.data[CU.NAME] || 'Unknown') : 'Walk-in';

    const siData = getSheetData(SHEETS.SALE_ITEMS);
    // Cols: 0=id, 1=sale_id, 2=item_stock_id, 3=item_code, 4=item_name, 5=qty, 6=rate, 7=line_total
    const items = siData.filter(si => si[SI.SALE_ID] == parseInt(id)).map(si => ({
      id: si[SI.ID], item_code: si[SI.ITEM_CODE], item_name: si[SI.ITEM_NAME] || '', qty: parseInt(si[SI.QTY]), rate: parseFloat(si[SI.RATE]), total: parseFloat(si[SI.TOTAL])
    }));

    const uData = getSheetData(SHEETS.USERS);
    const uMap = {}; uData.forEach(u => { uMap[u[U.ID]] = u[U.NAME]; });
    let payments = [];
    try {
      const pyData = getSheetData(SHEETS.PAYMENTS);
      payments = pyData.filter(p => p[PAY.SALE_ID] == parseInt(id) && p[PAY.TYPE] === 'customer_payment').map(p => ({
        id: p[PAY.ID], amount: parseFloat(p[PAY.AMT]), method: p[PAY.METHOD] || '', reference: p[PAY.REF] || '',
        date: p[PAY.DATE] instanceof Date ? p[PAY.DATE].toISOString().split('T')[0] : (p[PAY.DATE] || ''), created_by_name: uMap[p[PAY.CREATED_BY]] || 'Unknown'
      })).reverse();
    } catch(e) {}

    return {
      success: true,
      data: {
        id: s[SL.ID], invoice_no: s[SL.INV_NO], customer_id: s[SL.CUST_ID], customer_name: custName,
        total_items: parseInt(s[SL.ITEMS]), subtotal: parseFloat(s[SL.SUBTOTAL]) || 0, discount: parseFloat(s[SL.DISC]) || 0,
        grand_total: parseFloat(s[SL.TOTAL]), paid_amount: parseFloat(s[SL.PAID]), due_amount: parseFloat(s[SL.DUE]),
        payment_method: s[SL.METHOD], sale_date: s[SL.DATE] instanceof Date ? s[SL.DATE].toISOString() : s[SL.DATE],
        status: s[SL.STATUS], notes: s[SL.NOTES] || '', cashier_name: uMap[s[SL.CREATED_BY]] || 'Unknown', items, payments
      }
    };
  } catch (e) { return { success: false, message: 'Failed' }; }
}

function cancelSale(id, reason, userId, role) {
  try {
    const r = findRowByValue(SHEETS.SALES, SL.ID, parseInt(id));
    if (!r) return { success: false, message: 'Not found' };
    if (r.data[SL.STATUS] === 'cancelled') return { success: false, message: 'Already cancelled' };

    const row = r.data.slice();
    row[SL.STATUS] = 'cancelled'; row[SL.NOTES] = (row[SL.NOTES] || '') + ' | Cancelled: ' + (reason || 'No reason'); row[SL.UPDATED] = ts();
    getSheet(SHEETS.SALES).getRange(r.row, 1, 1, row.length).setValues([row]);

    const siData = getSheetData(SHEETS.SALE_ITEMS);
    const saleItems = siData.filter(si => si[SI.SALE_ID] == parseInt(id));
    const wsData = getSheetData(SHEETS.ITEM_STOCKS);
    const wsSh = getSheet(SHEETS.ITEM_STOCKS);
    const soldMap = {};
    saleItems.forEach(si => { soldMap[parseInt(si[SI.WS_ID])] = parseInt(si[SI.QTY]) || 1; });
    
    const nowTs = ts();
    wsData.forEach((w, idx) => {
      const wsId = w[WS.ID];
      if (soldMap[wsId] !== undefined) {
        const wrow = w.slice();
        const addBack = soldMap[wsId] || 1;
        const newQty = (parseInt(wrow[WS.QTY]) || 0) + addBack;
        wrow[WS.QTY] = newQty;
        wrow[WS.STATUS] = 'available'; // Set available in case it was sold out
        wrow[WS.BUY_PRICE] = Math.round(newQty * (parseFloat(wrow[WS.BUY_RATE]) || 0) * 100) / 100;
        wrow[WS.SELL_PRICE] = Math.round(newQty * (parseFloat(wrow[WS.SELL_RATE]) || 0) * 100) / 100;
        wrow[WS.UPDATED] = nowTs;
        wsSh.getRange(idx + 2, 1, 1, wrow.length).setValues([wrow]);
      }
    });

    if (row[SL.CUST_ID]) {
      const cr = findRowByValue(SHEETS.CUSTOMERS, CU.ID, parseInt(row[SL.CUST_ID]));
      if (cr) {
        const crow = cr.data.slice();
        crow[CU.TOTAL] = Math.max(0, (parseFloat(crow[CU.TOTAL]) || 0) - (parseFloat(row[SL.TOTAL]) || 0));
        crow[CU.PAID] = Math.max(0, (parseFloat(crow[CU.PAID]) || 0) - (parseFloat(row[SL.PAID]) || 0));
        crow[CU.DUE] = Math.round((crow[CU.TOTAL] - crow[CU.PAID]) * 100) / 100;
        getSheet(SHEETS.CUSTOMERS).getRange(cr.row, 1, 1, crow.length).setValues([crow]);
      }
    }

    try {
      const pySh = getSheet(SHEETS.PAYMENTS);
      const pyData = getSheetData(SHEETS.PAYMENTS);
      pyData.forEach((py, idx) => {
        if (py[PAY.SALE_ID] == parseInt(id) && py[PAY.TYPE] === 'customer_payment') {
          const prow = py.slice(); prow[PAY.NOTES] = '[CANCELLED] ' + (prow[PAY.NOTES] || '');
          pySh.getRange(idx + 2, 1, 1, prow.length).setValues([prow]);
        }
      });
    } catch(e) {}
    return { success: true, message: 'Sale cancelled' };
  } catch (e) { return { success: false, message: 'Failed to cancel' }; }
}

function updateSale(saleData, userId, role) {
  try {
    const { id, customer_id, discount, paid_amount, payment_method, notes, status } = saleData;
    const r = findRowByValue(SHEETS.SALES, SL.ID, parseInt(id));
    if (!r) return { success: false, message: 'Not found' };
    const row = r.data.slice();
    const oldCustId = row[SL.CUST_ID];
    const oldTotal = parseFloat(row[SL.TOTAL]) || 0;
    const oldPaid = parseFloat(row[SL.PAID]) || 0;

    const newCustId = customer_id !== undefined ? (customer_id ? parseInt(customer_id) : '') : row[SL.CUST_ID];
    row[SL.CUST_ID] = newCustId;
    const disc = discount !== undefined ? (parseFloat(discount) || 0) : (parseFloat(row[SL.DISC]) || 0);
    row[SL.DISC] = disc;
    const subtotal = parseFloat(row[SL.SUBTOTAL]) || 0;
    const newTotal = Math.round((subtotal - disc) * 100) / 100;
    row[SL.TOTAL] = newTotal;
    const paid = paid_amount !== undefined ? Math.min(parseFloat(paid_amount) || 0, newTotal) : Math.min(parseFloat(row[SL.PAID]) || 0, newTotal);
    row[SL.PAID] = paid;
    row[SL.DUE] = Math.round((newTotal - paid) * 100) / 100;
    if (payment_method) row[SL.METHOD] = payment_method;
    if (notes !== undefined) row[SL.NOTES] = notes;
    if (status && status !== 'cancelled') row[SL.STATUS] = status;
    row[SL.UPDATED] = ts();

    getSheet(SHEETS.SALES).getRange(r.row, 1, 1, row.length).setValues([row]);

    if (oldCustId && oldCustId != newCustId) {
      const ocr = findRowByValue(SHEETS.CUSTOMERS, CU.ID, parseInt(oldCustId));
      if (ocr) {
        const oc = ocr.data.slice();
        oc[CU.TOTAL] = Math.max(0, (parseFloat(oc[CU.TOTAL]) || 0) - oldTotal);
        oc[CU.PAID] = Math.max(0, (parseFloat(oc[CU.PAID]) || 0) - oldPaid);
        oc[CU.DUE] = Math.round((oc[CU.TOTAL] - oc[CU.PAID]) * 100) / 100;
        getSheet(SHEETS.CUSTOMERS).getRange(ocr.row, 1, 1, oc.length).setValues([oc]);
      }
    }
    if (newCustId) {
      const ncr = findRowByValue(SHEETS.CUSTOMERS, CU.ID, parseInt(newCustId));
      if (ncr) {
        const nc = ncr.data.slice();
        if (oldCustId == newCustId) {
          nc[CU.TOTAL] = Math.max(0, (parseFloat(nc[CU.TOTAL]) || 0) + (newTotal - oldTotal));
          nc[CU.PAID] = Math.max(0, (parseFloat(nc[CU.PAID]) || 0) + (paid - oldPaid));
        } else {
          nc[CU.TOTAL] = (parseFloat(nc[CU.TOTAL]) || 0) + newTotal;
          nc[CU.PAID] = (parseFloat(nc[CU.PAID]) || 0) + paid;
        }
        nc[CU.DUE] = Math.round((nc[CU.TOTAL] - nc[CU.PAID]) * 100) / 100;
        getSheet(SHEETS.CUSTOMERS).getRange(ncr.row, 1, 1, nc.length).setValues([nc]);
      }
    }
    return { success: true, message: 'Updated' };
  } catch (e) { return { success: false, message: 'Failed' }; }
}

function addSaleItem(saleId, itemStockId, sellRate, userId, role) {
  try {
    const sr = findRowByValue(SHEETS.SALES, SL.ID, parseInt(saleId));
    const wr = findRowByValue(SHEETS.ITEM_STOCKS, WS.ID, parseInt(itemStockId));
    if (!sr || !wr) return { success: false, message: 'Not found' };
    const w = wr.data;
    const qty = 1; // Defaulting to 1 for quick add
    const rate = parseFloat(sellRate) || parseFloat(w[WS.SELL_RATE]) || 0;
    const lineTotal = Math.round(qty * rate * 100) / 100;
    const now = ts();

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      getSheet(SHEETS.SALE_ITEMS).appendRow([getNextId(SHEETS.SALE_ITEMS), parseInt(saleId), parseInt(itemStockId), w[WS.ITEM_CODE] || '', w[WS.ITEM_NAME] || '', qty, rate, lineTotal]);

      const wrow = w.slice();
      const stockQty = parseInt(wrow[WS.QTY]) || 1;
      if (qty >= stockQty) {
        wrow[WS.STATUS] = 'sold';
        wrow[WS.QTY] = 0;
        wrow[WS.BUY_PRICE] = 0;
        wrow[WS.SELL_PRICE] = 0;
      }
      else {
        const remain = stockQty - qty; wrow[WS.QTY] = remain;
        wrow[WS.BUY_PRICE] = Math.round(remain * (parseFloat(wrow[WS.BUY_RATE]) || 0) * 100) / 100;
        wrow[WS.SELL_PRICE] = Math.round(remain * (parseFloat(wrow[WS.SELL_RATE]) || 0) * 100) / 100;
      }
      wrow[WS.UPDATED] = now;
      getSheet(SHEETS.ITEM_STOCKS).getRange(wr.row, 1, 1, wrow.length).setValues([wrow]);
      recalcSaleFromItems(sr, now);
      return { success: true, message: 'Added' };
    } finally { lock.releaseLock(); }
  } catch(e) { return { success: false, message: 'Failed' }; }
}

function returnSaleItem(saleItemId, saleId, reason, userId, role) {
  try {
    const sr = findRowByValue(SHEETS.SALES, SL.ID, parseInt(saleId));
    const siRow = findRowByValue(SHEETS.SALE_ITEMS, SI.ID, parseInt(saleItemId));
    if (!sr || !siRow) return { success: false, message: 'Not found' };

    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      const now = ts();
      const wr = findRowByValue(SHEETS.ITEM_STOCKS, WS.ID, parseInt(siRow.data[SI.WS_ID]));
      if (wr) {
        const wrow = wr.data.slice();
        const addBack = parseInt(siRow.data[SI.QTY]) || 1;
        const newQty = (parseInt(wrow[WS.QTY]) || 0) + addBack;
        wrow[WS.QTY] = newQty;
        wrow[WS.STATUS] = 'available';
        wrow[WS.BUY_PRICE] = Math.round(newQty * (parseFloat(wrow[WS.BUY_RATE]) || 0) * 100) / 100;
        wrow[WS.SELL_PRICE] = Math.round(newQty * (parseFloat(wrow[WS.SELL_RATE]) || 0) * 100) / 100;
        wrow[WS.UPDATED] = now;
        getSheet(SHEETS.ITEM_STOCKS).getRange(wr.row, 1, 1, wrow.length).setValues([wrow]);
      }
      getSheet(SHEETS.SALE_ITEMS).deleteRow(siRow.row);
      recalcSaleFromItems(findRowByValue(SHEETS.SALES, SL.ID, parseInt(saleId)), now);
      return { success: true, message: 'Returned' };
    } finally { lock.releaseLock(); }
  } catch(e) { return { success: false, message: 'Failed' }; }
}

function recalcSaleFromItems(sr, now) {
  const saleId = sr.data[SL.ID];
  const siData = getSheetData(SHEETS.SALE_ITEMS);
  const items = siData.filter(si => si[SI.SALE_ID] == saleId);
  let totalPcs = 0, subtotal = 0;
  items.forEach(si => { totalPcs += (parseInt(si[SI.QTY])||1); subtotal += (parseFloat(si[SI.TOTAL])||0); });

  const row = sr.data.slice();
  const oldTotal = parseFloat(row[SL.TOTAL]) || 0;
  const disc = parseFloat(row[SL.DISC]) || 0;
  const newTotal = Math.round((subtotal - disc) * 100) / 100;
  const paid = parseFloat(row[SL.PAID]) || 0;

  row[SL.ITEMS] = totalPcs; row[SL.SUBTOTAL] = Math.round(subtotal * 100) / 100; row[SL.TOTAL] = newTotal;
  row[SL.DUE] = Math.round(Math.max(0, newTotal - paid) * 100) / 100;
  row[SL.STATUS] = row[SL.DUE] <= 0 ? 'completed' : 'pending'; row[SL.UPDATED] = now;
  getSheet(SHEETS.SALES).getRange(sr.row, 1, 1, row.length).setValues([row]);

  if (row[SL.CUST_ID]) {
    const diff = newTotal - oldTotal;
    if (diff !== 0) {
      const cr = findRowByValue(SHEETS.CUSTOMERS, CU.ID, parseInt(row[SL.CUST_ID]));
      if (cr) {
        const crow = cr.data.slice();
        crow[CU.TOTAL] = Math.max(0, (parseFloat(crow[CU.TOTAL]) || 0) + diff);
        crow[CU.DUE] = Math.round((crow[CU.TOTAL] - (parseFloat(crow[CU.PAID]) || 0)) * 100) / 100;
        getSheet(SHEETS.CUSTOMERS).getRange(cr.row, 1, 1, crow.length).setValues([crow]);
      }
    }
  }
}

function deleteSale(id, userId, role) {
  try {
    const r = findRowByValue(SHEETS.SALES, SL.ID, parseInt(id));
    if (!r) return { success: false, message: 'Not found' };
    const row = r.data;
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      if (row[SL.STATUS] !== 'cancelled') {
        const saleItems = getSheetData(SHEETS.SALE_ITEMS).filter(si => si[SI.SALE_ID] == parseInt(id));
        saleItems.forEach(si => {
          const wr = findRowByValue(SHEETS.ITEM_STOCKS, WS.ID, parseInt(si[SI.WS_ID]));
          if (wr) {
            const wrow = wr.data.slice();
            const addBack = parseInt(si[SI.QTY]) || 1;
            const newQty = (parseInt(wrow[WS.QTY]) || 0) + addBack;
            wrow[WS.QTY] = newQty;
            wrow[WS.STATUS] = 'available';
            wrow[WS.BUY_PRICE] = Math.round(newQty * (parseFloat(wrow[WS.BUY_RATE]) || 0) * 100) / 100;
            wrow[WS.SELL_PRICE] = Math.round(newQty * (parseFloat(wrow[WS.SELL_RATE]) || 0) * 100) / 100;
            wrow[WS.UPDATED] = ts();
            getSheet(SHEETS.ITEM_STOCKS).getRange(wr.row, 1, 1, wrow.length).setValues([wrow]);
          }
        });
        if (row[SL.CUST_ID]) {
          const cr = findRowByValue(SHEETS.CUSTOMERS, CU.ID, parseInt(row[SL.CUST_ID]));
          if (cr) {
            const crow = cr.data.slice();
            crow[CU.TOTAL] = Math.max(0, (parseFloat(crow[CU.TOTAL]) || 0) - (parseFloat(row[SL.TOTAL]) || 0));
            crow[CU.PAID] = Math.max(0, (parseFloat(crow[CU.PAID]) || 0) - (parseFloat(row[SL.PAID]) || 0));
            crow[CU.DUE] = Math.round((crow[CU.TOTAL] - crow[CU.PAID]) * 100) / 100;
            getSheet(SHEETS.CUSTOMERS).getRange(cr.row, 1, 1, crow.length).setValues([crow]);
          }
        }
      }
      const siSh = getSheet(SHEETS.SALE_ITEMS);
      const pySh = getSheet(SHEETS.PAYMENTS);
      getSheetData(SHEETS.SALE_ITEMS).map((si, i) => ({ id: si[SI.SALE_ID], row: i + 2 })).filter(si => si.id == parseInt(id)).reverse().forEach(rn => siSh.deleteRow(rn.row));
      getSheetData(SHEETS.PAYMENTS).map((py, i) => ({ id: py[PAY.SALE_ID], type: py[PAY.TYPE], row: i + 2 })).filter(py => py.id == parseInt(id) && py.type === 'customer_payment').reverse().forEach(rn => pySh.deleteRow(rn.row));
      getSheet(SHEETS.SALES).deleteRow(r.row);
      return { success: true, message: 'Deleted' };
    } finally { lock.releaseLock(); }
  } catch (e) { return { success: false, message: 'Failed' }; }
}
// Expenses
function getExpenses(userId, role) {
  try {
    const data = getSheetData(SHEETS.EXPENSES);
    const uData = getSheetData(SHEETS.USERS);
    const uMap = {}; uData.forEach(r => { uMap[r[U.ID]] = r[U.NAME]; });
    const expenses = data.map(r => ({
      id: r[EX.ID], title: r[EX.TITLE], category: r[EX.CAT] || 'other', amount: parseFloat(r[EX.AMT]) || 0,
      expense_date: r[EX.DATE] instanceof Date ? r[EX.DATE].toISOString().split('T')[0] : (r[EX.DATE] || ''),
      notes: r[EX.NOTES] || '', created_by_name: uMap[r[EX.CREATED_BY]] || 'Unknown'
    }));
    return { success: true, data: expenses.reverse() };
  } catch (e) { return { success: false, message: 'Failed' }; }
}

function addExpense(exData, userId, role) {
  try {
    const sh = getSheet(SHEETS.EXPENSES);
    sh.appendRow([getNextId(SHEETS.EXPENSES), exData.title.trim(), exData.category || 'other', parseFloat(exData.amount), exData.expense_date || ts().split('T')[0], exData.notes || '', userId, ts()]);
    return { success: true, message: 'Added' };
  } catch (e) { return { success: false, message: 'Failed' }; }
}

function updateExpense(exData, userId, role) {
  try {
    const r = findRowByValue(SHEETS.EXPENSES, EX.ID, parseInt(exData.id));
    const row = r.data.slice();
    row[EX.TITLE] = exData.title.trim(); row[EX.CAT] = exData.category || 'other'; row[EX.AMT] = parseFloat(exData.amount); row[EX.DATE] = exData.expense_date || row[EX.DATE]; row[EX.NOTES] = exData.notes !== undefined ? exData.notes : row[EX.NOTES];
    getSheet(SHEETS.EXPENSES).getRange(r.row, 1, 1, row.length).setValues([row]);
    return { success: true, message: 'Updated' };
  } catch (e) { return { success: false, message: 'Failed' }; }
}

function deleteExpense(id, userId, role) {
  try {
    const r = findRowByValue(SHEETS.EXPENSES, EX.ID, parseInt(id));
    getSheet(SHEETS.EXPENSES).deleteRow(r.row);
    return { success: true, message: 'Deleted' };
  } catch (e) { return { success: false, message: 'Failed' }; }
}

// Import Logs
function getImportLogs(userId, role) {
  try {
    const data = getSheetData(SHEETS.IMPORT_LOGS);
    const catData = getSheetData(SHEETS.CATEGORIES);
    const uData = getSheetData(SHEETS.USERS);
    const catMap = {}; catData.forEach(r => { catMap[r[C.ID]] = r[C.NAME]; });
    const uMap = {}; uData.forEach(r => { uMap[r[U.ID]] = r[U.NAME]; });
    const logs = data.map(r => ({
      id: r[IL.ID], file_name: r[IL.FILE], category_id: r[IL.CAT_ID], category_name: catMap[r[IL.CAT_ID]] || 'Unknown',
      total_rows: parseInt(r[IL.TOTAL]) || 0, success_rows: parseInt(r[IL.SUCCESS]) || 0, failed_rows: parseInt(r[IL.FAILED]) || 0,
      status: r[IL.STATUS] || 'completed', error_log: r[IL.ERRORS] || '', created_by_name: uMap[r[IL.CREATED_BY]] || 'Unknown',
      created_at: r[IL.CREATED] instanceof Date ? r[IL.CREATED].toISOString() : r[IL.CREATED]
    }));
    return { success: true, data: logs.reverse() };
  } catch (e) { return { success: false, message: 'Failed' }; }
}

// Settings
function getSettings() {
  try {
    const data = getSheetData(SHEETS.SETTINGS);
    const settings = {}; data.forEach(r => { settings[r[ST.KEY]] = r[ST.VAL]; });
    return { success: true, data: settings };
  } catch(e) { return { success: true, data: {} }; }
}

function saveAllSettings(settingsObj, userId, role) {
  try {
    Object.entries(settingsObj).forEach(([key, value]) => {
      const r = findRowByValue(SHEETS.SETTINGS, ST.KEY, key);
      if (r) {
        const row = r.data.slice(); row[ST.VAL] = value; row[ST.UPDATED_BY] = userId; row[ST.UPDATED] = ts();
        getSheet(SHEETS.SETTINGS).getRange(r.row, 1, 1, row.length).setValues([row]);
      } else getSheet(SHEETS.SETTINGS).appendRow([getNextId(SHEETS.SETTINGS), key, value, userId, ts()]);
    });
    return { success: true, message: 'Settings saved' };
  } catch(e) { return { success: false, message: 'Failed' }; }
}

// Dashboard
function getDashboardStats(userId, role) {
  try {
    const now = new Date();
    const todayStr = now.toISOString().split('T')[0];
    const mStart = new Date(now.getFullYear(), now.getMonth(), 1);

    const slData = getSheetData(SHEETS.SALES);
    const activeSales = slData.filter(s => s[SL.STATUS] !== 'cancelled');
    const todaySales = activeSales.filter(s => { const d = s[SL.DATE] instanceof Date ? s[SL.DATE].toISOString().split('T')[0] : String(s[SL.DATE]).split('T')[0]; return d === todayStr; });

    const wsData = getSheetData(SHEETS.ITEM_STOCKS);
    const availStock = wsData.filter(w => w[WS.STATUS] === 'available');

    const pyData = getSheetData(SHEETS.PAYMENTS);
    const todayCollections = pyData.filter(p => p[PAY.TYPE] === 'customer_payment' && (p[PAY.DATE] instanceof Date ? p[PAY.DATE].toISOString().split('T')[0] : String(p[PAY.DATE])) === todayStr);
    const todayCollectionAmt = todayCollections.reduce((s, p) => s + (parseFloat(p[PAY.AMT]) || 0), 0);

    const custData = getSheetData(SHEETS.CUSTOMERS);
    const custMap = {}; custData.forEach(c => { custMap[c[CU.ID]] = c[CU.NAME]; });
    const totalCustDue = custData.reduce((s, c) => s + (parseFloat(c[CU.DUE]) || 0), 0);
    const puData = getSheetData(SHEETS.PURCHASES);
    const totalSuppDue = puData.filter(p => p[PU.STATUS] !== 'cancelled').reduce((s, p) => s + (parseFloat(p[PU.DUE]) || 0), 0);

    const stockValue = availStock.reduce((s, w) => s + (parseFloat(w[WS.SELL_PRICE]) || 0), 0);
    const exData = getSheetData(SHEETS.EXPENSES);
    const expThisMonth = exData.filter(e => new Date(e[EX.DATE]) >= mStart).reduce((s, e) => s + (parseFloat(e[EX.AMT]) || 0), 0);

    const dailySales = [];
    for (let i = 29; i >= 0; i--) {
      const d = new Date(now); d.setDate(now.getDate() - i);
      const ds = d.toISOString().split('T')[0];
      const dayTotal = activeSales.filter(s => { const sd = s[SL.DATE] instanceof Date ? s[SL.DATE].toISOString().split('T')[0] : String(s[SL.DATE]).split('T')[0]; return sd === ds; }).reduce((s, sl) => s + (parseFloat(sl[SL.TOTAL]) || 0), 0);
      dailySales.push({ date: ds.substring(5), amount: dayTotal });
    }

    const qtySum = (arr) => arr.reduce((s, w) => s + (parseInt(w[WS.QTY]) || 1), 0);
    const statusDist = { available: qtySum(availStock), sold: qtySum(wsData.filter(w => w[WS.STATUS]==='sold')), reserved: qtySum(wsData.filter(w => w[WS.STATUS]==='reserved')), damaged: qtySum(wsData.filter(w => w[WS.STATUS]==='damaged')) };

    const topCusts = custData.filter(c => parseFloat(c[CU.TOTAL]) > 0).map(c => ({ name: c[CU.NAME], total: parseFloat(c[CU.TOTAL]) || 0 })).sort((a, b) => b.total - a.total).slice(0, 10);

    const expBreakdown = {};
    exData.filter(e => new Date(e[EX.DATE]) >= mStart).forEach(e => { const cat = e[EX.CAT] || 'other'; expBreakdown[cat] = (expBreakdown[cat] || 0) + (parseFloat(e[EX.AMT]) || 0); });

    const recentSales = activeSales.slice(-10).reverse().map(s => ({ invoice_no: s[SL.INV_NO], total: parseFloat(s[SL.TOTAL]) || 0, customer: s[SL.CUST_ID] ? (custMap[s[SL.CUST_ID]] || 'Unknown') : 'Walk-in', date: s[SL.DATE] instanceof Date ? s[SL.DATE].toISOString() : s[SL.DATE] }));
    const recentPayments = pyData.filter(p => p[PAY.TYPE] === 'customer_payment').slice(-5).reverse().map(p => ({ amount: parseFloat(p[PAY.AMT]) || 0, method: p[PAY.METHOD] || '', date: p[PAY.DATE] instanceof Date ? p[PAY.DATE].toISOString().split('T')[0] : (p[PAY.DATE] || '') }));

    const todaySalesAmt = todaySales.reduce((s, sl) => s + (parseFloat(sl[SL.TOTAL]) || 0), 0);
    const myTodaySales = todaySales.filter(s => s[SL.CREATED_BY] == userId);
    const myTodaySalesAmt = myTodaySales.reduce((s, sl) => s + (parseFloat(sl[SL.TOTAL]) || 0), 0);
    const myTodayCollection = todayCollections.filter(p => p[PAY.CREATED_BY] == userId).reduce((s, p) => s + (parseFloat(p[PAY.AMT]) || 0), 0);
    const myMonthSales = activeSales.filter(s => s[SL.CREATED_BY] == userId && new Date(s[SL.DATE] instanceof Date ? s[SL.DATE] : s[SL.DATE]) >= mStart).reduce((s, sl) => s + (parseFloat(sl[SL.TOTAL]) || 0), 0);
    const myRecentSales = activeSales.filter(s => s[SL.CREATED_BY] == userId).slice(-10).reverse().map(s => ({ invoice_no: s[SL.INV_NO], total: parseFloat(s[SL.TOTAL]) || 0, customer: s[SL.CUST_ID] ? (custMap[s[SL.CUST_ID]] || 'Unknown') : 'Walk-in', date: s[SL.DATE] instanceof Date ? s[SL.DATE].toISOString() : s[SL.DATE] }));

    const todayImports = wsData.filter(w => { const c = w[WS.CREATED] instanceof Date ? w[WS.CREATED].toISOString().split('T')[0] : String(w[WS.CREATED]).split('T')[0]; return c === todayStr; }).length;
    const pendingPurchases = puData.filter(p => p[PU.STATUS] === 'pending').length;
    const recentItem = wsData.slice(-10).reverse().map(w => ({ item_code: w[WS.ITEM_CODE], item_name: w[WS.ITEM_NAME] || '', qty: parseInt(w[WS.QTY]) || 1, status: w[WS.STATUS], created_at: w[WS.CREATED] instanceof Date ? w[WS.CREATED].toISOString() : w[WS.CREATED] }));

    return { success: true, data: { todaySalesAmt, todayCollectionAmt, totalCustDue, totalSuppDue, stockValue, expThisMonth, dailySales, statusDist, topCusts, expBreakdown, recentSales, recentPayments, myTodaySalesAmt, myTodayCollection, myMonthSales, myRecentSales, availableStock: qtySum(availStock), todayImports, pendingPurchases, recentItem }};
  } catch (e) { return { success: false, message: 'Failed' }; }
}

function getLogs(userId, role, limit) {
  try {
    const data = getSheetData(SHEETS.LOGS);
    const logs = data.map(r => ({ id: r[L.ID], user_id: r[L.UID], username: r[L.UNAME], action: r[L.ACTION], table_name: r[L.TABLE], record_id: r[L.RID], details: r[L.DETAILS] || '', created_at: r[L.CREATED] instanceof Date ? r[L.CREATED].toISOString() : r[L.CREATED] }));
    const lim = Math.min(parseInt(limit) || 200, 500);
    return { success: true, data: logs.reverse().slice(0, lim) };
  } catch (e) { return { success: false, message: 'Failed' }; }
}

function logActivity(uid, action, table, rid, details, uname) {
  try { const sh = getSheet(SHEETS.LOGS); const name = uname || getUsernameById(uid); sh.appendRow([getNextId(SHEETS.LOGS), uid, name, action, table, rid || '', details || '', ts()]); } catch (e) {}
}

function getReportsData(reportType, filters, userId, role) {
  try {
    const f = filters || {};
    const dFrom = f.dateFrom ? new Date(f.dateFrom + 'T00:00:00') : null;
    const dTo = f.dateTo ? new Date(f.dateTo + 'T23:59:59') : null;
    const toD = v => v instanceof Date ? v : new Date(v);
    const toDs = v => { const d = toD(v); return isNaN(d) ? '' : d.toISOString().split('T')[0]; };
    const inRange = v => { if (!dFrom && !dTo) return true; const d = toD(v); if (isNaN(d)) return false; if (dFrom && d < dFrom) return false; if (dTo && d > dTo) return false; return true; };

    if (reportType === 'profit_loss') {
      const slData = getSheetData(SHEETS.SALES).filter(r => r[SL.STATUS] !== 'cancelled' && inRange(r[SL.DATE]));
      const puData = getSheetData(SHEETS.PURCHASES).filter(r => inRange(r[PU.DATE]));
      const exData = getSheetData(SHEETS.EXPENSES).filter(r => inRange(r[EX.DATE]));
      const totalSales = slData.reduce((s, r) => s + (parseFloat(r[SL.TOTAL]) || 0), 0);
      const totalPurchases = puData.reduce((s, r) => s + (parseFloat(r[PU.TOTAL]) || 0), 0);
      const totalExpenses = exData.reduce((s, r) => s + (parseFloat(r[EX.AMT]) || 0), 0);
      const dayMap = {};
      slData.forEach(r => { const d = toDs(r[SL.DATE]); if (d) { if (!dayMap[d]) dayMap[d] = { date: d, sales: 0, purchases: 0, expenses: 0 }; dayMap[d].sales += parseFloat(r[SL.TOTAL]) || 0; } });
      puData.forEach(r => { const d = toDs(r[PU.DATE]); if (d) { if (!dayMap[d]) dayMap[d] = { date: d, sales: 0, purchases: 0, expenses: 0 }; dayMap[d].purchases += parseFloat(r[PU.TOTAL]) || 0; } });
      exData.forEach(r => { const d = toDs(r[EX.DATE]); if (d) { if (!dayMap[d]) dayMap[d] = { date: d, sales: 0, purchases: 0, expenses: 0 }; dayMap[d].expenses += parseFloat(r[EX.AMT]) || 0; } });
      const dailyData = Object.values(dayMap).sort((a, b) => a.date.localeCompare(b.date));
      return { success: true, data: { totalSales, totalPurchases, totalExpenses, grossProfit: totalSales - totalPurchases, netProfit: totalSales - totalPurchases - totalExpenses, salesCount: slData.length, purchasesCount: puData.length, dailyData }};
    }

    if (reportType === 'sales_summary') {
      const slData = getSheetData(SHEETS.SALES);
      const filtered = slData.filter(r => inRange(r[SL.DATE]));
      const active = filtered.filter(r => r[SL.STATUS] !== 'cancelled');
      const totalAmount = active.reduce((s, r) => s + (parseFloat(r[SL.TOTAL]) || 0), 0);
      const totalCount = active.length;
      const avgTicket = totalCount ? totalAmount / totalCount : 0;
      const groupBy = f.groupBy || 'daily';
      const gMap = {};
      active.forEach(r => {
        const d = toD(r[SL.DATE]); if (isNaN(d)) return; let label;
        if (groupBy === 'daily') label = toDs(r[SL.DATE]);
        else if (groupBy === 'weekly') { const jan1 = new Date(d.getFullYear(), 0, 1); const wk = Math.ceil(((d - jan1) / 86400000 + jan1.getDay() + 1) / 7); label = d.getFullYear() + '-W' + String(wk).padStart(2, '0'); }
        else label = d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
        if (!gMap[label]) gMap[label] = { label, amount: 0, count: 0 }; gMap[label].amount += parseFloat(r[SL.TOTAL]) || 0; gMap[label].count++;
      });
      const groupedData = Object.values(gMap).sort((a, b) => a.label.localeCompare(b.label));
      const methodBreakdown = {}; active.forEach(r => { const m = r[SL.METHOD] || 'unknown'; methodBreakdown[m] = (methodBreakdown[m] || 0) + (parseFloat(r[SL.TOTAL]) || 0); });
      const statusBreakdown = {}; filtered.forEach(r => { const st = r[SL.STATUS] || 'unknown'; if (!statusBreakdown[st]) statusBreakdown[st] = { count: 0, amount: 0 }; statusBreakdown[st].count++; statusBreakdown[st].amount += parseFloat(r[SL.TOTAL]) || 0; });
      return { success: true, data: { totalAmount, totalCount, avgTicket, groupedData, methodBreakdown, statusBreakdown }};
    }

    if (reportType === 'stock_aging') {
      const wsData = getSheetData(SHEETS.ITEM_STOCKS);
      const now = Date.now();
      const buckets = [ { label: '0-30 days', min: 0, max: 30, count: 0, value: 0 }, { label: '31-60 days', min: 31, max: 60, count: 0, value: 0 }, { label: '61-90 days', min: 61, max: 90, count: 0, value: 0 }, { label: '90+ days', min: 91, max: Infinity, count: 0, value: 0 } ];
      const oldStock = []; let totalAvailable = 0, totalValue = 0;
      wsData.forEach(r => {
        if (r[WS.STATUS] !== 'available') return;
        const created = toD(r[WS.CREATED]); if (isNaN(created)) return;
        const days = Math.floor((now - created.getTime()) / 86400000);
        const qty = parseInt(r[WS.QTY]) || 1; const val = parseFloat(r[WS.SELL_PRICE]) || 0;
        totalAvailable += qty; totalValue += val;
        const b = buckets.find(b => days >= b.min && days <= b.max); if (b) { b.count += qty; b.value += val; }
        if (days > 90) oldStock.push({ item_code: r[WS.ITEM_CODE], item_name: r[WS.ITEM_NAME] || '', category_name: '', days, qty: qty, sell_price: val });
      });
      if (oldStock.length) {
        const catData = getSheetData(SHEETS.CATEGORIES); const catMap = catData.reduce((m, r) => (m[r[C.ID]] = r[C.NAME], m), {});
        const wsMap = wsData.reduce((m, r) => (m[r[WS.ITEM_CODE]] = r, m), {});
        oldStock.forEach(o => { const ws = wsMap[o.item_code]; if (ws) o.category_name = catMap[ws[WS.CAT_ID]] || ''; });
      }
      return { success: true, data: { buckets: buckets.map(b => ({ label: b.label, count: b.count, value: b.value })), oldStock: oldStock.sort((a, b) => b.days - a.days), totalAvailable, totalValue }};
    }

    if (reportType === 'customer_profit') {
      const slData = getSheetData(SHEETS.SALES).filter(r => r[SL.STATUS] !== 'cancelled' && inRange(r[SL.DATE]));
      const cuData = getSheetData(SHEETS.CUSTOMERS); const cuMap = cuData.reduce((m, r) => (m[r[CU.ID]] = { name: r[CU.NAME], phone: r[CU.PHONE] || '' }, m), {});
      const agg = {};
      slData.forEach(r => {
        const cid = r[SL.CUST_ID]; if (!cid) return;
        if (!agg[cid]) agg[cid] = { id: cid, name: '', phone: '', totalSales: 0, saleCount: 0 };
        agg[cid].totalSales += parseFloat(r[SL.TOTAL]) || 0; agg[cid].saleCount++;
      });
      const customers = Object.values(agg).map(c => { const cu = cuMap[c.id]; c.name = cu ? cu.name : 'Unknown'; c.phone = cu ? cu.phone : ''; c.avgTicket = c.saleCount ? c.totalSales / c.saleCount : 0; return c; }).sort((a, b) => b.totalSales - a.totalSales);
      return { success: true, data: { customers }};
    }

    if (reportType === 'category_report') {
      const catData = getSheetData(SHEETS.CATEGORIES);
      const wsData = getSheetData(SHEETS.ITEM_STOCKS);
      const puData = getSheetData(SHEETS.PURCHASES).filter(r => inRange(r[PU.DATE]));
      const slData = getSheetData(SHEETS.SALES).filter(r => r[SL.STATUS] !== 'cancelled' && inRange(r[SL.DATE]));
      const siData = getSheetData(SHEETS.SALE_ITEMS);

      const saleMap = slData.reduce((m, r) => (m[r[SL.ID]] = r, m), {});
      const wsCatMap = wsData.reduce((m, r) => (m[r[WS.ID]] = r[WS.CAT_ID], m), {});
      const puAgg = {};
      puData.forEach(r => {
        const cId = r[PU.CAT_ID];
        if (!puAgg[cId]) puAgg[cId] = { cost: 0, qty: 0 };
        puAgg[cId].cost += parseFloat(r[PU.TOTAL]) || 0;
        puAgg[cId].qty += parseInt(r[PU.QTY]) || 0;
      });

      const slAgg = {};
      siData.forEach(r => {
        if (!saleMap[r[SI.SALE_ID]]) return;
        const cId = wsCatMap[r[SI.WS_ID]]; if (!cId) return;
        if (!slAgg[cId]) slAgg[cId] = { revenue: 0 };
        slAgg[cId].revenue += parseFloat(r[SI.TOTAL]) || 0;
      });

      const categories_rep = catData.map(r => {
        const cId = r[C.ID];
        const pu = puAgg[cId] || { cost: 0, qty: 0 };
        const sl = slAgg[cId] || { revenue: 0 };
        return {
          id: cId, name: r[C.NAME],
          totalQty: pu.qty, purchaseCost: pu.cost, salesRevenue: sl.revenue, profit: sl.revenue - pu.cost
        };
      }).sort((a, b) => b.profit - a.profit);
      return { success: true, data: { categories_rep }};
    }
    return { success: false, message: 'Unknown report type' };
  } catch (e) { return { success: false, message: 'Failed' }; }
}

function getOverdueSummary(userId, role) {
  try {
    const cuData = getSheetData(SHEETS.CUSTOMERS); const slData = getSheetData(SHEETS.SALES);
    const puData = getSheetData(SHEETS.PURCHASES); const spData = getSheetData(SHEETS.SUPPLIERS);
    const now = Date.now();
    const toD = v => v instanceof Date ? v : new Date(v);
    const daysBetween = v => { const d = toD(v); return isNaN(d) ? 0 : Math.floor((now - d.getTime()) / 86400000); };
    const bucket = days => days <= 7 ? 'within_7' : days <= 15 ? 'within_15' : days <= 30 ? 'within_30' : 'over_30';
    const emptyBuckets = () => ({ within_7: 0, within_15: 0, within_30: 0, over_30: 0 });

    const custOldest = {};
    slData.forEach(r => {
      if (r[SL.STATUS] === 'cancelled' || (parseFloat(r[SL.DUE]) || 0) <= 0) return;
      const cid = r[SL.CUST_ID]; if (!cid) return;
      const d = toD(r[SL.DATE]); if (isNaN(d)) return;
      if (!custOldest[cid] || d < custOldest[cid]) custOldest[cid] = d;
    });

    const custBuckets = emptyBuckets();
    const customers = cuData.filter(r => (parseFloat(r[CU.DUE]) || 0) > 0).map(r => {
      const id = r[CU.ID], due = parseFloat(r[CU.DUE]) || 0; const oldest = custOldest[id];
      const days = oldest ? daysBetween(oldest) : 0; if (days > 0) custBuckets[bucket(days)]++;
      return { id, name: r[CU.NAME], phone: r[CU.PHONE] || '', total_due: due, oldest_sale_date: oldest ? oldest.toISOString() : '', days_overdue: days };
    }).sort((a, b) => b.days_overdue - a.days_overdue);

    const suppAgg = {};
    puData.forEach(r => {
      if (r[PU.STATUS] === 'cancelled' || (parseFloat(r[PU.DUE]) || 0) <= 0) return;
      const sid = r[PU.SUPPLIER_ID]; const d = toD(r[PU.DATE]); const due = parseFloat(r[PU.DUE]) || 0;
      if (!suppAgg[sid]) suppAgg[sid] = { total_due: 0, oldest: null };
      suppAgg[sid].total_due += due;
      if (!isNaN(d) && (!suppAgg[sid].oldest || d < suppAgg[sid].oldest)) suppAgg[sid].oldest = d;
    });

    const spMap = spData.reduce((m, r) => (m[r[SP.ID]] = r, m), {});
    const suppBuckets = emptyBuckets();
    const suppliers = Object.keys(suppAgg).map(sid => {
      const a = suppAgg[sid], sp = spMap[sid]; const days = a.oldest ? daysBetween(a.oldest) : 0;
      if (days > 0) suppBuckets[bucket(days)]++;
      return { id: parseInt(sid), name: sp ? sp[SP.NAME] : 'Unknown', phone: sp ? (sp[SP.PHONE] || '') : '', total_due: Math.round(a.total_due * 100) / 100, oldest_date: a.oldest ? a.oldest.toISOString() : '', days_overdue: days };
    }).sort((a, b) => b.days_overdue - a.days_overdue);

    const totalCustDue = customers.reduce((s, c) => s + c.total_due, 0); const totalSuppDue = suppliers.reduce((s, c) => s + c.total_due, 0);
    return { success: true, data: { customers, suppliers, buckets: { customers: custBuckets, suppliers: suppBuckets }, totalCustDue: Math.round(totalCustDue * 100) / 100, totalSuppDue: Math.round(totalSuppDue * 100) / 100 } };
  } catch (e) { return { success: false, message: 'Failed' }; }
}

function getSaleForDuplicate(saleId, userId, role) {
  try {
    const sr = findRowByValue(SHEETS.SALES, SL.ID, parseInt(saleId));
    if (!sr) return { success: false, message: 'Not found' };
    const siData = getSheetData(SHEETS.SALE_ITEMS);
    const saleItems = siData.filter(r => r[SI.SALE_ID] == parseInt(saleId));

    const wsData = getSheetData(SHEETS.ITEM_STOCKS);
    const wsSerialMap = {}; wsData.forEach(r => { if (r[WS.STATUS] === 'available') wsSerialMap[r[WS.ITEM_CODE]] = r[WS.ID]; });

    const custId = sr.data[SL.CUST_ID]; let custName = '';
    if (custId) { const cr = findRowByValue(SHEETS.CUSTOMERS, CU.ID, parseInt(custId)); if (cr) custName = cr.data[CU.NAME]; }

    const items = saleItems.map(r => ({
      serial: r[SI.ITEM_CODE] || '', qty: parseInt(r[SI.QTY]) || 1, rate: parseFloat(r[SI.RATE]) || 0,
      original_total: parseFloat(r[SI.TOTAL]) || 0, available_item_id: wsSerialMap[r[SI.ITEM_CODE]] || null
    }));
    return { success: true, data: { customer_id: custId || null, customer_name: custName, items } };
  } catch (e) { return { success: false, message: 'Failed' }; }
}

function setupDemoData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let tmp = ss.insertSheet('_tmp_setup');
  ss.getSheets().forEach(s => { if (s.getName() !== '_tmp_setup') ss.deleteSheet(s); });
  const now = ts();

  let sh = ss.insertSheet(SHEETS.USERS);
  sh.appendRow(['ID','Full Name','Email','Phone','Password','Role','Avatar','Is Active','Created At','Updated At','OTP','OTP Expires']);
  sh.getRange(1, 1, 1, 12).setFontWeight('bold').setBackground('#001f3f').setFontColor('white');
  sh.appendRow([1, 'Admin User', 'admin@demo.com', '03001000001', 'admin123', 'admin', '', 1, now, now, '', '']);
  sh.appendRow([2, 'Manager 1', 'manager1@demo.com', '03001000002', 'manager123', 'manager', '', 1, now, '', '', '']);
  sh.appendRow([3, 'Cashier 1', 'cashier1@demo.com', '03001000003', 'cashier123', 'cashier', '', 1, now, '', '', '']);
  sh.appendRow([4, 'Warehouse 1', 'warehouse1@demo.com', '03001000005', 'warehouse123', 'warehouse_staff', '', 1, now, '', '', '']);

  sh = ss.insertSheet(SHEETS.CATEGORIES);
  sh.appendRow(['ID','Name','Description','Is Active','Created By','Created At']);
  sh.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#001f3f').setFontColor('white');
  sh.appendRow([1, 'Electronics', 'Electronic Items', 1, 1, now]);
  sh.appendRow([2, 'Books', 'Publications', 1, 1, now]);
  sh.appendRow([3, 'Gadgets', 'Smart gadgets', 1, 1, now]);

  sh = ss.insertSheet(SHEETS.SUPPLIERS);
  sh.appendRow(['ID','Name','Phone','Address','Is Active','Created By','Created At']);
  sh.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#001f3f').setFontColor('white');
  sh.appendRow([1, 'Supplier 1', '03001000010', 'Demo City', 1, 1, now]);
  sh.appendRow([2, 'Supplier 2', '03001000011', 'Demo City', 1, 1, now]);

  sh = ss.insertSheet(SHEETS.ITEM_STOCKS);
  sh.appendRow(['ID','Item Code','Item Name','Category ID','Purchase ID','Buy Rate','Buy Price','Sell Rate','Sell Price','Status','Notes','Created By','Created At','Updated At','Qty','Image']);
  sh.getRange(1, 1, 1, 16).setFontWeight('bold').setBackground('#001f3f').setFontColor('white');
  var demoStocks = [
    [1,'ITEM-001','Laptop',1,1, 100,0, 150,0,'available','',1,now,now,10,''],
    [2,'ITEM-002','Mouse',1,1, 200,0, 250,0,'available','',1,now,now,5,''],
    [3,'ITEM-003','Notebook',2,2, 50,0, 80,0,'available','',1,now,now,20,''],
    [4,'ITEM-004','Pen',2,2, 500,0, 600,0,'sold','',1,now,now,2,''],
    [5,'ITEM-005','Smartwatch',3,3, 1000,0, 1200,0,'available','',2,now,now,3,''],
  ];
  demoStocks.forEach(function(r) { var q = r[14] || 1; r[6] = q * r[5]; r[8] = q * r[7]; sh.appendRow(r); });

  sh = ss.insertSheet(SHEETS.CUSTOMERS);
  sh.appendRow(['ID','Name','Phone','Address','Total Purchase','Total Paid','Total Due','Is Active','Created By','Created At']);
  sh.getRange(1, 1, 1, 10).setFontWeight('bold').setBackground('#001f3f').setFontColor('white');
  sh.appendRow([1, 'Customer 1', '03001000020', 'Demo City', 1200, 1200, 0, 1, 1, now]);
  sh.appendRow([2, 'Customer 2', '03001000021', 'Demo City', 0, 0, 0, 1, 1, now]);

  sh = ss.insertSheet(SHEETS.PURCHASES);
  sh.appendRow(['ID','Purchase No','Supplier ID','Category ID','Total Qty','Avg Rate','Total Amount','Paid Amount','Due Amount','Purchase Date','Notes','Status','Created By','Created At','Updated At']);
  sh.getRange(1, 1, 1, 15).setFontWeight('bold').setBackground('#001f3f').setFontColor('white');
  sh.appendRow([1, 'PUR-20260410-001', 1, 1, 15, 133.33, 2000, 2000, 0, '2026-04-10', '', 'completed', 1, now, now]);
  sh.appendRow([2, 'PUR-20260415-001', 2, 2, 22, 90.90, 2000, 1000, 1000, '2026-04-15', '', 'pending', 1, now, now]);
  sh.appendRow([3, 'PUR-20260418-001', 1, 3, 3, 1000, 3000, 1500, 1500, '2026-04-18', '', 'pending', 2, now, now]);

  sh = ss.insertSheet(SHEETS.SALES);
  sh.appendRow(['ID','Invoice No','Customer ID','Total Items','Subtotal','Discount','Grand Total','Paid Amount','Due Amount','Payment Method','Sale Date','Status','Notes','Created By','Created At','Updated At']);
  sh.getRange(1, 1, 1, 16).setFontWeight('bold').setBackground('#001f3f').setFontColor('white');
  sh.appendRow([1, 'INV-20260423-001', 1, 2, 1200, 0, 1200, 1200, 0, 'cash', '2026-04-23T10:00:00Z', 'completed', '', 3, now, now]);

  sh = ss.insertSheet(SHEETS.SALE_ITEMS);
  sh.appendRow(['ID','Sale ID','Item Stock ID','Item Code','Item Name','Qty','Rate','Line Total']);
  sh.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#001f3f').setFontColor('white');
  sh.appendRow([1, 1, 4, 'ITEM-004', 'Pen', 2, 600, 1200]);

  sh = ss.insertSheet(SHEETS.EXPENSES);
  sh.appendRow(['ID','Title','Category','Amount','Expense Date','Notes','Created By','Created At']);
  sh.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#001f3f').setFontColor('white');
  sh.appendRow([1, 'Transport fee', 'transport', 500, '2026-04-24', '', 1, now]);
  sh.appendRow([2, 'Office rent', 'rent', 5000, '2026-04-01', '', 1, now]);

  sh = ss.insertSheet(SHEETS.PAYMENTS);
  sh.appendRow(['ID','Sale ID','Purchase ID','Payment Type','Amount','Payment Method','Reference No','Payment Date','Notes','Created By','Created At']);
  sh.getRange(1, 1, 1, 11).setFontWeight('bold').setBackground('#001f3f').setFontColor('white');
  sh.appendRow([1, 1, '', 'customer_payment', 1200, 'cash', '', '2026-04-23', 'Full', 3, now]);
  sh.appendRow([2, '', 1, 'supplier_payment', 2000, 'bank', '', '2026-04-10', 'Full', 1, now]);
  sh.appendRow([3, '', 2, 'supplier_payment', 1000, 'bkash', '', '2026-04-15', 'Partial', 1, now]);
  sh.appendRow([4, '', 3, 'supplier_payment', 1500, 'cash', '', '2026-04-18', 'Partial', 2, now]);

  sh = ss.insertSheet(SHEETS.IMPORT_LOGS);
  sh.appendRow(['ID','File Name','Category ID','Total Rows','Success Rows','Failed Rows','Status','Error Log','Created By','Created At']);
  sh.getRange(1, 1, 1, 10).setFontWeight('bold').setBackground('#001f3f').setFontColor('white');

  sh = ss.insertSheet(SHEETS.SETTINGS);
  sh.appendRow(['ID','Setting Key','Setting Value','Updated By','Updated At']);
  sh.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#001f3f').setFontColor('white');
  sh.appendRow([1, 'business_name', 'CDS Publication', 1, now]);
  sh.appendRow([2, 'business_address', 'Chattogram, Bangladesh', 1, now]);
  sh.appendRow([3, 'business_phone', '03001000001', 1, now]);
  sh.appendRow([4, 'currency_symbol', '৳', 1, now]);
  sh.appendRow([5, 'default_buy_rate', '100', 1, now]);
  sh.appendRow([6, 'default_sell_rate', '150', 1, now]);
  sh.appendRow([7, 'invoice_prefix', 'INV', 1, now]);
  sh.appendRow([8, 'purchase_prefix', 'PUR', 1, now]);
  sh.appendRow([9, 'invoice_footer', 'Thank you for your business!', 1, now]);

  sh = ss.insertSheet(SHEETS.LOGS);
  sh.appendRow(['ID','User ID','Username','Action','Table','Record ID','Details','Created At']);
  sh.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#001f3f').setFontColor('white');

  ss.deleteSheet(tmp);
  return 'Setup done! Login: admin@demo.com / admin123';
}
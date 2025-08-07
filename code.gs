const SHEET_NAMES = {
  ITEMS: 'items',
  GROUPS: 'groups',
  LOGBOOK: 'logbook',
  FORECAST: 'forecast'
};
const PHOTO_FOLDER_ID = 'REPLACE_WITH_DRIVE_FOLDER_ID';
const LOW_STOCK_RECIPIENTS = ['example@example.com'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Inventory Manager');
}

function initializeSpreadsheet() {
  const ss = SpreadsheetApp.getActive();
  if (!ss.getSheetByName(SHEET_NAMES.ITEMS)) {
    const sh = ss.insertSheet(SHEET_NAMES.ITEMS);
    sh.appendRow(['ID','Name','Make','Model','Quantity','CriticalQty']);
  }
  if (!ss.getSheetByName(SHEET_NAMES.GROUPS)) {
    const sh = ss.insertSheet(SHEET_NAMES.GROUPS);
    sh.appendRow(['GroupName','ItemIDs']);
  }
  if (!ss.getSheetByName(SHEET_NAMES.LOGBOOK)) {
    const sh = ss.insertSheet(SHEET_NAMES.LOGBOOK);
    sh.appendRow(['Timestamp','Direction','ItemID','Name','Quantity','PlantCode','ServiceCode','RequisitionNo','Remarks','PhotoURL','User']);
  }
  if (!ss.getSheetByName(SHEET_NAMES.FORECAST)) {
    const sh = ss.insertSheet(SHEET_NAMES.FORECAST);
    sh.appendRow(['ItemID','Name','AvgMonthlyUsage','NextQuarterNeed','LastUpdated']);
  }
}

function getItems() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAMES.ITEMS);
  const values = sh.getDataRange().getValues();
  const items = [];
  for (let i = 1; i < values.length; i++) {
    const [id,name,make,model,qty,critical] = values[i];
    items.push({id,name,make,model,qty,critical});
  }
  return items;
}

function getGroups() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAMES.GROUPS);
  const values = sh.getDataRange().getValues();
  const groups = [];
  for (let i = 1; i < values.length; i++) {
    const [name, ids] = values[i];
    if (!name) continue;
    groups.push({name, ids: ids.split(',').map(id=>id.trim())});
  }
  return groups;
}

function getStock(itemId) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAMES.ITEMS);
  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] == itemId) {
      return {qty: values[i][4], critical: values[i][5], name: values[i][1]};
    }
  }
  return null;
}

function recordTransaction(data) {
  if (data.groupName) {
    const groups = getGroups();
    const group = groups.find(g=>g.name === data.groupName);
    if (group) {
      group.ids.forEach(id => {
        const clone = Object.assign({}, data, {itemId: id, groupName: ''});
        recordTransaction(clone);
      });
    }
    return 'Group transaction recorded';
  }
  const ss = SpreadsheetApp.getActive();
  const itemSheet = ss.getSheetByName(SHEET_NAMES.ITEMS);
  const logSheet = ss.getSheetByName(SHEET_NAMES.LOGBOOK);
  const values = itemSheet.getDataRange().getValues();
  let row = -1;
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] == data.itemId) {
      row = i + 1;
      break;
    }
  }
  if (row === -1) throw new Error('Item not found');
  const currentQty = itemSheet.getRange(row,5).getValue();
  const newQty = data.direction === 'IN' ? currentQty + data.quantity : currentQty - data.quantity;
  itemSheet.getRange(row,5).setValue(newQty);
  const photoUrl = savePhoto(data.photo, data.itemId);
  logSheet.appendRow([
    new Date(),
    data.direction,
    data.itemId,
    itemSheet.getRange(row,2).getValue(),
    data.quantity,
    data.plantCode,
    data.serviceCode,
    data.requisition,
    data.remarks,
    photoUrl,
    Session.getActiveUser().getEmail()
  ]);
  return 'Recorded';
}

function savePhoto(photo, itemId) {
  if (!photo) return '';
  const folder = DriveApp.getFolderById(PHOTO_FOLDER_ID);
  const contentType = photo.match(/^data:(image\/\w+);base64,/)[1];
  const bytes = Utilities.base64Decode(photo.split(',')[1]);
  const blob = Utilities.newBlob(bytes, contentType, itemId + '_' + Date.now() + '.jpg');
  const file = folder.createFile(blob);
  return file.getUrl();
}

function checkLowStock() {
  const items = getItems();
  const low = items.filter(it => it.qty <= it.critical);
  if (!low.length) return;
  let html = '<table border="1" cellpadding="5"><tr><th>Item</th><th>Qty</th><th>Critical</th></tr>';
  low.forEach(it=> {
    html += '<tr><td>'+it.name+'</td><td>'+it.qty+'</td><td>'+it.critical+'</td></tr>';
  });
  html += '</table>';
  MailApp.sendEmail({
    to: LOW_STOCK_RECIPIENTS.join(','),
    subject: 'Low Stock Alert',
    htmlBody: html
  });
}

function forecastUsage() {
  const ss = SpreadsheetApp.getActive();
  const logSheet = ss.getSheetByName(SHEET_NAMES.LOGBOOK);
  const forecastSheet = ss.getSheetByName(SHEET_NAMES.FORECAST);
  const data = logSheet.getDataRange().getValues();
  const usage = {};
  const today = new Date();
  const threeMonthsAgo = new Date(today.getFullYear(), today.getMonth()-3, today.getDate());
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const ts = row[0];
    const direction = row[1];
    const itemId = row[2];
    const qty = row[4];
    if (direction === 'OUT' && ts >= threeMonthsAgo) {
      if (!usage[itemId]) usage[itemId] = 0;
      usage[itemId] += qty;
    }
  }
  forecastSheet.clearContents();
  forecastSheet.appendRow(['ItemID','Name','AvgMonthlyUsage','NextQuarterNeed','LastUpdated']);
  const items = getItems();
  items.forEach(it=>{
    const outQty = usage[it.id] || 0;
    const monthly = outQty / 3;
    const need = Math.ceil(monthly * 3);
    forecastSheet.appendRow([it.id, it.name, monthly, need, new Date()]);
  });
}

function createMonthlyTrigger() {
  ScriptApp.newTrigger('monthlyTasks')
    .timeBased()
    .everyMonths(1)
    .create();
}

function monthlyTasks() {
  checkLowStock();
  forecastUsage();
}

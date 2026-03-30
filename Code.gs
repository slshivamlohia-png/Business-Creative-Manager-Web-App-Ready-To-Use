// ==================== CREATIVE MANAGER v7 ====================
const SHEETS = { CREATIVES:'Creatives', USERS:'Users', SETTINGS:'Settings', PRODUCTS:'Products', ATTRIBUTES:'Attributes', ACTIVITY:'ActivityLog' };

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Neodrift Creative Manager')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ==================== INIT ====================
function initializeSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const expectedHeaders = ['Code','Name','Type','Status','FileId','MimeType','ThumbnailUrl','UploadedBy','UploadDate','ApprovedBy','ApprovalDate','Rating','ChangesRequired','TestedGoogle','TestedMeta','TestedAmazon','TestedDate','ProductCategory','AINamed','TestedOthers','TestedOthersText','OriginalFileName','LastAIRun','CustomAttributes'];
  let cs = ss.getSheetByName(SHEETS.CREATIVES);
  if (!cs) { cs = ss.insertSheet(SHEETS.CREATIVES); cs.getRange(1,1,1,expectedHeaders.length).setValues([expectedHeaders]); cs.setFrozenRows(1); }
  else {
    const lc = Math.max(cs.getLastColumn(), expectedHeaders.length);
    const cur = cs.getRange(1,1,1,lc).getValues()[0];
    for (let i = 0; i < expectedHeaders.length; i++) { if (String(cur[i]||'').trim() !== expectedHeaders[i]) cs.getRange(1,i+1).setValue(expectedHeaders[i]); }
  }
  // NOTE: deduplicateCreativeCodes removed from init to avoid masking real duplicates
  let us = ss.getSheetByName(SHEETS.USERS);
  if (!us) { us = ss.insertSheet(SHEETS.USERS); us.getRange(1,1,1,4).setValues([['Username','Password','Role','DisplayName']]); us.getRange(2,1,1,4).setValues([['admin','admin123','admin','Administrator']]); us.setFrozenRows(1); }
  let st = ss.getSheetByName(SHEETS.SETTINGS);
  if (!st) { st = ss.insertSheet(SHEETS.SETTINGS); st.getRange(1,1,4,2).setValues([['Setting','Value'],['NomenclatureFormat','{product}_{platform}_{variant}_{date}'],['ClaudeApiKey',''],['DriveFolderId','']]); }
  let pr = ss.getSheetByName(SHEETS.PRODUCTS);
  if (!pr) { pr = ss.insertSheet(SHEETS.PRODUCTS); pr.getRange(1,1,1,1).setValues([['ProductName']]); pr.setFrozenRows(1); }
  let at = ss.getSheetByName(SHEETS.ATTRIBUTES);
  if (!at) { at = ss.insertSheet(SHEETS.ATTRIBUTES); at.getRange(1,1,1,3).setValues([['AttributeName','AttributeType','Options']]); at.setFrozenRows(1); }
  let al = ss.getSheetByName(SHEETS.ACTIVITY);
  if (!al) { al = ss.insertSheet(SHEETS.ACTIVITY); al.getRange(1,1,1,7).setValues([['Timestamp','Action','Code','OldValue','NewValue','User','UndoJSON']]); al.setFrozenRows(1); }
  return { success: true, message: 'Initialized. Login: admin / admin123' };
}

// ==================== ACTIVITY LOG ====================
function logActivity(action, code, oldVal, newVal, user, undoJSON) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.ACTIVITY);
    if (!sheet) return;
    sheet.appendRow([new Date().toISOString(), action, code || '', oldVal || '', newVal || '', user || '', undoJSON || '']);
  } catch(e) {}
}

function getActivityLog(limit) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.ACTIVITY);
  if (!sheet || sheet.getLastRow() <= 1) return [];
  const data = sheet.getDataRange().getValues();
  const results = [];
  for (let i = data.length - 1; i >= 1; i--) {
    results.push({ rowIndex: i+1, timestamp: data[i][0], action: data[i][1], code: data[i][2], oldValue: data[i][3], newValue: data[i][4], user: data[i][5], undoJSON: data[i][6] });
    if (results.length >= (limit || 100)) break;
  }
  return results;
}

function undoAction(rowIndex) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.ACTIVITY);
  const data = sheet.getDataRange().getValues();
  if (rowIndex < 2 || rowIndex > data.length) return { success: false, message: 'Invalid log entry' };
  const row = data[rowIndex - 1];
  const undoStr = row[6];
  if (!undoStr) return { success: false, message: 'No undo data for this action' };
  try {
    const undo = JSON.parse(undoStr);
    const cs = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CREATIVES);
    const cdata = cs.getDataRange().getValues();
    const h = cdata[0];
    
    if (undo.type === 'update') {
      for (let i = 1; i < cdata.length; i++) {
        if (String(cdata[i][0]) === String(undo.code)) {
          Object.keys(undo.fields).forEach(key => {
            const ci = h.indexOf(key);
            if (ci !== -1) cs.getRange(i+1, ci+1).setValue(undo.fields[key]);
          });
          break;
        }
      }
    } else if (undo.type === 'delete_restore') {
      cs.appendRow(undo.rowData);
    } else if (undo.type === 'bulk_update') {
      undo.items.forEach(item => {
        for (let i = 1; i < cdata.length; i++) {
          if (String(cdata[i][0]) === String(item.code)) {
            Object.keys(item.fields).forEach(key => {
              const ci = h.indexOf(key);
              if (ci !== -1) cs.getRange(i+1, ci+1).setValue(item.fields[key]);
            });
            break;
          }
        }
      });
    }
    // Mark as undone in log
    sheet.getRange(rowIndex, 2).setValue('[UNDONE] ' + row[1]);
    sheet.getRange(rowIndex, 7).setValue('');
    logActivity('Undo: ' + row[1], undo.code || '', row[4], row[3], 'admin', '');
    return { success: true };
  } catch(e) { return { success: false, message: e.message }; }
}

// ==================== UNIQUE CODE ====================
function generateUniqueCode() {
  const d = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyMMdd');
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CREATIVES).getDataRange().getValues();
  const allCodes = new Set(); for (let i=1;i<data.length;i++) allCodes.add(String(data[i][0]));
  const prefix = 'ND-' + d + '-';
  let max = 0;
  for (const c of allCodes) { if (c.startsWith(prefix)) { const n = parseInt(c.substring(prefix.length)); if (n > max) max = n; } }
  let candidate; do { max++; candidate = prefix + String(max).padStart(4,'0'); } while (allCodes.has(candidate));
  return candidate;
}

function makeFilenameUnique(name, currentFileId) {
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CREATIVES).getDataRange().getValues();
  const h = data[0]; const ni = h.indexOf('Name'), fi = h.indexOf('FileId');
  const existing = new Set();
  for (let i=1;i<data.length;i++) { if (String(data[i][fi]) !== String(currentFileId)) existing.add(String(data[i][ni]).toLowerCase()); }
  if (!existing.has(name.toLowerCase())) return name;
  let c = 2, cand; do { cand = name + '_v' + c; c++; } while (existing.has(cand.toLowerCase()));
  return cand;
}

// ==================== AI PASSWORD (IST) ====================
function verifyAIPassword(entered) {
  const now = new Date();
  const ist = Utilities.formatDate(now, 'Asia/Kolkata', 'ddMMyy,HH').split(',');
  return { success: entered === ist[0] + ist[1] };
}

// ==================== INTEGRITY REPORT ====================
function getIntegrityReport() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CREATIVES);
  const data = sheet.getDataRange().getValues();
  const h = data[0];
  const fi = h.indexOf('FileId'), ni = h.indexOf('Name'), ci = 0;
  
  const sheetCount = data.length - 1;
  const dupCodes = [], dupNames = [], missingFiles = [];
  const codeMap = new Map(), nameMap = new Map();
  
  for (let i=1;i<data.length;i++) {
    const code = String(data[i][ci]), name = String(data[i][ni]).toLowerCase(), fid = String(data[i][fi]);
    if (codeMap.has(code)) { dupCodes.push({ row: i+1, code }); if (!dupCodes.find(d => d.row === codeMap.get(code))) dupCodes.unshift({ row: codeMap.get(code), code }); }
    else codeMap.set(code, i+1);
    if (nameMap.has(name)) { dupNames.push({ row: i+1, name: String(data[i][ni]) }); if (!dupNames.find(d => d.row === nameMap.get(name))) dupNames.unshift({ row: nameMap.get(name), name: String(data[i][ni]) }); }
    else nameMap.set(name, i+1);
  }
  
  let driveCount = 0;
  try {
    const folderId = getOrCreateFolder();
    const folder = DriveApp.getFolderById(folderId);
    function countFiles(f) {
      const files = f.getFiles(); while (files.hasNext()) { files.next(); driveCount++; }
      const subs = f.getFolders(); while (subs.hasNext()) countFiles(subs.next());
    }
    countFiles(folder);
  } catch(e) {}
  
  const fileIds = [];
  for (let i=1;i<data.length;i++) fileIds.push(String(data[i][fi]));
  for (let i=0;i<fileIds.length;i++) {
    try { DriveApp.getFileById(fileIds[i]); } catch(e) { missingFiles.push({ row: i+2, fileId: fileIds[i] }); }
  }
  
  return {
    sheetCount, driveCount,
    duplicateCodes: dupCodes,
    duplicateNames: dupNames,
    missingFiles: missingFiles,
    inSync: sheetCount === driveCount && dupCodes.length === 0 && missingFiles.length === 0
  };
}

// ==================== FIX FUNCTIONS ====================
function fixDuplicateCodes() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CREATIVES);
  const data = sheet.getDataRange().getValues();
  const seen = new Map(); // code -> first row index
  let fixed = 0;
  for (let i = 1; i < data.length; i++) {
    const code = String(data[i][0]);
    if (seen.has(code)) {
      const d = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyMMdd');
      const allCodes = new Set();
      // Gather all current codes including already-fixed ones
      for (let j = 1; j < data.length; j++) allCodes.add(String(data[j][0]));
      for (const c of seen.keys()) allCodes.add(c);
      const prefix = 'ND-' + d + '-';
      let max = 0;
      for (const c of allCodes) { if (c.startsWith(prefix)) { const n = parseInt(c.substring(prefix.length)); if (n > max) max = n; } }
      let nc; do { max++; nc = prefix + String(max).padStart(4,'0'); } while (allCodes.has(nc));
      const oldCode = code;
      sheet.getRange(i+1, 1).setValue(nc);
      data[i][0] = nc; // Update in-memory too
      seen.set(nc, i+1);
      logActivity('Fix duplicate code', oldCode, oldCode, nc, 'system', JSON.stringify({type:'update',code:nc,fields:{Code:oldCode}}));
      fixed++;
    } else {
      seen.set(code, i+1);
    }
  }
  return { success: true, fixed };
}

function fixDuplicateNames() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CREATIVES);
  const data = sheet.getDataRange().getValues();
  const h = data[0]; const ni = h.indexOf('Name');
  const seen = new Map();
  let fixed = 0;
  for (let i = 1; i < data.length; i++) {
    const name = String(data[i][ni]);
    const key = name.toLowerCase();
    if (seen.has(key)) {
      // Make unique by appending _v2, _v3 etc
      const allNames = new Set();
      for (let j = 1; j < data.length; j++) allNames.add(String(data[j][ni]).toLowerCase());
      let c = 2, newName;
      do { newName = name + '_v' + c; c++; } while (allNames.has(newName.toLowerCase()));
      sheet.getRange(i+1, ni+1).setValue(newName);
      data[i][ni] = newName;
      logActivity('Fix duplicate name', String(data[i][0]), name, newName, 'system', JSON.stringify({type:'update',code:String(data[i][0]),fields:{Name:name}}));
      fixed++;
    } else {
      seen.set(key, i+1);
    }
  }
  return { success: true, fixed };
}

function fixMissingFiles() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CREATIVES);
  const data = sheet.getDataRange().getValues();
  const h = data[0]; const fi = h.indexOf('FileId');
  const rowsToDelete = [];
  for (let i = 1; i < data.length; i++) {
    const fid = String(data[i][fi]);
    try { DriveApp.getFileById(fid); } catch(e) { rowsToDelete.push(i+1); }
  }
  for (let i = rowsToDelete.length - 1; i >= 0; i--) {
    logActivity('Remove missing file entry', String(data[rowsToDelete[i]-1][0]), '', '', 'system', '');
    sheet.deleteRow(rowsToDelete[i]);
  }
  return { success: true, removed: rowsToDelete.length };
}

// ==================== AUTH / URLS ====================
function getBackendUrls() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return { sheetUrl: ss.getUrl(), driveUrl: 'https://drive.google.com/drive/folders/' + getOrCreateFolder() };
}

function login(u, p) {
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.USERS).getDataRange().getValues();
  for (let i=1;i<data.length;i++) { if (data[i][0]===u && data[i][1]===p) return { success:true, user:{username:data[i][0],role:data[i][2],displayName:data[i][3]} }; }
  return { success:false, message:'Invalid credentials' };
}

// ==================== DRIVE ====================
function getOrCreateFolder() {
  const settings = getSettings();
  let fid = settings.DriveFolderId;
  if (fid) { try { DriveApp.getFolderById(fid); return fid; } catch(e) {} }
  const folder = DriveApp.createFolder('Neodrift Creatives');
  fid = folder.getId();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.SETTINGS);
  const data = sheet.getDataRange().getValues();
  for (let i=1;i<data.length;i++) { if (data[i][0]==='DriveFolderId') { sheet.getRange(i+1,2).setValue(fid); break; } }
  return fid;
}

// ==================== UPLOAD ====================
function initResumableUpload(fileName, mimeType, fileSize) {
  const folderId = getOrCreateFolder();
  const token = ScriptApp.getOAuthToken();
  const resp = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=resumable', {
    method:'POST', headers:{'Authorization':'Bearer '+token,'Content-Type':'application/json; charset=UTF-8','X-Upload-Content-Type':mimeType,'X-Upload-Content-Length':fileSize},
    payload:JSON.stringify({name:fileName,parents:[folderId]}), muteHttpExceptions:true
  });
  if (resp.getResponseCode()===200) return { success:true, uploadUri:resp.getHeaders()['Location']||resp.getHeaders()['location'], token };
  return { success:false, message:resp.getContentText() };
}

function registerUploadedFile(fileId, fileName, mimeType, uploadedBy, originalFileName) {
  try {
    DriveApp.getFileById(fileId).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const code = generateUniqueCode();
    const type = mimeType.startsWith('image/') ? 'Photo' : 'Video';
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CREATIVES).appendRow([
      code,fileName,type,'Pending Approval',fileId,mimeType,'',uploadedBy,new Date().toISOString(),'','','','',false,false,false,'','',false,false,'',originalFileName||fileName,'','{}']);
    moveToProductFolder(fileId,'Unknown');
    logActivity('Upload', code, '', fileName, uploadedBy, '');
    return { success:true, code, name:fileName, type, fileId };
  } catch(e) { return { success:false, message:e.message }; }
}

function getAccessToken() { return ScriptApp.getOAuthToken(); }

function uploadFile(fileData, fileName, mimeType, uploadedBy, originalFileName) {
  const folder = DriveApp.getFolderById(getOrCreateFolder());
  const blob = Utilities.newBlob(Utilities.base64Decode(fileData), mimeType, fileName);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const code = generateUniqueCode();
  const type = mimeType.startsWith('image/') ? 'Photo' : 'Video';
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CREATIVES).appendRow([
    code,fileName,type,'Pending Approval',file.getId(),mimeType,'',uploadedBy,new Date().toISOString(),'','','','',false,false,false,'','',false,false,'',originalFileName||fileName,'','{}']);
  moveToProductFolder(file.getId(),'Unknown');
  logActivity('Upload', code, '', fileName, uploadedBy, '');
  return { success:true, code, name:fileName, type, fileId:file.getId() };
}

// ==================== REPLACE CREATIVE FILE ====================
function replaceCreativeFile(code, fileData, fileName, mimeType) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CREATIVES);
  const data = sheet.getDataRange().getValues();
  const h = data[0];
  for (let i=1;i<data.length;i++) {
    if (String(data[i][0]) === String(code)) {
      const oldFileId = data[i][h.indexOf('FileId')];
      const category = data[i][h.indexOf('ProductCategory')] || 'Unknown';
      const mainFolder = DriveApp.getFolderById(getOrCreateFolder());
      let targetFolder;
      const sf = mainFolder.getFoldersByName(category);
      if (sf.hasNext()) targetFolder = sf.next(); else targetFolder = mainFolder.createFolder(category);
      const blob = Utilities.newBlob(Utilities.base64Decode(fileData), mimeType, fileName);
      const newFile = targetFolder.createFile(blob);
      newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      const newFileId = newFile.getId();
      sheet.getRange(i+1, h.indexOf('FileId')+1).setValue(newFileId);
      sheet.getRange(i+1, h.indexOf('MimeType')+1).setValue(mimeType);
      sheet.getRange(i+1, h.indexOf('Status')+1).setValue('Pending Approval');
      sheet.getRange(i+1, h.indexOf('ChangesRequired')+1).setValue('');
      sheet.getRange(i+1, h.indexOf('Type')+1).setValue(mimeType.startsWith('image/')?'Photo':'Video');
      try { DriveApp.getFileById(oldFileId).setTrashed(true); } catch(e) {}
      logActivity('Replace file', code, oldFileId, newFileId, '', '');
      return { success:true, newFileId };
    }
  }
  return { success:false, message:'Creative not found' };
}

// ==================== CREATIVES CRUD ====================
function getCreatives(filters) {
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CREATIVES).getDataRange().getValues();
  const h = data[0]; const results = [];
  for (let i=1;i<data.length;i++) {
    const row = {}; h.forEach((k,j) => row[k] = data[i][j]);
    row.rowIndex = i+1;
    ['AINamed','TestedGoogle','TestedMeta','TestedAmazon','TestedOthers'].forEach(f => { row[f] = row[f]===true||row[f]==='TRUE'||row[f]==='true'; });
    try { row.CustomAttributes = JSON.parse(row.CustomAttributes || '{}'); } catch(e) { row.CustomAttributes = {}; }
    if (filters) {
      if (filters.status && row.Status !== filters.status) continue;
      if (filters.type && row.Type !== filters.type) continue;
      if (filters.uploadedBy && row.UploadedBy !== filters.uploadedBy) continue;
      if (filters.search) { const s = filters.search.toLowerCase(); if (!String(row.Name).toLowerCase().includes(s) && !String(row.Code).toLowerCase().includes(s)) continue; }
      if (filters.customAttr && filters.customAttrValue) {
        const cv = row.CustomAttributes[filters.customAttr];
        if (cv !== filters.customAttrValue) continue;
      }
    }
    results.push(row);
  }
  return results.reverse();
}

function updateCreative(code, updates) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CREATIVES);
  const data = sheet.getDataRange().getValues(); const h = data[0];
  for (let i=1;i<data.length;i++) {
    if (String(data[i][0]) === String(code)) {
      Object.keys(updates).forEach(key => { const ci = h.indexOf(key); if (ci !== -1) sheet.getRange(i+1,ci+1).setValue(updates[key]); });
      return { success:true };
    }
  }
  return { success:false, message:'Not found' };
}

function updateCustomAttribute(code, attrName, attrValue) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CREATIVES);
  const data = sheet.getDataRange().getValues();
  const h = data[0]; const caIdx = h.indexOf('CustomAttributes');
  for (let i=1;i<data.length;i++) {
    if (String(data[i][0]) === String(code)) {
      let attrs = {};
      try { attrs = JSON.parse(data[i][caIdx] || '{}'); } catch(e) {}
      attrs[attrName] = attrValue;
      sheet.getRange(i+1, caIdx+1).setValue(JSON.stringify(attrs));
      return { success:true };
    }
  }
  return { success:false };
}

function approveCreative(code, rating, approvedBy) {
  // Get old values for undo
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CREATIVES).getDataRange().getValues();
  const h = data[0];
  for (let i=1;i<data.length;i++) {
    if (String(data[i][0]) === String(code)) {
      const oldStatus = data[i][h.indexOf('Status')];
      const oldRating = data[i][h.indexOf('Rating')];
      logActivity('Approve', code, oldStatus, 'Approved (★'+rating+')', approvedBy,
        JSON.stringify({type:'update',code,fields:{Status:oldStatus,Rating:oldRating,ApprovedBy:data[i][h.indexOf('ApprovedBy')],ApprovalDate:data[i][h.indexOf('ApprovalDate')],ChangesRequired:data[i][h.indexOf('ChangesRequired')]}}));
      break;
    }
  }
  return updateCreative(code, {Status:'Approved',Rating:rating,ApprovedBy:approvedBy,ApprovalDate:new Date().toISOString(),ChangesRequired:''});
}

function rejectCreative(code, changes, rejectedBy) {
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CREATIVES).getDataRange().getValues();
  const h = data[0];
  for (let i=1;i<data.length;i++) {
    if (String(data[i][0]) === String(code)) {
      const oldStatus = data[i][h.indexOf('Status')];
      logActivity('Reject', code, oldStatus, 'Changes Required', rejectedBy,
        JSON.stringify({type:'update',code,fields:{Status:oldStatus,ChangesRequired:data[i][h.indexOf('ChangesRequired')],ApprovedBy:data[i][h.indexOf('ApprovedBy')],ApprovalDate:data[i][h.indexOf('ApprovalDate')]}}));
      break;
    }
  }
  return updateCreative(code, {Status:'Changes Required',ChangesRequired:changes,ApprovedBy:rejectedBy,ApprovalDate:new Date().toISOString()});
}

function revertToPending(code, user) {
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CREATIVES).getDataRange().getValues();
  const h = data[0];
  for (let i=1;i<data.length;i++) {
    if (String(data[i][0]) === String(code)) {
      const oldStatus = data[i][h.indexOf('Status')];
      const oldRating = data[i][h.indexOf('Rating')];
      logActivity('Revert to Pending', code, oldStatus, 'Pending Approval', user,
        JSON.stringify({type:'update',code,fields:{Status:oldStatus,Rating:oldRating,ApprovedBy:data[i][h.indexOf('ApprovedBy')],ApprovalDate:data[i][h.indexOf('ApprovalDate')]}}));
      break;
    }
  }
  return updateCreative(code, {Status:'Pending Approval',Rating:'',ApprovedBy:'',ApprovalDate:'',ChangesRequired:''});
}

function markTested(code, platform, value, othersText) {
  const updates = {}; updates['Tested'+platform] = value;
  if (platform === 'Others') updates['TestedOthersText'] = othersText || '';
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CREATIVES).getDataRange().getValues();
  const h = data[0];
  for (let i=1;i<data.length;i++) {
    if (String(data[i][0])===String(code)) {
      let g=data[i][h.indexOf('TestedGoogle')],m=data[i][h.indexOf('TestedMeta')],a=data[i][h.indexOf('TestedAmazon')],o=data[i][h.indexOf('TestedOthers')];
      if (platform==='Google')g=value;if(platform==='Meta')m=value;if(platform==='Amazon')a=value;if(platform==='Others')o=value;
      updates.Status=(g||m||a||o)?'Tested on Ads':'Approved';
      if((g||m||a||o)&&!data[i][h.indexOf('TestedDate')])updates.TestedDate=new Date().toISOString();
      const oldStatus = data[i][h.indexOf('Status')];
      logActivity('Test '+platform+(value?' on':' off'), code, oldStatus, updates.Status, '',
        JSON.stringify({type:'update',code,fields:{['Tested'+platform]:data[i][h.indexOf('Tested'+platform)],Status:oldStatus,TestedDate:data[i][h.indexOf('TestedDate')],TestedOthersText:data[i][h.indexOf('TestedOthersText')]||''}}));
      break;
    }
  }
  return updateCreative(code, updates);
}

function updateCreativeCode(oldCode, newCode) {
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CREATIVES).getDataRange().getValues();
  for (let i=1;i<data.length;i++) { if (String(data[i][0])===String(newCode)&&String(data[i][0])!==String(oldCode)) return {success:false,message:'Code exists'}; }
  logActivity('Rename code', oldCode, oldCode, newCode, '', JSON.stringify({type:'update',code:newCode,fields:{Code:oldCode}}));
  return updateCreative(oldCode, {Code:newCode});
}

function updateCreativeName(code, newName) {
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CREATIVES).getDataRange().getValues();
  const h = data[0];
  for (let i=1;i<data.length;i++) {
    if (String(data[i][0]) === String(code)) {
      const oldName = data[i][h.indexOf('Name')];
      logActivity('Rename', code, oldName, newName, '', JSON.stringify({type:'update',code,fields:{Name:oldName}}));
      break;
    }
  }
  return updateCreative(code, {Name:newName});
}

function deleteCreative(code, user) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CREATIVES);
  const data = sheet.getDataRange().getValues();
  const h = data[0];
  for (let i=1;i<data.length;i++) {
    if (String(data[i][0])===String(code)) {
      const rowData = data[i];
      logActivity('Delete', code, String(data[i][h.indexOf('Name')]), '', user||'', JSON.stringify({type:'delete_restore',code,rowData:rowData}));
      try { DriveApp.getFileById(data[i][4]).setTrashed(true); } catch(e) {}
      sheet.deleteRow(i+1); return { success:true };
    }
  }
  return { success:false, message:'Not found' };
}

function bulkDeleteCreatives(codes, user) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CREATIVES);
  const data = sheet.getDataRange().getValues();
  const codeSet = new Set(codes);
  const rowsToDelete = [];
  for (let i=1;i<data.length;i++) {
    if (codeSet.has(String(data[i][0]))) {
      logActivity('Bulk delete', String(data[i][0]), String(data[i][1]), '', user||'', '');
      try { DriveApp.getFileById(data[i][4]).setTrashed(true); } catch(e) {}
      rowsToDelete.push(i+1);
    }
  }
  for (let i=rowsToDelete.length-1;i>=0;i--) sheet.deleteRow(rowsToDelete[i]);
  return { success:true, deleted: rowsToDelete.length };
}

function resubmitCreative(code) { return updateCreative(code, {Status:'Pending Approval',ChangesRequired:''}); }

function bulkUpdateCategory(fileIds, category, user) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CREATIVES);
  const data = sheet.getDataRange().getValues();
  const h = data[0]; const catIdx=h.indexOf('ProductCategory'),fidIdx=h.indexOf('FileId');
  const fset = new Set(fileIds); let updated=0;
  const undoItems = [];
  for (let i=1;i<data.length;i++) {
    if (fset.has(String(data[i][fidIdx]))) {
      const oldCat = data[i][catIdx];
      undoItems.push({code:String(data[i][0]),fields:{ProductCategory:oldCat}});
      sheet.getRange(i+1,catIdx+1).setValue(category);
      try { moveToProductFolder(String(data[i][fidIdx]), category); } catch(e) {}
      updated++;
    }
  }
  logActivity('Bulk category → '+category, updated+' items', '', category, user||'', JSON.stringify({type:'bulk_update',items:undoItems}));
  return { success:true, updated };
}

// ==================== DASHBOARD ====================
function getDashboardStats() {
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CREATIVES).getDataRange().getValues();
  const today = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd');
  const s = {total:0,pendingApproval:0,changesRequired:0,approved:0,testedOnAds:0,photos:0,videos:0,uploadedToday:0,approvedToday:0,testedToday:0,dailyUploads:{},dailyApprovals:{},dailyTested:{}};
  for (let i=1;i<data.length;i++) {
    s.total++;
    const st=data[i][3],tp=data[i][2],ud=data[i][8]?String(data[i][8]).substring(0,10):'',ad=data[i][10]?String(data[i][10]).substring(0,10):'';
    if(st==='Pending Approval')s.pendingApproval++;else if(st==='Changes Required')s.changesRequired++;else if(st==='Approved')s.approved++;else if(st==='Tested on Ads')s.testedOnAds++;
    if(tp==='Photo')s.photos++;else s.videos++;
    if(ud){s.dailyUploads[ud]=(s.dailyUploads[ud]||0)+1;if(ud===today)s.uploadedToday++}
    if(ad&&(st==='Approved'||st==='Tested on Ads')){s.dailyApprovals[ad]=(s.dailyApprovals[ad]||0)+1;if(ad===today)s.approvedToday++}
    if(st==='Tested on Ads'){const td=data[i][16]?String(data[i][16]).substring(0,10):ad;if(td){s.dailyTested[td]=(s.dailyTested[td]||0)+1;if(td===today)s.testedToday++}}
  }
  return s;
}

// ==================== SETTINGS / PRODUCTS / USERS ====================
function getSettings() { const d=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.SETTINGS).getDataRange().getValues();const s={};for(let i=1;i<d.length;i++)s[d[i][0]]=d[i][1];return s; }
function updateSetting(k,v) { const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.SETTINGS);const d=sh.getDataRange().getValues();for(let i=1;i<d.length;i++){if(d[i][0]===k){sh.getRange(i+1,2).setValue(v);return{success:true}}}sh.appendRow([k,v]);return{success:true}; }
function getProducts() { const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.PRODUCTS);if(!sh)return[];return sh.getDataRange().getValues().slice(1).filter(r=>r[0]).map((r,i)=>({rowIndex:i+2,name:r[0]})); }
function addProduct(n) { SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.PRODUCTS).appendRow([n]);return{success:true}; }
function deleteProduct(ri) { SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.PRODUCTS).deleteRow(ri);return{success:true}; }
function getUsers() { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.USERS).getDataRange().getValues().slice(1).map(r=>({username:r[0],role:r[2],displayName:r[3]})); }
function addUser(u,p,r,d) { const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.USERS);const da=sh.getDataRange().getValues();for(let i=1;i<da.length;i++){if(da[i][0]===u)return{success:false,message:'Exists'}}sh.appendRow([u,p,r,d]);return{success:true}; }
function deleteUser(u) { if(u==='admin')return{success:false,message:'Cannot delete admin'};const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.USERS);const d=sh.getDataRange().getValues();for(let i=1;i<d.length;i++){if(d[i][0]===u){sh.deleteRow(i+1);return{success:true}}}return{success:false}; }

// ==================== CUSTOM ATTRIBUTES ====================
function getAttributes() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.ATTRIBUTES);
  if (!sh) return [];
  const data = sh.getDataRange().getValues();
  return data.slice(1).filter(r => r[0]).map((r,i) => ({
    rowIndex: i+2, name: r[0], type: r[1] || 'dropdown',
    options: r[2] ? String(r[2]).split(',').map(s => s.trim()) : []
  }));
}
function addAttribute(name, type, options) { SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.ATTRIBUTES).appendRow([name, type || 'dropdown', options || '']);return { success: true }; }
function deleteAttribute(rowIndex) { SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.ATTRIBUTES).deleteRow(rowIndex);return { success: true }; }

// ==================== AI NAMING ====================
function analyzeCreativeWithAI(fileId, mimeType) {
  const settings = getSettings();
  const apiKey = settings.ClaudeApiKey;
  if (!apiKey) return { success:false, message:'API key not set' };
  const products = getProducts();
  const productList = products.map(p => '- '+p.name).join('\n');
  const productNames = products.map(p => p.name);
  let imageBase64 = '';
  try {
    const resp = UrlFetchApp.fetch('https://drive.google.com/thumbnail?id='+fileId+'&sz=w800',{headers:{'Authorization':'Bearer '+ScriptApp.getOAuthToken()},muteHttpExceptions:true});
    if (resp.getResponseCode()!==200) return {success:false,message:'Thumbnail error'};
    imageBase64 = Utilities.base64Encode(resp.getContent());
  } catch(e) { return {success:false,message:e.message}; }
  try {
    const resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method:'POST',headers:{'Content-Type':'application/json','x-api-key':apiKey,'anthropic-version':'2023-06-01'},
      payload:JSON.stringify({model:'claude-haiku-4-5-20251001',max_tokens:300,
        messages:[{role:'user',content:[{type:'image',source:{type:'base64',media_type:'image/png',data:imageBase64}},{type:'text',text:'Creative naming for Neodrift auto accessories. Identify product from catalog, suggest filename: {product}_{platform}_{variant}_{date}\n\nCATALOG:\n'+(productList||'(empty)')+'\n\nRules: lowercase_underscores, platform if visible, variant type.\n\nPRODUCT: <name or Unknown>\nFILENAME: <suggestion>'}]}]}),muteHttpExceptions:true
    });
    const result = JSON.parse(resp.getContentText());
    if (result.error) return {success:false,message:result.error.message};
    if (!result.content||!result.content[0]) return {success:false,message:'No response'};
    const text = result.content[0].text.trim();
    const pm=text.match(/PRODUCT:\s*(.+)/i),fm=text.match(/FILENAME:\s*(.+)/i);
    let sn=fm?fm[1].trim():text, pc=pm?pm[1].trim():'Unknown';
    if (pc.toLowerCase()==='unknown') pc='Unknown';
    else { const mp=productNames.find(p=>p.toLowerCase()===pc.toLowerCase()); pc=mp||'Unknown'; }
    if (pc!=='Unknown') {
      if(sn.toLowerCase().startsWith('unknown_'))sn=sn.substring(8);
      sn=sn.replace(/_unknown_/gi,'_').replace(/_unknown$/gi,'');
      const ps=pc.toLowerCase().replace(/\s+/g,'_');
      if(!sn.toLowerCase().includes(ps))sn=ps+'_'+sn;
    }
    sn = makeFilenameUnique(sn, fileId);
    const sr = saveAIResult(fileId, sn, pc);
    if (!sr.success) return {success:false,message:sr.message};
    logActivity('AI Name', sr.code||'', '', sn+' ['+pc+']', '', '');
    return {success:true,suggestedName:sn,productCategory:pc,fileId};
  } catch(e) { return {success:false,message:e.message}; }
}

function saveAIResult(fileId, name, category) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CREATIVES);
    const data = sheet.getDataRange().getValues(); const h = data[0];
    const ci = {}; h.forEach((x,j) => { ci[String(x).trim()] = j; });
    if (ci.FileId===undefined) return {success:false,message:'FileId missing'};
    for (let i=1;i<data.length;i++) {
      if (String(data[i][ci.FileId]).trim()===String(fileId).trim()) {
        if(ci.Name!==undefined)sheet.getRange(i+1,ci.Name+1).setValue(name);
        if(ci.ProductCategory!==undefined)sheet.getRange(i+1,ci.ProductCategory+1).setValue(category);
        if(ci.AINamed!==undefined)sheet.getRange(i+1,ci.AINamed+1).setValue(true);
        if(ci.LastAIRun!==undefined)sheet.getRange(i+1,ci.LastAIRun+1).setValue(new Date().toISOString());
        try{moveToProductFolder(fileId,category)}catch(e){}
        return {success:true,code:String(data[i][0])};
      }
    }
    return {success:false,message:'FileId not found'};
  } catch(e) { return {success:false,message:e.message}; }
}

// ==================== SYNC / FOLDER ====================
function syncDriveWithSheet() {
  const folderId=getOrCreateFolder();const folder=DriveApp.getFolderById(folderId);
  const sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CREATIVES);
  const data=sheet.getDataRange().getValues();const h=data[0];const cols={};h.forEach((x,i)=>{cols[x]=i});
  const sheetFids=new Set();for(let i=1;i<data.length;i++){if(data[i][cols.FileId])sheetFids.add(String(data[i][cols.FileId]))}
  const driveFiles=[];
  function scan(f,fn){const fi=f.getFiles();while(fi.hasNext()){const file=fi.next();driveFiles.push({id:file.getId(),name:file.getName(),mimeType:file.getMimeType(),folder:fn})}const s=f.getFolders();while(s.hasNext()){const sf=s.next();scan(sf,sf.getName())}}
  scan(folder,'');
  let added=0;
  for(const f of driveFiles){if(!sheetFids.has(f.id)&&(f.mimeType.startsWith('image/')||f.mimeType.startsWith('video/'))){
    const code=generateUniqueCode();sheet.appendRow([code,f.name,f.mimeType.startsWith('image/')?'Photo':'Video','Pending Approval',f.id,f.mimeType,'','sync',new Date().toISOString(),'','','','',false,false,false,'',f.folder||'Unknown',false,false,'',f.name,'','{}']);added++}}
  const dids=new Set(driveFiles.map(f=>f.id));let removed=0;
  for(let i=data.length-1;i>=1;i--){const fid=String(data[i][cols.FileId]);if(fid&&!dids.has(fid)){sheet.deleteRow(i+1);removed++}}
  logActivity('Sync', '', '+'+added, '-'+removed, '', '');
  return {success:true,added,removed,sheetCount:data.length-1+added-removed};
}

function moveToProductFolder(fileId, cat) {
  try {
    const mf=DriveApp.getFolderById(getOrCreateFolder());const file=DriveApp.getFileById(fileId);
    const c=cat||'Unknown';let sub;const f=mf.getFoldersByName(c);if(f.hasNext())sub=f.next();else sub=mf.createFolder(c);
    file.moveTo(sub);return{success:true};
  } catch(e) { return {success:false,message:e.message}; }
}

// ==================== IMPORT ====================
function listDriveFolders(query) {
  const results=[],mid=getOrCreateFolder();
  try {
    const folders=query?DriveApp.searchFolders('title contains "'+query.replace(/"/g,'\\"')+'"'):DriveApp.searchFolders('mimeType = "application/vnd.google-apps.folder"');
    let c=0;while(folders.hasNext()&&c<50){const f=folders.next();if(f.getId()===mid)continue;let fc=0;const fi=f.getFiles();while(fi.hasNext()){fi.next();fc++}results.push({id:f.getId(),name:f.getName(),fileCount:fc});c++}
  } catch(e) { return {success:false,message:e.message}; }
  return {success:true,folders:results};
}

function importFromDriveLink(url) {
  let m=url.match(/\/folders\/([a-zA-Z0-9_-]+)/);if(!m)m=url.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  const fid=m?m[1]:(/^[a-zA-Z0-9_-]{10,}$/.test(url.trim())?url.trim():null);
  if(!fid) return {success:false,message:'Invalid URL'};
  try { const f=DriveApp.getFolderById(fid);return{success:true,folderId:fid,folderName:f.getName()}; }
  catch(e) { return {success:false,message:'Cannot access folder'}; }
}

function importFromDriveFolder(sourceFolderId, preserveFolderName) {
  try {
    const mf=DriveApp.getFolderById(getOrCreateFolder());const sf=DriveApp.getFolderById(sourceFolderId);
    const catName=preserveFolderName||sf.getName();
    const sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CREATIVES);
    const data=sheet.getDataRange().getValues();const efids=new Set();const fidx=data[0].indexOf('FileId');
    for(let i=1;i<data.length;i++)efids.add(String(data[i][fidx]));
    let tf;const sfs=mf.getFoldersByName(catName);if(sfs.hasNext())tf=sfs.next();else tf=mf.createFolder(catName);
    const files=[];function scan(f){const fi=f.getFiles();while(fi.hasNext()){const file=fi.next();if(file.getMimeType().startsWith('image/')||file.getMimeType().startsWith('video/'))files.push(file)}const s=f.getFolders();while(s.hasNext())scan(s.next())}
    scan(sf);
    let imported=0,skipped=0;
    for(const file of files){if(efids.has(file.getId())){skipped++;continue}const copy=file.makeCopy(file.getName(),tf);copy.setSharing(DriveApp.Access.ANYONE_WITH_LINK,DriveApp.Permission.VIEW);const code=generateUniqueCode();
    sheet.appendRow([code,file.getName(),file.getMimeType().startsWith('image/')?'Photo':'Video','Pending Approval',copy.getId(),file.getMimeType(),'','import',new Date().toISOString(),'','','','',false,false,false,'',catName,false,false,'',file.getName(),'','{}']);efids.add(copy.getId());imported++}
    logActivity('Import from '+sf.getName(), '', '', imported+' imported', '', '');
    return {success:true,imported,skipped,total:files.length,folderName:sf.getName(),categoryName:catName};
  } catch(e) { return {success:false,message:e.message}; }
}

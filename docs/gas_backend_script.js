function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      logErrorToSheet("データが空、またはLINE以外からのアクセスです");
      return ContentService.createTextOutput("OK");
    }

    const json = JSON.parse(e.postData.contents);
    const events = json.events;
    
    if (!events || events.length === 0) return ContentService.createTextOutput("OK");
    const event = events[0];
    if (event.type !== 'message' || event.message.type !== 'text') return ContentService.createTextOutput("OK");
    
    const text = event.message.text;
    
    // ▼ 1. 古いシートへの書き込み処理
    try {
      const oldData = extractDataOld(text);
      if (!oldData.error) {
        const oldSpreadsheet = SpreadsheetApp.getActiveSpreadsheet(); 
        if (!oldSpreadsheet) throw new Error("アクティブなスプレッドシートが見つかりません。");

        const currentYear = new Date().getFullYear();
        const oldSheetName = currentYear + "年" + oldData.month + "月";
        
        let oldSheet = oldSpreadsheet.getSheetByName(oldSheetName);
        if (!oldSheet) {
          oldSheet = oldSpreadsheet.insertSheet(oldSheetName);
          // A列をエラー詳細に変更
          oldSheet.appendRow(["エラー詳細", "日付", "コマ数", "コード名", "稼働者", "配達件数", "受託率", "配達完了率"]);
          oldSheet.getRange(1, 1, 1, 8).setBorder(true, true, true, true, true, true);
          oldSheet.setFrozenRows(1);
        }
        oldSheet.appendRow(oldData.data);
        const oldLastRow = oldSheet.getLastRow();
        oldSheet.getRange(oldLastRow, 1, 1, 8).setBorder(true, true, true, true, true, true);
      }
    } catch(errOld) {
      logErrorToSheet("古いシートへの書き込みエラー: " + errOld.message);
    }

    // ▼ 2. 新しいシートへの書き込み処理
    try {
      const newData = extractDataNew(text);
      if (!newData.error) {
        const newSpreadsheet = SpreadsheetApp.openById("1j0w86gZU_sCvVVcQVQ8R6rzDRaBieTfCnZqPkboREN0");
        const currentYear = new Date().getFullYear();
        const newSheetName = currentYear + "年" + newData.month + "月";
        
        let newSheet = newSpreadsheet.getSheetByName(newSheetName);
        if (!newSheet) {
          newSheet = newSpreadsheet.insertSheet(newSheetName);
          // A列をエラー詳細に変更
          newSheet.appendRow(["エラー詳細", "日付", "コマ数", "コード名", "メール稼働者", "稼働者(申告)", "配達件数", "受託率", "配達完了率", "未確認ドライバー"]);
          newSheet.getRange(1, 1, 1, 10).setBorder(true, true, true, true, true, true);
          newSheet.setFrozenRows(1);
        }
        newSheet.appendRow(newData.data);
        const newLastRow = newSheet.getLastRow();
        newSheet.getRange(newLastRow, 1, 1, 10).setBorder(true, true, true, true, true, true);
      }
    } catch(errNew) {
      logErrorToSheet("新しいシートへの書き込みエラー: " + errNew.message);
    }
  } catch(err) {
    logErrorToSheet("システム全体エラー: " + err.message + "\n" + err.stack);
  }
  return ContentService.createTextOutput("OK");
}

function logErrorToSheet(errorMessage) {
  try {
    const spreadsheet = SpreadsheetApp.openById("1j0w86gZU_sCvVVcQVQ8R6rzDRaBieTfCnZqPkboREN0");
    const sheet = spreadsheet.getSheets()[0]; 
    sheet.appendRow([new Date(), "【システムエラー】", errorMessage]);
  } catch (e) {}
}

function extractDataOld(text) {
  function toHalfWidth(str) {
    if (!str) return '';
    return str.replace(/[！-～]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0)).replace(/[ 　\t]+/g, '');
  }
  const normalizedText = toHalfWidth(text);
  const dateMatch = normalizedText.match(/日付.*?([0-9]+)月([0-9]+)日/);
  const slotLineMatch = normalizedText.match(/コマ数([^\n]+)/);
  const codeMatch = normalizedText.match(/コード名([^\n]+)/);
  const workerMatch = normalizedText.match(/稼働者([^\n]+)/);
  
  // 「配達件数」と「配達件」の両方に対応
  const deliveryCountMatch = normalizedText.match(/配達件(?:数)?.*?([0-9]+).*?件/);
  
  const acceptanceRateMatch = normalizedText.match(/受[託諾]率.*?([0-9]+).*?%/);
  const completionRateMatch = normalizedText.match(/配達完了率.*?([0-9]+).*?%/);

  if (!dateMatch || !workerMatch) return { error: true };
  let slotValue = '';
  if (slotLineMatch) {
    const line = slotLineMatch[1].toUpperCase();
    if (line.includes('A') && line.includes('B')) slotValue = 'AB';
    else if (line.includes('A')) slotValue = 'A';
    else if (line.includes('B')) slotValue = 'B';
  }
  
  let formatErrorMsg = '';
  if (!deliveryCountMatch) {
    formatErrorMsg = 'フォーマットエラー: 配達件数が読み取れません';
  }

  const data = [
    formatErrorMsg, dateMatch ? dateMatch[1] + "月" + dateMatch[2] + "日" : '', slotValue,
    codeMatch ? codeMatch[1].trim() : '', workerMatch ? workerMatch[1].trim() : '',
    deliveryCountMatch ? parseInt(deliveryCountMatch[1], 10) : '',
    acceptanceRateMatch ? parseInt(acceptanceRateMatch[1], 10) : '',
    completionRateMatch ? parseInt(completionRateMatch[1], 10) : ''
  ];
  return { error: false, data: data, month: dateMatch[1] };
}

function extractDataNew(text) {
  function toHalfWidth(str) {
    if (!str) return '';
    return str.replace(/[！-～]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0)).replace(/[ 　\t]+/g, '');
  }
  const normalizedText = toHalfWidth(text);
  const dateMatch = normalizedText.match(/日付.*?([0-9]+)月([0-9]+)日/);
  const slotLineMatch = normalizedText.match(/コマ数([^\n]+)/);
  const codeMatch = normalizedText.match(/コード名([^\n]+)/);
  const workerMatch = normalizedText.match(/稼働者([^\n]+)/);
  
  // 「配達件数」と「配達件」の両方に対応
  const deliveryCountMatch = normalizedText.match(/配達件(?:数)?.*?([0-9]+).*?件/);
  
  const acceptanceRateMatch = normalizedText.match(/受[託諾]率.*?([0-9]+).*?%/);
  const completionRateMatch = normalizedText.match(/配達完了率.*?([0-9]+).*?%/);

  if (!dateMatch || !workerMatch) return { error: true };
  let slotValue = '';
  if (slotLineMatch) {
    const line = slotLineMatch[1].toUpperCase();
    if (line.includes('A') && line.includes('B')) slotValue = 'AB';
    else if (line.includes('A')) slotValue = 'A';
    else if (line.includes('B')) slotValue = 'B';
  }

  const targetMonth = dateMatch[1];
  const targetDay = dateMatch[2];
  
  const namesDict = getVerifiedNames(targetMonth, targetDay);
  
  const codeNameRaw = codeMatch ? codeMatch[1].trim() : '';
  const codeNameNoSpace = codeNameRaw.replace(/\s+/g, '');
  const workerNameRaw = workerMatch[1].trim(); 
  
  const allVerifiedNames = namesDict.a.concat(namesDict.b);
  const isVerified = allVerifiedNames.includes(codeNameNoSpace);
  
  let systemName = '';
  let jColumnMsg = '';
  if (isVerified) {
    systemName = codeNameRaw;
  } else {
    let listStr = allVerifiedNames.join(",");
    if (listStr === "") listStr = "メール見つからず (" + namesDict.debug + ")";
    jColumnMsg = "未確認: " + codeNameRaw + " (リスト: " + listStr + ")";
  }

  // A列に入れるエラー内容を判定
  let formatErrorMsg = '';
  if (!deliveryCountMatch) {
    formatErrorMsg = 'フォーマットエラー: 配達件数が読み取れません';
  } else if (!isVerified) {
    formatErrorMsg = '未確認エラー: メールリストに名前がありません';
  }

  const data = [
    formatErrorMsg, dateMatch[1] + "月" + dateMatch[2] + "日", slotValue,
    codeNameRaw, systemName, workerNameRaw,
    deliveryCountMatch ? parseInt(deliveryCountMatch[1], 10) : '',
    acceptanceRateMatch ? parseInt(acceptanceRateMatch[1], 10) : '',
    completionRateMatch ? parseInt(completionRateMatch[1], 10) : '',
    jColumnMsg
  ];
  return { error: false, data: data, month: targetMonth };
}

function getVerifiedNames(targetMonth, targetDay) {
  const monthStr = ('0' + targetMonth).slice(-2);
  const dayStr = ('0' + targetDay).slice(-2);
  const targetDateStr = monthStr + '-' + dayStr; 
  
  const threads = GmailApp.search('newer_than:7d');
  let listA = []; let listB = [];
  let debugLog = `検索件数:${threads.length}件`; 
  
  for (let i = 0; i < threads.length; i++) {
    const messages = threads[i].getMessages();
    for (let j = 0; j < messages.length; j++) {
      const msg = messages[j];
      const subject = msg.getSubject();
      
      if (subject.includes('ドライバー稼働状況') && subject.includes(targetDateStr)) {
        debugLog += ` | 対象メール発見!`;
        const timeMatch = subject.match(/(\d{1,2}):(\d{2})\s*時点/);
        if (timeMatch) {
          const hour = parseInt(timeMatch[1], 10);
          const min = parseInt(timeMatch[2], 10);
          
          if (hour === 14 && min >= 0 && min <= 20) {
            const body = msg.getPlainBody();
            const names = extractNamesFromBody(body);
            debugLog += `(14時台:${names.length}名抽出)`;
            listA = listA.concat(names);
          } 
          else if (hour === 20 && min >= 0 && min <= 20) {
            const body = msg.getPlainBody();
            const names = extractNamesFromBody(body);
            debugLog += `(20時台:${names.length}名抽出)`;
            listB = listB.concat(names);
          } else {
            debugLog += `(時間外:${hour}時)`;
          }
        } else {
          debugLog += `(時間形式不一致)`;
        }
      }
    }
  }
  return { a: listA, b: listB, debug: debugLog };
}

function extractNamesFromBody(body) {
  const names = [];
  const lines = body.split('\n');
  let inStatusSection = false;
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    if (line.includes('ドライバー別ステータス')) { inStatusSection = true; continue; }
    if (inStatusSection) {
      if (line.includes('※')) { inStatusSection = false; break; }
      const parts = line.split('/');
      if (parts.length >= 4) {
        const name = parts[3].trim();
        if (name) { names.push(name.replace(/\s+/g, '')); }
      }
    }
  }
  return names;
}

// ▼ ダッシュボード連携用API
function doGet(e) {
  var targetMonth = '';
  if (e.parameter && e.parameter.month) {
    targetMonth = e.parameter.month;
  } else {
    var date = new Date();
    targetMonth = date.getFullYear() + '年' + (date.getMonth() + 1) + '月';
  }

  var spreadsheet = SpreadsheetApp.openById("1j0w86gZU_sCvVVcQVQ8R6rzDRaBieTfCnZqPkboREN0");
  var sheet = spreadsheet.getSheetByName(targetMonth);
  
  if (!sheet) {
    return ContentService.createTextOutput(JSON.stringify({ error: targetMonth + " のシートが見つかりません" })).setMimeType(ContentService.MimeType.JSON);
  }

  var data = sheet.getDataRange().getValues();
  
  // 列名（ヘッダー）は1行目
  var headers = data[0]; 
  var result = [];
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      if (headers[j]) {
        obj[headers[j]] = row[j];
      }
    }
    if (obj['日付'] || obj['コード名']) {
      result.push(obj);
    }
  }
  
  var responseData = {
    sheetName: targetMonth,
    data: result
  };
  
  return ContentService.createTextOutput(JSON.stringify(responseData))
    .setMimeType(ContentService.MimeType.JSON);
}

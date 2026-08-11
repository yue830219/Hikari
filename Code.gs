const SHEET_CONFIG = {
  'clients': {
    name: '客戶資料',
    headers: ['ID', '姓名', '性別', '是否會員', '儲值餘額', '贈送課程數', '生日', '職業', '手機', '地址', '過敏史', '肌膚狀況', '生活習慣'],
    keys: ['id', 'name', 'gender', 'isMember', 'balance', 'freeClasses', 'birthday', 'occupation', 'phone', 'address', 'allergy', 'skinConditions', 'lifestyle']
  },
  'products': {
    name: '產品資料',
    headers: ['ID', '代號', '類別', '產品名稱', '定價', '會員價', '非會員加購價', '會員加購價'],
    keys: ['id', 'code', 'category', 'name', 'price', 'memberPrice', 'nonMemberAddonPrice', 'memberAddonPrice']
  },
  'sales': {
    name: '銷售紀錄',
    headers: ['ID', '日期', '客戶名稱', '是否會員', '類型', '銷售項目明細', '額外折扣', '總計金額', '儲值金折抵', '實收金額', '付款方式', '手續費率', '手續費', '付款明細(JSON)', '扣除免費課程', '備註'],
    keys: ['id', 'date', 'clientName', 'isMember', 'type', 'items', 'discount', 'total', 'paidByBalance', 'paidByActual', 'paymentMethod', 'paymentFeeRate', 'paymentFee', 'paymentBreakdown', 'freeClassesDeducted', 'note']
  },
  'expenses': {
    name: '支出紀錄',
    headers: ['ID', '日期', '支出類別', '金額', '備註'],
    keys: ['id', 'date', 'category', 'amount', 'note']
  },
  'users': {
    name: '系統帳號',
    headers: ['ID', '帳號', '密碼', '權限', '名稱', '狀態'],
    keys: ['id', 'username', 'password', 'role', 'displayname', 'status']
  }
};

// 1. 初始化檢查：如果分頁不存在則建立，並寫入中文表頭
function initSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  for (let key in SHEET_CONFIG) {
    let conf = SHEET_CONFIG[key];
    let sheet = ss.getSheetByName(conf.name);
    if (!sheet) {
      sheet = ss.insertSheet(conf.name);
      sheet.appendRow(conf.headers);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, conf.headers.length).setFontWeight("bold").setBackground("#EBE3DB");
    } else {
      // 如果已存在分頁但表頭不符，更新表頭以確保包含必要欄位（例如 '類型' / type）
      try {
        const existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        // 若現有表頭數量或內容與預期不同，覆寫表頭（保守處理：只覆寫第一列）
        if (existingHeaders.length !== conf.headers.length || existingHeaders.join('|') !== conf.headers.join('|')) {
          sheet.getRange(1, 1, 1, conf.headers.length).clearContent();
          sheet.getRange(1, 1, 1, conf.headers.length).setValues([conf.headers]);
          sheet.setFrozenRows(1);
          sheet.getRange(1, 1, 1, conf.headers.length).setFontWeight("bold").setBackground("#EBE3DB");
        }
      } catch (err) {
        // 若檢查失敗，跳過以避免阻斷啟動
      }
    }
  }
}

// 2. 處理讀取請求 (網頁 onload 時觸發)
function doGet(e) {
  initSheets(); 
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let responseData = { status: "success", clients: [], products: [], sales: [], expenses: [], users: [] };  
  for (let key in SHEET_CONFIG) {
    let conf = SHEET_CONFIG[key];
    let sheet = ss.getSheetByName(conf.name);
    
    // 同時抓取「底層原始資料」與「表層顯示文字」
    let dataRange = sheet.getDataRange();
    let data = dataRange.getValues(); 
    let displayData = dataRange.getDisplayValues(); 
    
    if (data.length > 1) {
      for (let i = 1; i < data.length; i++) {
        let row = data[i];
        let displayRow = displayData[i]; 
        let obj = {};
        
        conf.keys.forEach((k, index) => {
          let val = row[index];

          // 處理日期：直接用肉眼所見純文字
          if (k === 'birthday' || k === 'date') {
            val = displayRow[index];
          }
          else if (val instanceof Date) {
            val = Utilities.formatDate(val, "Asia/Taipei", "yyyy-MM-dd");
          }

          // 處理手機號碼開頭 0 被吃掉的問題
          if (k === 'phone' && val !== '' && val !== null) {
            let phoneStr = String(val);
            if (phoneStr.length === 9 && phoneStr.startsWith('9')) {
              val = '0' + phoneStr;
            } else {
              val = phoneStr;
            }
          }

          // 💡【解析銷售明細】兼容舊版 JSON 與新版「頓號」排版文字
          if (k === 'items') {
            if (typeof val === 'string' && val.trim() !== '') {
              if (val.startsWith('[')) { 
                // 讀取舊的 JSON 資料
                try { val = JSON.parse(val); } catch(e) { val = []; }
              } else {
                // 解析新版的人類可讀格式 (毛孔淨化 x1 ($560)、水潤養膚 x1 ($280))
                let parsedItems = [];
                let parts = val.split('、'); // 改用頓號分割
                parts.forEach(part => {
                  let match = part.match(/(.+?) x(\d+) \(\$(\d+)\)/);
                  if (match) {
                    parsedItems.push({
                      name: match[1].trim(),
                      qty: parseInt(match[2]),
                      price: parseInt(match[3])
                    });
                  }
                });
                val = parsedItems;
              }
            } else {
              val = [];
            }
          }

          // 💡【解析付款明細】兼容舊版 JSON 與新版「頓號」排版文字
          if (k === 'paymentBreakdown') {
            if (typeof val === 'string' && val.trim() !== '') {
              if (val.startsWith('{')) {
                // 讀取舊的 JSON 資料
                try { val = JSON.parse(val); } catch(e) { val = {cash:0, transfer:0, line:0}; }
              } else {
                // 解析新版的人類可讀格式
                let pb = {cash: 0, transfer: 0, line: 0};
                let parts = val.split('、'); // 改用頓號分割
                parts.forEach(part => {
                  if (part.includes('現金')) pb.cash = parseInt(part.replace(/[^\d]/g, '')) || 0;
                  if (part.includes('匯款')) pb.transfer = parseInt(part.replace(/[^\d]/g, '')) || 0;
                  // 接受多種可能的 Line Pay 標示：'LinePay', 'Line Pay', 'LinePay:' 等
                  if (part.toLowerCase().includes('line')) pb.line = parseInt(part.replace(/[^\d]/g, '')) || 0;
                });
                val = pb;
              }
            } else {
              val = {cash: 0, transfer: 0, line: 0};
            }
          }
          
          // 針對 肌膚狀況 和 生活習慣
          if (['skinConditions', 'lifestyle'].includes(k)) {
            if (typeof val === 'string' && val.trim() !== '') {
              if (val.startsWith('[')) {
                try { val = JSON.parse(val); } catch(e) { val = val.split('、'); }
              } else {
                val = val.split('、');
              }
            } else {
              val = [];
            }
          }

          if (k === 'isMember') val = (val === true || val === 'TRUE' || val === 'true');
          obj[k] = val;
        });
        responseData[key].push(obj);
      }
    }
  }
  
  return ContentService.createTextOutput(JSON.stringify(responseData)).setMimeType(ContentService.MimeType.JSON);
}

// 3. 處理寫入/更新/刪除請求
function doPost(e) {
  initSheets();
  try {
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action; 
    const dataType = payload.dataType; 
    const data = payload.data; 
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const conf = SHEET_CONFIG[dataType];
    if (!conf) throw new Error("無效的資料類型");
    const sheet = ss.getSheetByName(conf.name);
    
    if (action === 'reorder') {
      if (!Array.isArray(data)) {
        throw new Error("排序資料格式錯誤，請傳入完整資料陣列");
      }
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        sheet.getRange(2, 1, lastRow - 1, conf.keys.length).clearContent();
      }
      const newRows = data.map(item => formatDataToRow(item, conf.keys));
      if (newRows.length > 0) {
        sheet.getRange(2, 1, newRows.length, conf.keys.length).setValues(newRows);
      }
      return ContentService.createTextOutput(JSON.stringify({ status: "success" })).setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'insert') {
      sheet.appendRow(formatDataToRow(data, conf.keys));
    } 
    else if (action === 'update' || action === 'delete') {
      let allData = sheet.getDataRange().getValues();
      let targetId = data.id;
      for (let i = 1; i < allData.length; i++) {
        if (allData[i][0] === targetId) { 
          if (action === 'delete') {
            sheet.deleteRow(i + 1);
          } else {
            sheet.getRange(i + 1, 1, 1, conf.keys.length).setValues([formatDataToRow(data, conf.keys)]);
          }
          break;
        }
      }
    }
    return ContentService.createTextOutput(JSON.stringify({ status: "success" })).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.message })).setMimeType(ContentService.MimeType.JSON);
  }
}

// 將前端 JSON 轉為對應的 Google Sheet 陣列，並處理特定格式
function formatDataToRow(data, keys) {
  return keys.map(k => {
    let val = data[k];
    
    // 💡【寫入銷售明細】轉為「頓號」排版的純文字
    if (k === 'items') {
       if (Array.isArray(val) && val.length > 0) {
           return val.map(i => `${i.name} x${i.qty} ($${i.price})`).join('、');
       }
       return '';
    }
    
    // 💡【寫入付款明細】轉為「頓號」排版的純文字
    if (k === 'paymentBreakdown') {
       if (val && typeof val === 'object') {
         return `現金: $${val.cash || 0}、匯款: $${val.transfer || 0}、Line Pay: $${val.line || 0}`;
       }
       return '';
    }
    
    if (['skinConditions', 'lifestyle'].includes(k)) {
       if (Array.isArray(val)) {
           return val.join('、');
       }
       return val || '';
    }
    
    if (val === undefined || val === null) return '';
    return val;
  });
}
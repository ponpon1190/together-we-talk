/**
 * The Hope | Together We Talk 一起聽見 ‧ 搬家神隊友 數位發聲行動
 * 後端 Google Apps Script (GAS) 核心資料庫管線
 * 
 * 核心特色：
 * 1. 三層資料分流與最小權限架構：
 *    - 分頁 1: 【所有問卷原始紀錄】(管理者總表，含審核狀態)
 *    - 分頁 2: 【同行包_FAQ題庫】(去識別化專案籌備層，僅保留 Q2 標籤與 Q3 敲碗內容，無姓名無隱私)
 *    - 分頁 3: 【教牧代禱_保密專區】(核心教牧層，僅同步姓名、角色、Q5 代禱事項，絕對不公開)
 * 2. 安全性保護：
 *    - doGet() 僅回傳分頁 1 內審核狀態為「已核准」的公開展示資料 (Q1 暱稱/角色、Q2 標籤、Q4 屬靈看見)
 *    - doGet() 程式碼層級嚴格排除 Q3 實務提問與 Q5 守望代禱，杜絕前台滲漏！
 * 3. 跨網域 CORS 支援 (無縫對接 Netlify 與前端靜態網頁)
 */

const SHEET_NAMES = {
  ALL_RECORDS: '所有問卷原始紀錄',
  FAQ_BANK: '同行包_FAQ題庫',
  PASTORAL_PRAYER: '教牧代禱_保密專區'
};

/**
 * 處理前端 POST 提交問卷
 */
function doPost(e) {
  try {
    let data;
    if (e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    } else if (e.parameter) {
      data = e.parameter;
    } else {
      throw new Error('未接收到任何有效資料');
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    initSheetsIfNotExist(ss);

    const timestamp = new Date();
    const name = (data.name || '匿名家人').trim().substring(0, 10);
    const roleId = data.roleId || 'congregation';
    const roleName = data.roleName || getRoleNameById(roleId);
    const topicPrimary = data.topicPrimary || '';
    const topicSecondary = data.topicSecondary || '';
    const q3Question = (data.q3Question || '').trim().substring(0, 100);
    const q4Faith = (data.q4Faith || '').trim().substring(0, 50);
    const q5Prayer = (data.q5Prayer || '').trim().substring(0, 200);
    const wantsReply = data.wantsReply === '是' ? '是' : '否';
    const email = wantsReply === '是' ? (data.email || '').trim().substring(0, 100) : '';
    const status = '已核准'; // 預設核准，同工可手動改為「隱藏」

    // 1. 寫入【所有問卷原始紀錄】
    const sheetAll = ss.getSheetByName(SHEET_NAMES.ALL_RECORDS);
    sheetAll.appendRow([
      timestamp,
      name,
      roleId,
      roleName,
      topicPrimary,
      topicSecondary,
      q3Question,
      q4Faith,
      q5Prayer,
      status,
      wantsReply,
      email
    ]);

    // 2. 寫入【同行包_FAQ題庫】(去識別化：不含姓名、無 Q5 代禱)
    if (q3Question || topicSecondary) {
      const sheetFaq = ss.getSheetByName(SHEET_NAMES.FAQ_BANK);
      sheetFaq.appendRow([
        timestamp,
        topicPrimary,
        topicSecondary,
        q3Question || '(僅選擇關注議題，無額外提問)',
        '待分類解答' // 處理狀態
      ]);
    }

    // 3. 寫入【教牧代禱_保密專區】(僅限牧者查看，絕對不公開)
    if (q5Prayer) {
      const sheetPrayer = ss.getSheetByName(SHEET_NAMES.PASTORAL_PRAYER);
      sheetPrayer.appendRow([
        timestamp,
        name,
        roleName,
        q5Prayer,
        '待守望', // 牧養跟進狀態
        wantsReply,
        email
      ]);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ success: true, message: '問卷送出成功' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * 處理前端 GET 請求 (拉取公開展示牆之審核通過名單)
 * 嚴格遵循最小權限：僅回傳姓名、角色、議題標籤與 Q4 屬靈看見
 */
function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    initSheetsIfNotExist(ss);
    const sheetAll = ss.getSheetByName(SHEET_NAMES.ALL_RECORDS);
    const data = sheetAll.getDataRange().getValues();

    if (data.length <= 1) {
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, data: [] }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 標題列在第 0 列，內容從第 1 列開始
    // 欄位索引：0:時間, 1:姓名, 2:roleId, 3:roleName, 4:主議題, 5:子選項, 6:Q3敲碗, 7:Q4看見, 8:Q5代禱, 9:狀態
    const publicList = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const status = row[9] || '已核准';
      
      if (status === '已核准') {
        publicList.push({
          name: row[1] || '匿名家人',
          roleId: row[2] || 'congregation',
          roleName: row[3] || '會眾',
          topicPrimary: row[4] || '',
          topicSecondary: row[5] || '',
          q4Faith: row[7] || '',
          timestamp: row[0]
        });
      }
    }

    // 倒序排列，讓最新發言在前
    publicList.reverse();

    return ContentService
      .createTextOutput(JSON.stringify({ success: true, data: publicList }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * 自動建立並格式化三張標準分頁
 */
function initSheetsIfNotExist(ss) {
  // 1. 所有問卷原始紀錄
  let s1 = ss.getSheetByName(SHEET_NAMES.ALL_RECORDS);
  if (!s1) {
    s1 = ss.insertSheet(SHEET_NAMES.ALL_RECORDS);
    const headers1 = [
      '送出時間', '填答者姓名', '角色代碼', '角色稱謂',
      '核心議題', '具體子選項', 'Q3 實務敲碗(50字)',
      'Q4 屬靈看見(20字)', 'Q5 守望代禱(保密)', '審核狀態',
      '是否願意收到回覆', 'Email'
    ];
    s1.appendRow(headers1);
    s1.getRange(1, 1, 1, headers1.length)
      .setBackground('#23442e')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    s1.setFrozenRows(1);
  }

  // 2. 同行包_FAQ題庫
  let s2 = ss.getSheetByName(SHEET_NAMES.FAQ_BANK);
  if (!s2) {
    s2 = ss.insertSheet(SHEET_NAMES.FAQ_BANK);
    const headers2 = ['提問時間', '核心議題', '具體關注點', '會眾實務提問內容', 'FAQ解答與進度'];
    s2.appendRow(headers2);
    s2.getRange(1, 1, 1, headers2.length)
      .setBackground('#d88126')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    s2.setFrozenRows(1);
  }

  // 3. 教牧代禱_保密專區
  let s3 = ss.getSheetByName(SHEET_NAMES.PASTORAL_PRAYER);
  if (!s3) {
    s3 = ss.insertSheet(SHEET_NAMES.PASTORAL_PRAYER);
    const headers3 = ['填答時間', '填答者姓名', '角色身分', '守望代禱內容(絕對保密)', '牧養關懷狀態', '是否願意收到回覆', 'Email'];
    s3.appendRow(headers3);
    s3.getRange(1, 1, 1, headers3.length)
      .setBackground('#14281c')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    s3.setFrozenRows(1);
  }
}

function getRoleNameById(roleId) {
  const map = {
    'leader': '小組長',
    'serving_lead': '服事領袖',
    'member': '小組員',
    'volunteer': '服事夥伴',
    'congregation': '會眾'
  };
  return map[roleId] || '會眾';
}

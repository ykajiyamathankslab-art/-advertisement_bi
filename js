/**
 * ============================================================
 * 熊野町 ふるさと納税 統合広告運用レポート v2.0
 * - 履歴保存機能（UPSERT方式）
 * - 前月比・前年比・12ヶ月トレンド対応
 * ============================================================
 */

// ★★★ 設定: 履歴保存用スプレッドシートID ★★★
// 自治体ごとに専用のスプレッドシートを作成し、そのIDをここに設定
const HISTORY_SPREADSHEET_ID = 'スプレッドシートID';  // ← 後で設定

/**
 * ★ デバッグ用: スプレッドシート接続テスト ★
 * Apps Scriptエディタで この関数を直接実行してください
 */
function testSpreadsheetConnection() {
  Logger.log("=== スプレッドシート接続テスト開始 ===");
  Logger.log("設定されたID: " + HISTORY_SPREADSHEET_ID);
  
  if (!HISTORY_SPREADSHEET_ID || HISTORY_SPREADSHEET_ID === 'YOUR_SPREADSHEET_ID_HERE') {
    Logger.log("エラー: IDが未設定です");
    return;
  }
  
  try {
    const ss = SpreadsheetApp.openById(HISTORY_SPREADSHEET_ID);
    Logger.log("成功: スプレッドシートに接続できました");
    Logger.log("スプレッドシート名: " + ss.getName());
    Logger.log("シート一覧: " + ss.getSheets().map(s => s.getName()).join(", "));
  } catch (e) {
    Logger.log("エラー: " + e.toString());
    Logger.log("エラーメッセージ: " + e.message);
  }
  
  Logger.log("=== テスト終了 ===");
}

/**
 * 1. 画面表示用関数
 */
function doGet() {
  return HtmlService.createTemplateFromFile('index').evaluate()
    .setTitle('熊野町 ふるさと納税 統合広告運用レポート')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * 2. クライアント(HTML)から呼び出されるメイン処理関数
 */
function processExcelReport(base64Data, fileName) {
  let fileId = null;
  try {
    const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), MimeType.MICROSOFT_EXCEL, fileName);
    
    const resource = {
      title: "Temp_Analysis_" + new Date().getTime(),
      mimeType: MimeType.GOOGLE_SHEETS
    };
    const newFile = Drive.Files.insert(resource, blob, {convert: true});
    fileId = newFile.id;
    
    const ss = SpreadsheetApp.openById(fileId);
    const data = analyzeData(ss);
    
    // 履歴データも取得して返す（エラーでもメイン処理は継続）
    try {
      const historyData = getHistoryData();
      data.historyData = historyData;
    } catch (historyError) {
      Logger.log("履歴取得エラー（無視して続行）: " + historyError.toString());
      data.historyData = { exists: false, message: '履歴取得でエラーが発生しました', records: [] };
    }
    
    return data;

  } catch (e) {
    Logger.log("Error: " + e.toString());
    throw new Error("処理エラー: " + e.message);
  } finally {
    if (fileId) {
      try { DriveApp.getFileById(fileId).setTrashed(true); } catch(e) {}
    }
  }
}

/**
 * 3. 履歴に保存（UPSERT方式）
 */
function saveToHistory(yearMonth, summaryData) {
  try {
    // ID未設定チェック
    if (!HISTORY_SPREADSHEET_ID || HISTORY_SPREADSHEET_ID === 'YOUR_SPREADSHEET_ID_HERE') {
      return { success: false, message: '履歴スプレッドシートIDが未設定です。Code.gsの HISTORY_SPREADSHEET_ID を設定してください。' };
    }
    
    let ss;
    try {
      ss = SpreadsheetApp.openById(HISTORY_SPREADSHEET_ID);
    } catch (accessError) {
      return { success: false, message: 'スプレッドシートにアクセスできません。共有設定を確認してください。' };
    }
    
    let sheet = ss.getSheetByName('月次履歴');
    
    // シートがなければ作成
    if (!sheet) {
      sheet = ss.insertSheet('月次履歴');
      // ヘッダー行を追加
      const headers = [
        '年月', '更新日時',
        '楽天_予算', '楽天_消化', '楽天_売上', '楽天_CV', '楽天_ROAS', '楽天_CPA', '楽天_クリック', '楽天_CTR', '楽天_CVR',
        'チョイス_予算', 'チョイス_消化', 'チョイス_売上', 'チョイス_CV', 'チョイス_ROAS', 'チョイス_CPA', 'チョイス_クリック', 'チョイス_imp', 'チョイス_CTR', 'チョイス_CVR',
        '合計_予算', '合計_消化', '合計_売上', '合計_CV', '合計_ROAS', '合計_CPA'
      ];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
    
    // 既存データを検索（日付型も考慮）
    const data = sheet.getDataRange().getValues();
    let existingRow = -1;
    for (let i = 1; i < data.length; i++) {
      let cellValue = data[i][0];
      // 日付オブジェクトの場合は文字列に変換
      if (cellValue instanceof Date) {
        cellValue = Utilities.formatDate(cellValue, Session.getScriptTimeZone(), "yyyy/MM");
      } else {
        cellValue = String(cellValue || '');
      }
      if (cellValue === yearMonth) {
        existingRow = i + 1; // 1-indexed
        break;
      }
    }
    
    // 保存データを構築
    const r = summaryData.rakuten;
    const c = summaryData.choice;
    const t = summaryData.total;
    const now = new Date();
    
    const rowData = [
      yearMonth,
      Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss"),
      r.budget, r.cost, r.sales, r.cv, r.roas, r.cpa, r.clicks, r.ctr, r.cvr,
      c.budget, c.cost, c.sales, c.cv, c.roas, c.cpa, c.clicks, c.imp, c.ctr, c.cvr,
      t.budget, t.cost, t.sales, t.cv, t.roas, t.cpa
    ];
    
    if (existingRow > 0) {
      // 上書き
      sheet.getRange(existingRow, 1, 1, rowData.length).setValues([rowData]);
      return { success: true, message: `${yearMonth} のデータを更新しました`, mode: 'update' };
    } else {
      // 新規追加
      sheet.appendRow(rowData);
      // 年月でソート（降順）
      const lastRow = sheet.getLastRow();
      if (lastRow > 2) {
        sheet.getRange(2, 1, lastRow - 1, rowData.length).sort({column: 1, ascending: false});
      }
      return { success: true, message: `${yearMonth} のデータを新規保存しました`, mode: 'insert' };
    }
    
  } catch (e) {
    Logger.log("saveToHistory Error: " + e.toString());
    return { success: false, message: "保存エラー: " + e.message };
  }
}

/**
 * 4. 履歴データを取得
 */
function getHistoryData() {
  try {
    // IDが未設定の場合
    if (!HISTORY_SPREADSHEET_ID || HISTORY_SPREADSHEET_ID === 'YOUR_SPREADSHEET_ID_HERE') {
      return { exists: false, message: '履歴スプレッドシートが未設定です', records: [] };
    }
    
    let ss;
    try {
      ss = SpreadsheetApp.openById(HISTORY_SPREADSHEET_ID);
    } catch (accessError) {
      // アクセス権限エラーの場合
      Logger.log("スプレッドシートアクセスエラー: " + accessError.toString());
      return { exists: false, message: 'スプレッドシートにアクセスできません。共有設定を確認してください。', records: [] };
    }
    
    const sheet = ss.getSheetByName('月次履歴');
    
    if (!sheet || sheet.getLastRow() < 2) {
      return { exists: true, message: '履歴データがありません', records: [] };
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const records = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;
      
      // 年月を文字列に変換（日付オブジェクトの場合も対応）
      let yearMonth = row[0];
      if (yearMonth instanceof Date) {
        yearMonth = Utilities.formatDate(yearMonth, Session.getScriptTimeZone(), "yyyy/MM");
      } else {
        yearMonth = String(yearMonth);
      }
      
      records.push({
        yearMonth: yearMonth,
        updatedAt: row[1] instanceof Date ? Utilities.formatDate(row[1], Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss") : String(row[1] || ''),
        rakuten: {
          budget: Number(row[2]) || 0, cost: Number(row[3]) || 0, sales: Number(row[4]) || 0, cv: Number(row[5]) || 0,
          roas: Number(row[6]) || 0, cpa: Number(row[7]) || 0, clicks: Number(row[8]) || 0, ctr: Number(row[9]) || 0, cvr: Number(row[10]) || 0
        },
        choice: {
          budget: Number(row[11]) || 0, cost: Number(row[12]) || 0, sales: Number(row[13]) || 0, cv: Number(row[14]) || 0,
          roas: Number(row[15]) || 0, cpa: Number(row[16]) || 0, clicks: Number(row[17]) || 0, imp: Number(row[18]) || 0, ctr: Number(row[19]) || 0, cvr: Number(row[20]) || 0
        },
        total: {
          budget: Number(row[21]) || 0, cost: Number(row[22]) || 0, sales: Number(row[23]) || 0, cv: Number(row[24]) || 0,
          roas: Number(row[25]) || 0, cpa: Number(row[26]) || 0
        }
      });
    }
    
    // 年月順にソート（降順）- 文字列として比較
    records.sort((a, b) => {
      const aStr = String(a.yearMonth || '');
      const bStr = String(b.yearMonth || '');
      return bStr.localeCompare(aStr);
    });
    
    return { exists: true, records: records };
    
  } catch (e) {
    Logger.log("getHistoryData Error: " + e.toString());
    return { exists: false, message: "履歴取得エラー: " + e.message, records: [] };
  }
}

/**
 * 5. 比較データを計算（前月比・前年比）
 */
function getComparisonData(currentYearMonth) {
  const history = getHistoryData();
  if (!history.exists || history.records.length === 0) {
    return { hasPrevMonth: false, hasPrevYear: false };
  }
  
  // 年月をパース
  const [year, month] = currentYearMonth.split('/').map(Number);
  
  // 前月
  let prevMonth = month - 1;
  let prevMonthYear = year;
  if (prevMonth === 0) {
    prevMonth = 12;
    prevMonthYear = year - 1;
  }
  const prevMonthKey = `${prevMonthYear}/${String(prevMonth).padStart(2, '0')}`;
  
  // 前年同月
  const prevYearKey = `${year - 1}/${String(month).padStart(2, '0')}`;
  
  const result = {
    hasPrevMonth: false,
    hasPrevYear: false,
    prevMonth: null,
    prevYear: null
  };
  
  for (const record of history.records) {
    if (record.yearMonth === prevMonthKey) {
      result.hasPrevMonth = true;
      result.prevMonth = record;
    }
    if (record.yearMonth === prevYearKey) {
      result.hasPrevYear = true;
      result.prevYear = record;
    }
  }
  
  return result;
}

/**
 * 6. データ分析・抽出ロジック（既存のanalyzeData関数）
 */
function analyzeData(ss) {
  
  // --- ヘルパー関数 ---
  const parseNumeric = (val) => {
    if (typeof val === 'number') return val;
    if (!val) return 0;
    const cleanStr = String(val)
      .replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0))
      .replace(/,/g, '')
      .replace(/円/g, '')
      .replace(/%/g, '')
      .trim();
    const num = parseFloat(cleanStr);
    return isNaN(num) ? 0 : num;
  };

  const formatDate = (date) => {
    if (date instanceof Date) {
      return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy/MM/dd");
    }
    return String(date);
  };

  const cleanRakutenName = (name) => (!name) ? "" : String(name).replace(/^【ふるさと納税】/, '').trim();
  const cleanChoiceName = (name) => (!name) ? "" : String(name).replace(/^熊野化粧筆/, '').trim();

  const getRawData = (sheetName) => {
    const sheet = ss.getSheetByName(sheetName);
    return sheet ? sheet.getDataRange().getValues() : [];
  };

  const getTableData = (sheetName, headerKeyword) => {
    const raw = getRawData(sheetName);
    let headerRowIndex = -1;
    for (let r = 0; r < raw.length; r++) {
      if (raw[r].some(cell => String(cell).includes(headerKeyword))) {
        headerRowIndex = r;
        break;
      }
    }
    if (headerRowIndex === -1) return [];

    const headers = raw[headerRowIndex].map(h => String(h).trim());
    const dataRows = raw.slice(headerRowIndex + 1);

    return dataRows.map(row => {
      let obj = {};
      headers.forEach((h, i) => {
        if (h) obj[h] = row[i];
      });
      return obj;
    }).filter(obj => Object.keys(obj).length > 0);
  };

  const findValueInRow = (rawMatrix, keyword) => {
    for (let r = 0; r < rawMatrix.length; r++) {
      for (let c = 0; c < rawMatrix[r].length; c++) {
        if (String(rawMatrix[r][c]).includes(keyword)) {
          for (let nextC = c + 1; nextC < rawMatrix[r].length; nextC++) {
            const val = rawMatrix[r][nextC];
            if (val !== "" && val !== null) return parseNumeric(val);
          }
        }
      }
    }
    return 0;
  };

  // --- 期間特定 ---
  const dailyRPPData = getTableData('日次データ_全体', '日付');
  let targetYearMonth = "";
  let periodDisplay = "対象期間不明";
  if (dailyRPPData.length > 0) {
    const firstDateStr = formatDate(dailyRPPData[0]['日付']);
    const dateObj = new Date(firstDateStr);
    if (!isNaN(dateObj.getTime())) {
      targetYearMonth = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "yyyy/MM"); 
      periodDisplay = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "yyyy年MM月");
    }
  }

  // --- 処理実行 ---
  const rakutenData = processRakutenData(ss, dailyRPPData, targetYearMonth);
  const choiceData = processChoiceData(ss, targetYearMonth);

  // --- 全体統合 ---
  const totalData = {
    budget: rakutenData.budgetAnalysis.budget + choiceData.budgetAnalysis.budget,
    cost: rakutenData.budgetAnalysis.actualSpend + choiceData.budgetAnalysis.actualSpend,
    sales: rakutenData.budgetAnalysis.actualSales + choiceData.budgetAnalysis.actualSales,
    cv: rakutenData.rppOverallContribution.totalRppCv + choiceData.metrics.totalCv,
    clicks: rakutenData.monthlySummaryDetails.clicks + choiceData.metrics.totalClicks,
  };
  totalData.roas = totalData.cost > 0 ? (totalData.sales / totalData.cost) * 100 : 0;
  totalData.cpa = totalData.cv > 0 ? (totalData.cost / totalData.cv) : 0;
  totalData.cpc = totalData.clicks > 0 ? (totalData.cost / totalData.clicks) : 0;

  return {
    meta: {
      municipality: rakutenData.meta.municipality || '熊野町',
      period: periodDisplay,
      yearMonth: targetYearMonth  // 保存用キー
    },
    totalData,
    rakutenData,
    choiceData
  };

  // ---------------------------------------------------------
  //  楽天データ処理
  // ---------------------------------------------------------
  function processRakutenData(ss, dailyData, targetYm) {
    const budgetRaw = getRawData('設定_予算');
    const summaryRaw = getRawData('月次サマリー');
    const budgetTableData = getTableData('設定_予算', '年月');
    const storeMonthlyData = getTableData('店舗月次集計', '日付');
    const rppList = getTableData('RPP出稿リスト', '商品管理番号');
    const optimizationReport = getTableData('全体最適レポート', '商品管理番号');
    const keywordSettings = getTableData('キーワード設定', '商品管理番号');
    const keywordRanking = getTableData('過去1年間流入キーワードランキング', '順位');
    const pastYearSales = getTableData('過去1年間売り上げ', '商品名') || getTableData('過去1年間売り上げ', '商品管理番号');
    const workHistory = getTableData('作業履歴', '日付');

    let budgetAmount = 0;
    let targetRoas = 2000;
    const targetBudgetRow = budgetTableData.find(row => {
      const d = row['年月'] instanceof Date ? row['年月'] : new Date(String(row['年月']).replace('年','/').replace('月','/1'));
      return !isNaN(d) && Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy/MM") === targetYm;
    });
    if (targetBudgetRow) {
      budgetAmount = parseNumeric(targetBudgetRow['予算(税抜)']) || parseNumeric(targetBudgetRow['予算(税込)']);
      targetRoas = parseNumeric(targetBudgetRow['目標ROAS(%)']);
    }

    const md = {
      ctr: findValueInRow(summaryRaw, 'CTR平均'),
      clicks: findValueInRow(summaryRaw, 'クリック数'),
      cost: findValueInRow(summaryRaw, '実績額'), 
      salesTotal: findValueInRow(summaryRaw, '売上金額(合計'),
      cvTotal: findValueInRow(summaryRaw, '売上件数(合計'),
      cvrTotal: findValueInRow(summaryRaw, 'CVR(合計'),
      roasTotal: findValueInRow(summaryRaw, 'ROAS(合計'),
      salesNew: findValueInRow(summaryRaw, '売上金額(新規'),
      cvNew: findValueInRow(summaryRaw, '売上件数(新規'),
      salesExisting: findValueInRow(summaryRaw, '売上金額(既存'),
      cvExisting: findValueInRow(summaryRaw, '売上件数(既存')
    };
    md.cpc = md.clicks > 0 ? Math.round(md.cost / md.clicks) : 0;

    const contAnalysis = [];
    let totalStoreSales = 0;
    let totalRPPCv = 0;
    const rppMap = {};
    dailyData.forEach(row => {
      rppMap[formatDate(row['日付'])] = {
        sales: parseNumeric(row['売上金額(合計720時間)']),
        cost: parseNumeric(row['実績額(合計)']),
        cv: parseNumeric(row['売上件数(合計720時間)']),
        cvr: parseNumeric(row['CVR(合計720時間)(%)'])
      };
    });
    storeMonthlyData.forEach(row => {
      const date = formatDate(row['日付']);
      const sSales = parseNumeric(row['売上'] || row['売上金額']);
      totalStoreSales += sSales;
      const rData = rppMap[date] || { sales: 0, cost: 0, cv: 0, cvr: 0 };
      totalRPPCv += rData.cv;
      contAnalysis.push({
        date: date, storeSales: sSales, rppSales: rData.sales, rppCost: rData.cost,
        contributionRatio: sSales > 0 ? (rData.sales / sSales) * 100 : 0,
        rppCv: rData.cv, rppCvr: rData.cvr
      });
    });
    contAnalysis.sort((a,b)=> new Date(a.date)-new Date(b.date));

    const registeredIds = new Set(rppList.map(r => String(r['商品管理番号'])));
    const keywordConfiguredIds = new Set(keywordSettings.map(k => String(k['商品管理番号'])));
    const productNameMap = {};
    pastYearSales.forEach(row => {
      const id = row['商品管理番号'] || row['商品番号'];
      if(id) productNameMap[String(id)] = cleanRakutenName(row['商品名']);
    });
    const rankingList = keywordRanking.map(r => ({ rank: parseNumeric(r['順位']), keyword: String(r['検索キーワード']) })).sort((a,b) => a.rank - b.rank);

    const soldProductsList = [];
    const recommendations = { add: [], warn: [] };
    optimizationReport.forEach(row => {
      const cvCount = parseNumeric(row['売上件数(合計720時間)']);
      const id = String(row['商品管理番号']);
      const name = productNameMap[id] || cleanRakutenName(row['商品名']) || "商品名不明";
      const amount = parseNumeric(row['売上金額(合計720時間)']);
      const roas = parseNumeric(row['ROAS(合計720時間)(%)']);
      const spend = parseNumeric(row['実績額(合計)']);

      if (cvCount > 0) {
        let matchedRank = 99999, matchedKw = "";
        for (const r of rankingList) {
          if (name.includes(r.keyword) || r.keyword.includes(name)) { matchedRank = r.rank; matchedKw = r.keyword; break; }
        }
        soldProductsList.push({
          id, name, count: cvCount, amount, rank: matchedRank === 99999 ? '-' : matchedRank, rankKw: matchedKw,
          isRpp: registeredIds.has(id), isKw: keywordConfiguredIds.has(id)
        });
      }
      if (roas >= 400 && !registeredIds.has(id)) recommendations.add.push({ id, name, roas, spend, sales: cvCount });
      if (spend > 1000 && cvCount === 0) recommendations.warn.push({ id, name, spend, clicks: row['クリック数(合計)'] });
    });
    soldProductsList.sort((a, b) => b.amount - a.amount);

    const missingKeywords = [];
    const currentKeywords = new Set(keywordSettings.map(k => k['キーワード']));
    keywordRanking.slice(0, 20).forEach(row => {
      if (!currentKeywords.has(row['検索キーワード'])) missingKeywords.push({ rank: row['順位'], keyword: row['検索キーワード'], users: parseNumeric(row['アクセス人数']) });
    });

    const activeRppList = rppList.map(row => ({ id: row['商品管理番号'], name: cleanRakutenName(row['商品名']), price: parseNumeric(row['価格']), cpc: parseNumeric(row['商品CPC']) })).filter(item => item.id);
    const activeKeywordList = keywordSettings.map(row => ({ id: row['商品管理番号'], name: cleanRakutenName(row['商品名']), keyword: row['キーワード'], cpc: parseNumeric(row['キーワードCPC'] || row['目安CPC']) })).filter(item => item.keyword);
    const formattedHistory = workHistory.map(row => ({ date: formatDate(row['日付']), type: row['作業種別'], target: row['対象商品'], detail: row['備考'] || `${row['変更前']||''}→${row['変更後']||''}` }));

    return {
      meta: { municipality: ((raw)=>{for(let r=0;r<raw.length;r++){for(let c=0;c<raw[r].length;c++){if(String(raw[r][c]).includes('自治体名')&&raw[r][c+1])return raw[r][c+1]}}} )(budgetRaw) },
      monthlySummaryDetails: md,
      budgetAnalysis: { budget: budgetAmount, targetRoas, actualSpend: md.cost, actualRoas: md.roasTotal, actualSales: md.salesTotal },
      rppOverallContribution: { totalStoreSales, totalRppSales: md.salesTotal, ratio: totalStoreSales > 0 ? (md.salesTotal/totalStoreSales)*100 : 0, totalRppCv: totalRPPCv, totalRppCvr: md.cvrTotal },
      contributionAnalysis: contAnalysis,
      soldProductsList, recommendations, missingKeywords, activeRppList, activeKeywordList, workHistory: formattedHistory
    };
  }

  // ---------------------------------------------------------
  //  チョイスデータ処理
  // ---------------------------------------------------------
  function processChoiceData(ss, targetYm) {
    const budgetTable = getTableData('ふるさとチョイス設定_予算', '年月');
    const dailyData = getTableData('ふるさとチョイス日次データ_全体', '日付');
    const productReport = getTableData('ふるさとチョイス全体最適レポート', '品ID');
    const keywordReport = getTableData('ふるさとチョイス設定キーワード', 'キーワード');
    const workHistory = getTableData('ふるさとチョイス作業履歴', '日付');

    let budgetAmount = 0;
    let targetRoas = 2000;
    const targetBudgetRow = budgetTable.find(row => {
      const d = row['年月'] instanceof Date ? row['年月'] : new Date(String(row['年月']).replace('年','/').replace('月','/1'));
      return !isNaN(d) && Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy/MM") === targetYm;
    });
    if (targetBudgetRow) {
      budgetAmount = parseNumeric(targetBudgetRow['予算(税抜)']) || parseNumeric(targetBudgetRow['予算(税込)']);
      targetRoas = parseNumeric(targetBudgetRow['目標ROAS(%)']);
    }

    let totalCost = 0, totalSales = 0, totalCv = 0, totalClicks = 0, totalImp = 0;
    const chartData = [];

    dailyData.forEach(row => {
      const cost = parseNumeric(row['消化額']);
      const sales = parseNumeric(row['寄付経由額']);
      const cv = parseNumeric(row['寄付件数']);
      const clicks = parseNumeric(row['クリック'] || row['クリック数']);
      const imp = parseNumeric(row['imp'] || row['表示回数']);

      totalCost += cost;
      totalSales += sales;
      totalCv += cv;
      totalClicks += clicks;
      totalImp += imp;

      chartData.push({ date: formatDate(row['日付']), cost, sales, cv, clicks });
    });
    chartData.sort((a,b) => new Date(a.date) - new Date(b.date));

    const monthlySummaryDetails = {
      cost: totalCost, sales: totalSales, cv: totalCv, clicks: totalClicks, imp: totalImp,
      ctr: totalImp > 0 ? (totalClicks / totalImp) * 100 : 0,
      cpc: totalClicks > 0 ? Math.round(totalCost / totalClicks) : 0,
      roas: totalCost > 0 ? (totalSales / totalCost) * 100 : 0,
      cvr: totalClicks > 0 ? (totalCv / totalClicks) * 100 : 0
    };

    const productList = [];
    productReport.forEach(row => {
      const cv = parseNumeric(row['寄付件数']);
      if (cv > 0 || parseNumeric(row['消化額']) > 0) {
        productList.push({
          id: row['品ID'],
          name: cleanChoiceName(row['品名']),
          cv: cv,
          sales: parseNumeric(row['寄付経由額']),
          cost: parseNumeric(row['消化額']),
          roas: parseNumeric(row['ROAS'])
        });
      }
    });
    productList.sort((a,b) => b.sales - a.sales);

    const keywordList = [];
    keywordReport.forEach(row => {
      keywordList.push({
        keyword: row['キーワード'],
        product: cleanChoiceName(row['品名']),
        cv: parseNumeric(row['寄付件数']),
        sales: parseNumeric(row['寄付経由額']),
        cost: parseNumeric(row['消化額'])
      });
    });
    keywordList.sort((a,b) => b.sales - a.sales);

    const formattedHistory = workHistory.map(row => ({
      date: formatDate(row['日付']),
      type: row['作業種別'] || row['作業'] || '',
      target: row['対象'] || row['対象商品'] || '',
      detail: row['詳細'] || row['備考'] || ''
    }));

    return {
      budgetAnalysis: { budget: budgetAmount, targetRoas, actualSpend: totalCost, actualRoas: monthlySummaryDetails.roas, actualSales: totalSales },
      metrics: { totalCv, totalClicks, cpc: monthlySummaryDetails.cpc, cvr: monthlySummaryDetails.cvr },
      monthlySummaryDetails,
      chartData, productList, keywordList,
      workHistory: formattedHistory
    };
  }
}

// ===== 成績レポート取得 GAS（Webアプリとしてデプロイ） =====
// デプロイ設定：「次のユーザーとして実行」= 自分 ／「アクセスできるユーザー」= 全員

const CONFIG = {
  // ① 成績管理表
  SEISEKI_ID: '19oAUC2liN34I5QwQ72l60e1puD-VR547MBLEjVE2O1k',
  SEISEKI_SHEETS: ['アポ', '前確', '商談'],
  SEISEKI_START_COL: 2,       // B列から（A列は含めない）
  MONTH_ROW: 1,               // この行から「2026年9月」形式のセルを自動で探す
  TOTAL_LABEL: '合計',        // B列でこの文字がある行を最終行＆売上合計行とみなす

  // ② 勤怠管理（月ごとに別ファイル。ファイル名で自動検索）
  KINTAI_NAME: (y, m) => y + '.' + ('0' + m).slice(-2) + 'MEOアポ勤怠管理',
  KINTAI_SHEET: '稼働時間',
  KINTAI_SALES_LABEL: '当月売上',   // このセルの右隣を当月売上とみなす
  KINTAI_END_LABEL: '時間単価',     // A列でこの文字がある行までを画像範囲にする
  KINTAI_LAST_COL: 6                // F列まで（A1:F17）
};

function doGet(e) {
  let out;
  try {
    out = buildReport();
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
  }
  return ContentService
    .createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

function buildReport() {
  const now = new Date();
  const y = Number(Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy'));
  const m = Number(Utilities.formatDate(now, 'Asia/Tokyo', 'M'));

  // ---------- ① 成績管理表 ----------
  const ss = SpreadsheetApp.openById(CONFIG.SEISEKI_ID);
  const seiseki = CONFIG.SEISEKI_SHEETS.map(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return { name: name, error: 'タブ「' + name + '」が見つかりません' };

    const lastRow = sh.getLastRow();
    const colB = sh.getRange(1, 2, lastRow, 1).getValues().map(r => String(r[0]).trim());
    let endRow = colB.indexOf(CONFIG.TOTAL_LABEL) + 1;
    if (endRow < 1) endRow = lastRow;

    const header = sh.getRange(4, 1, 1, sh.getLastColumn()).getValues()[0];
    let endCol = 0;
    header.forEach((v, i) => { if (String(v).trim() !== '') endCol = i + 1; });
    if (!endCol) endCol = sh.getLastColumn();

    const sc = CONFIG.SEISEKI_START_COL;
    const block = grabRange(sh, 1, sc, endRow, endCol - sc + 1);
    const month = findMonth(sh, endCol);
    return Object.assign({
      name: name,
      spreadsheetName: ss.getName(),
      url: ss.getUrl() + '#gid=' + sh.getSheetId(),
      monthLabel: month.label, monthCell: month.cell,
      totalSales: Number(sh.getRange(endRow, 3).getValue()) || 0,
      totalCell: 'C' + endRow
    }, block);
  });

  // ---------- ② 勤怠管理 ----------
  const kName = CONFIG.KINTAI_NAME(y, m);
  let kintai;
  const files = DriveApp.getFilesByName(kName);
  if (!files.hasNext()) {
    kintai = { name: CONFIG.KINTAI_SHEET, spreadsheetName: kName,
               error: 'Driveに「' + kName + '」というファイルが見つかりません' };
  } else {
    const file = files.next();
    const kss = SpreadsheetApp.open(file);
    const sh = kss.getSheetByName(CONFIG.KINTAI_SHEET);
    if (!sh) {
      kintai = { name: CONFIG.KINTAI_SHEET, spreadsheetName: kss.getName(), url: kss.getUrl(),
                 error: 'タブ「' + CONFIG.KINTAI_SHEET + '」が見つかりません' };
    } else {
      const salesCell = sh.createTextFinder(CONFIG.KINTAI_SALES_LABEL).matchEntireCell(true).findNext();
      const endCell = sh.createTextFinder(CONFIG.KINTAI_END_LABEL).matchEntireCell(true).findNext();
      let endRow = endCell ? endCell.getRow() : (salesCell ? salesCell.getRow() + 1 : sh.getLastRow());
      const block = grabRange(sh, 1, 1, endRow, CONFIG.KINTAI_LAST_COL);
      kintai = Object.assign({
        name: CONFIG.KINTAI_SHEET,
        spreadsheetName: kss.getName(),
        url: kss.getUrl() + '#gid=' + sh.getSheetId(),
        totalSales: salesCell ? (Number(salesCell.offset(0, 1).getValue()) || 0) : null,
        totalCell: salesCell ? salesCell.offset(0, 1).getA1Notation() : null,
        salesLabelFound: !!salesCell
      }, block);
    }
  }

  return {
    ok: true,
    generatedAt: Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm'),
    thisMonth: { year: y, month: m, label: y + '年' + m + '月' },
    seiseki: seiseki,
    kintai: kintai
  };
}

// 1行目から「YYYY年M月」形式のセルを探す
function findMonth(sh, endCol) {
  const vals = sh.getRange(CONFIG.MONTH_ROW, 1, 1, Math.max(endCol, sh.getLastColumn())).getDisplayValues()[0];
  for (let i = 0; i < vals.length; i++) {
    if (/\d{4}\s*年\s*\d{1,2}\s*月/.test(vals[i])) {
      return { label: vals[i].match(/\d{4}\s*年\s*\d{1,2}\s*月/)[0].replace(/\s/g, ''),
               cell: sh.getRange(CONFIG.MONTH_ROW, i + 1).getA1Notation() };
    }
  }
  return { label: '', cell: '' };
}

// 指定範囲の値と見た目（背景・文字色・太字・サイズ・寄せ・結合・列幅・行高）をまとめて取る
function grabRange(sh, r, c, nr, nc) {
  const range = sh.getRange(r, c, nr, nc);
  const merges = range.getMergedRanges().map(mr => ({
    row: mr.getRow() - r, col: mr.getColumn() - c,
    rows: mr.getNumRows(), cols: mr.getNumColumns()
  }));
  const colWidths = [], rowHeights = [];
  for (let i = 0; i < nc; i++) colWidths.push(sh.getColumnWidth(c + i));
  for (let i = 0; i < nr; i++) rowHeights.push(sh.getRowHeight(r + i));
  return {
    rows: nr, cols: nc,
    display: range.getDisplayValues(),
    backgrounds: range.getBackgrounds(),
    fontColors: range.getFontColors(),
    fontWeights: range.getFontWeights(),
    fontSizes: range.getFontSizes(),
    hAlign: range.getHorizontalAlignments(),
    vAlign: range.getVerticalAlignments(),
    wraps: range.getWrapStrategies().map(r => r.map(w => String(w))),
    merges: merges, colWidths: colWidths, rowHeights: rowHeights
  };
}

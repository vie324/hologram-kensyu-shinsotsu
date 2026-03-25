/**
 * HOLOGRAM マナー研修 — 進捗管理 Google Apps Script
 *
 * 【セットアップ手順】
 * 1. Google Drive で新規スプレッドシートを作成
 * 2. 「拡張機能 > Apps Script」を開く
 * 3. このファイルの内容を貼り付ける
 * 4. 「デプロイ > 新しいデプロイ」→ 種類「ウェブアプリ」を選択
 *    - 実行するユーザー: 自分
 *    - アクセスできるユーザー: 全員
 * 5. デプロイして表示されるURLを、index.html の GAS_URL 定数に貼り付ける
 * 6. スプレッドシートIDを下の SPREADSHEET_ID に設定（URLの /d/XXXXX/edit の XXXXX 部分）
 *
 * 【シート構成】（初回実行時に自動作成されます）
 * - 「進捗ログ」: 全イベントの生データ
 * - 「受講者一覧」: 受講者ごとの最新サマリー
 * - 「管理ダッシュボード」: 管理者向けの集計ビュー
 */

// ── 設定 ──
const SPREADSHEET_ID = ''; // 空の場合はスクリプトにバインドされたスプレッドシートを使用

function getSpreadsheet() {
  if (SPREADSHEET_ID) {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

// ── シート初期化 ──
function initializeSheets() {
  const ss = getSpreadsheet();

  // 進捗ログシート
  let logSheet = ss.getSheetByName('進捗ログ');
  if (!logSheet) {
    logSheet = ss.insertSheet('進捗ログ');
    logSheet.appendRow([
      'タイムスタンプ', '受講者名', 'イベント種別',
      'ステップ', 'モジュールID', 'モジュール名',
      'レッスンID', 'レッスン名',
      'クイズ正答数', 'クイズ問題数', 'クイズ合格',
      '記述回答内容', '記述問題文'
    ]);
    logSheet.getRange(1, 1, 1, 13).setFontWeight('bold').setBackground('#5B0E2D').setFontColor('#FFFFFF');
    logSheet.setFrozenRows(1);
    logSheet.setColumnWidth(1, 160);
    logSheet.setColumnWidth(2, 120);
    logSheet.setColumnWidth(3, 100);
    logSheet.setColumnWidth(12, 400);
    logSheet.setColumnWidth(13, 300);
  }

  // 受講者一覧シート
  let summarySheet = ss.getSheetByName('受講者一覧');
  if (!summarySheet) {
    summarySheet = ss.insertSheet('受講者一覧');
    summarySheet.appendRow([
      '受講者名', '最終アクセス', '完了レッスン数', '全レッスン数',
      '進捗率', 'クイズ合格数', '全クイズ数', 'クイズ合格率',
      '記述回答数', 'STEP1進捗', 'STEP2進捗', 'STEP3進捗'
    ]);
    summarySheet.getRange(1, 1, 1, 12).setFontWeight('bold').setBackground('#5B0E2D').setFontColor('#FFFFFF');
    summarySheet.setFrozenRows(1);
  }

  // 管理ダッシュボードシート
  let dashSheet = ss.getSheetByName('管理ダッシュボード');
  if (!dashSheet) {
    dashSheet = ss.insertSheet('管理ダッシュボード');
    dashSheet.getRange('A1').setValue('HOLOGRAM マナー研修 管理ダッシュボード')
      .setFontSize(16).setFontWeight('bold').setFontColor('#5B0E2D');
    dashSheet.getRange('A2').setValue('最終更新: ' + new Date().toLocaleString('ja-JP'))
      .setFontColor('#8C7B72');
    dashSheet.setColumnWidth(1, 200);
    dashSheet.setColumnWidth(2, 150);
    dashSheet.setColumnWidth(3, 150);
    dashSheet.setColumnWidth(4, 150);
  }

  // デフォルトの Sheet1 を削除（あれば）
  const defaultSheet = ss.getSheetByName('Sheet1') || ss.getSheetByName('シート1');
  if (defaultSheet && ss.getSheets().length > 1) {
    try { ss.deleteSheet(defaultSheet); } catch (e) { /* ignore */ }
  }

  return ss;
}

// ── Web API エンドポイント ──
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;

    let result;
    switch (action) {
      case 'log_progress':
        result = logProgress(data);
        break;
      case 'sync_full':
        result = syncFullProgress(data);
        break;
      case 'get_dashboard':
        result = getDashboardData(data);
        break;
      default:
        result = { error: 'Unknown action: ' + action };
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'get_dashboard';
  const password = (e && e.parameter && e.parameter.password) || '';

  if (action === 'get_dashboard') {
    const result = getDashboardData({ password: password });
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput(JSON.stringify({ status: 'ok', message: 'HOLOGRAM Training API' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── 進捗ログの記録 ──
function logProgress(data) {
  const ss = initializeSheets();
  const logSheet = ss.getSheetByName('進捗ログ');
  const timestamp = new Date().toLocaleString('ja-JP');
  const name = data.traineeName || '未入力';

  const events = data.events || [];
  events.forEach(ev => {
    logSheet.appendRow([
      timestamp,
      name,
      ev.type || '',        // lesson_complete, quiz_submit, written_submit
      ev.dayId || '',
      ev.moduleId || '',
      ev.moduleName || '',
      ev.lessonId || '',
      ev.lessonName || '',
      ev.quizCorrect || '',
      ev.quizTotal || '',
      ev.quizPassed ? '合格' : (ev.type === 'quiz_submit' ? '不合格' : ''),
      ev.writtenAnswer || '',
      ev.writtenQuestion || ''
    ]);
  });

  // 受講者サマリー更新
  updateSummary(ss, name, data);

  return { status: 'ok', logged: events.length };
}

// ── 全進捗の一括同期 ──
function syncFullProgress(data) {
  const ss = initializeSheets();
  const name = data.traineeName || '未入力';

  // サマリー更新
  updateSummary(ss, name, data);

  return { status: 'ok', synced: true };
}

// ── 受講者サマリーの更新 ──
function updateSummary(ss, name, data) {
  const sheet = ss.getSheetByName('受講者一覧');
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  let rowIndex = -1;

  // 既存行を検索
  if (lastRow > 1) {
    const names = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < names.length; i++) {
      if (names[i][0] === name) {
        rowIndex = i + 2;
        break;
      }
    }
  }

  const progress = data.progress || {};
  const row = [
    name,
    new Date().toLocaleString('ja-JP'),
    progress.completedLessons || 0,
    progress.totalLessons || 0,
    progress.totalLessons ? Math.round((progress.completedLessons || 0) / progress.totalLessons * 100) + '%' : '0%',
    progress.completedQuizzes || 0,
    progress.totalQuizzes || 0,
    progress.totalQuizzes ? Math.round((progress.completedQuizzes || 0) / progress.totalQuizzes * 100) + '%' : '0%',
    progress.writtenCount || 0,
    progress.day1 || '0%',
    progress.day2 || '0%',
    progress.day3 || '0%'
  ];

  if (rowIndex > 0) {
    sheet.getRange(rowIndex, 1, 1, 12).setValues([row]);
  } else {
    sheet.appendRow(row);
  }

  // ダッシュボード更新
  updateDashboard(ss);
}

// ── 管理ダッシュボードの更新 ──
function updateDashboard(ss) {
  const dashSheet = ss.getSheetByName('管理ダッシュボード');
  const summarySheet = ss.getSheetByName('受講者一覧');
  if (!dashSheet || !summarySheet) return;

  // 既存データをクリア
  const lastRow = dashSheet.getLastRow();
  if (lastRow > 2) {
    dashSheet.getRange(3, 1, lastRow - 2, 6).clear();
  }

  dashSheet.getRange('A2').setValue('最終更新: ' + new Date().toLocaleString('ja-JP'));

  // サマリーデータ取得
  const summaryLastRow = summarySheet.getLastRow();
  if (summaryLastRow <= 1) return;

  const summaryData = summarySheet.getRange(2, 1, summaryLastRow - 1, 12).getValues();

  // 全体統計
  dashSheet.getRange('A4').setValue('全体統計').setFontWeight('bold').setFontSize(12).setFontColor('#5B0E2D');
  dashSheet.getRange('A5').setValue('受講者数');
  dashSheet.getRange('B5').setValue(summaryData.length + '名');
  dashSheet.getRange('A6').setValue('全員完了');

  const allComplete = summaryData.filter(r => r[4] === '100%').length;
  dashSheet.getRange('B6').setValue(allComplete + '名 / ' + summaryData.length + '名');

  // 受講者別進捗
  dashSheet.getRange('A8').setValue('受講者別進捗').setFontWeight('bold').setFontSize(12).setFontColor('#5B0E2D');
  dashSheet.getRange('A9:F9').setValues([['受講者名', '進捗率', 'STEP1', 'STEP2', 'STEP3', '最終アクセス']])
    .setFontWeight('bold').setBackground('#F3EDE7');

  summaryData.forEach((row, i) => {
    dashSheet.getRange(10 + i, 1, 1, 6).setValues([[
      row[0], row[4], row[9], row[10], row[11], row[1]
    ]]);

    // 進捗率に応じて色分け
    const pct = parseInt(row[4]) || 0;
    const color = pct === 100 ? '#E8F5E9' : pct >= 50 ? '#FFF8E1' : '#FFF3F0';
    dashSheet.getRange(10 + i, 2).setBackground(color);
  });
}

// ── ダッシュボードデータ取得（フロントエンド用） ──
function getDashboardData(data) {
  const ss = initializeSheets();
  const summarySheet = ss.getSheetByName('受講者一覧');
  const logSheet = ss.getSheetByName('進捗ログ');

  const result = { trainees: [], recentLogs: [] };

  // 受講者一覧
  if (summarySheet && summarySheet.getLastRow() > 1) {
    const rows = summarySheet.getRange(2, 1, summarySheet.getLastRow() - 1, 12).getValues();
    result.trainees = rows.map(r => ({
      name: r[0],
      lastAccess: r[1],
      completedLessons: r[2],
      totalLessons: r[3],
      progressRate: r[4],
      completedQuizzes: r[5],
      totalQuizzes: r[6],
      quizRate: r[7],
      writtenCount: r[8],
      day1: r[9],
      day2: r[10],
      day3: r[11]
    }));
  }

  // 直近ログ（最新50件）
  if (logSheet && logSheet.getLastRow() > 1) {
    const lastRow = logSheet.getLastRow();
    const startRow = Math.max(2, lastRow - 49);
    const rows = logSheet.getRange(startRow, 1, lastRow - startRow + 1, 13).getValues();
    result.recentLogs = rows.reverse().map(r => ({
      timestamp: r[0],
      name: r[1],
      type: r[2],
      day: r[3],
      moduleId: r[4],
      moduleName: r[5],
      lessonId: r[6],
      lessonName: r[7],
      quizCorrect: r[8],
      quizTotal: r[9],
      quizPassed: r[10],
      writtenAnswer: r[11],
      writtenQuestion: r[12]
    }));
  }

  return result;
}

// ── メニュー追加（手動初期化用） ──
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('研修管理')
    .addItem('シートを初期化', 'initializeSheets')
    .addItem('ダッシュボード更新', 'refreshDashboard')
    .addToUi();
}

function refreshDashboard() {
  const ss = initializeSheets();
  updateDashboard(ss);
  SpreadsheetApp.getUi().alert('ダッシュボードを更新しました。');
}

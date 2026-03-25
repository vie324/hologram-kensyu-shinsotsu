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
 * - 「受講者マスタ」: 受講者のID・名前・作成日時
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

// ── ユニークID生成 ──
function generateId() {
  const chars = 'abcdefghijklmnopqrstuvwxyz0123456789';
  let id = '';
  for (let i = 0; i < 8; i++) {
    id += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return id;
}

// ── シート初期化 ──
function initializeSheets() {
  const ss = getSpreadsheet();

  // 受講者マスタシート
  let masterSheet = ss.getSheetByName('受講者マスタ');
  if (!masterSheet) {
    masterSheet = ss.insertSheet('受講者マスタ');
    masterSheet.appendRow(['受講者ID', '受講者名', '作成日時', '詳細進捗JSON']);
    masterSheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#5B0E2D').setFontColor('#FFFFFF');
    masterSheet.setFrozenRows(1);
    masterSheet.setColumnWidth(1, 120);
    masterSheet.setColumnWidth(2, 150);
    masterSheet.setColumnWidth(3, 180);
    masterSheet.setColumnWidth(4, 400);
  }

  // 進捗ログシート
  let logSheet = ss.getSheetByName('進捗ログ');
  if (!logSheet) {
    logSheet = ss.insertSheet('進捗ログ');
    logSheet.appendRow([
      'タイムスタンプ', '受講者ID', '受講者名', 'イベント種別',
      'ステップ', 'モジュールID', 'モジュール名',
      'レッスンID', 'レッスン名',
      'クイズ正答数', 'クイズ問題数', 'クイズ合格',
      '記述回答内容', '記述問題文'
    ]);
    logSheet.getRange(1, 1, 1, 14).setFontWeight('bold').setBackground('#5B0E2D').setFontColor('#FFFFFF');
    logSheet.setFrozenRows(1);
    logSheet.setColumnWidth(1, 160);
    logSheet.setColumnWidth(2, 100);
    logSheet.setColumnWidth(3, 120);
    logSheet.setColumnWidth(4, 100);
    logSheet.setColumnWidth(13, 400);
    logSheet.setColumnWidth(14, 300);
  }

  // 受講者一覧シート
  let summarySheet = ss.getSheetByName('受講者一覧');
  if (!summarySheet) {
    summarySheet = ss.insertSheet('受講者一覧');
    summarySheet.appendRow([
      '受講者ID', '受講者名', '最終アクセス', '完了レッスン数', '全レッスン数',
      '進捗率', 'クイズ合格数', '全クイズ数', 'クイズ合格率',
      '記述回答数', 'STEP1進捗', 'STEP2進捗', 'STEP3進捗'
    ]);
    summarySheet.getRange(1, 1, 1, 13).setFontWeight('bold').setBackground('#5B0E2D').setFontColor('#FFFFFF');
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
      case 'register_trainee':
        result = registerTrainee(data);
        break;
      case 'log_progress':
        result = logProgress(data);
        break;
      case 'sync_full':
        result = syncFullProgress(data);
        break;
      case 'get_dashboard':
        result = getDashboardData(data);
        break;
      case 'get_trainee_progress':
        result = getTraineeProgress(data);
        break;
      case 'list_trainees':
        result = listTrainees(data);
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
  const traineeId = (e && e.parameter && e.parameter.id) || '';

  let result;
  switch (action) {
    case 'get_dashboard':
      result = getDashboardData({});
      break;
    case 'get_trainee_progress':
      result = traineeId ? getTraineeProgress({ traineeId: traineeId }) : { error: 'id is required' };
      break;
    case 'list_trainees':
      result = listTrainees({});
      break;
    default:
      result = { status: 'ok', message: 'HOLOGRAM Training API' };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── 受講者登録 ──
function registerTrainee(data) {
  const ss = initializeSheets();
  const masterSheet = ss.getSheetByName('受講者マスタ');
  const name = data.traineeName || '未入力';

  // 既存チェック（同名がいないか）
  const lastRow = masterSheet.getLastRow();
  if (lastRow > 1) {
    const names = masterSheet.getRange(2, 2, lastRow - 1, 1).getValues();
    for (let i = 0; i < names.length; i++) {
      if (names[i][0] === name) {
        const existingId = masterSheet.getRange(i + 2, 1).getValue();
        return { status: 'ok', traineeId: existingId, name: name, message: '既に登録済みです' };
      }
    }
  }

  // 新規登録（クライアントからIDが渡された場合はそれを使用）
  const id = data.traineeId || generateId();
  masterSheet.appendRow([id, name, new Date().toLocaleString('ja-JP'), '{}']);

  return { status: 'ok', traineeId: id, name: name, message: '登録完了' };
}

// ── 受講者一覧取得 ──
function listTrainees(data) {
  const ss = initializeSheets();
  const masterSheet = ss.getSheetByName('受講者マスタ');
  const summarySheet = ss.getSheetByName('受講者一覧');

  const trainees = [];
  const lastRow = masterSheet.getLastRow();
  if (lastRow <= 1) return { trainees: [] };

  const masterData = masterSheet.getRange(2, 1, lastRow - 1, 3).getValues();

  // サマリーデータも取得
  const summaryMap = {};
  const sLastRow = summarySheet.getLastRow();
  if (sLastRow > 1) {
    const sData = summarySheet.getRange(2, 1, sLastRow - 1, 13).getValues();
    sData.forEach(r => { summaryMap[r[0]] = r; });
  }

  masterData.forEach(row => {
    const id = row[0];
    const summary = summaryMap[id];
    trainees.push({
      id: id,
      name: row[1],
      createdAt: row[2],
      progressRate: summary ? summary[5] : '0%',
      completedLessons: summary ? summary[3] : 0,
      totalLessons: summary ? summary[4] : 0,
      lastAccess: summary ? summary[2] : '-',
    });
  });

  return { trainees: trainees };
}

// ── 受講者の詳細進捗取得（ページ復元用） ──
function getTraineeProgress(data) {
  const ss = initializeSheets();
  const masterSheet = ss.getSheetByName('受講者マスタ');
  const traineeId = data.traineeId;

  if (!traineeId) return { error: 'traineeId is required' };

  const lastRow = masterSheet.getLastRow();
  if (lastRow <= 1) return { error: 'trainee not found' };

  const ids = masterSheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === traineeId) {
      const rowIdx = i + 2;
      const name = masterSheet.getRange(rowIdx, 2).getValue();
      const jsonStr = masterSheet.getRange(rowIdx, 4).getValue() || '{}';
      let state = {};
      try { state = JSON.parse(jsonStr); } catch (e) { state = {}; }
      return {
        status: 'ok',
        traineeId: traineeId,
        name: name,
        state: state, // { done, qDone, qAns, qSub, wAns }
      };
    }
  }

  return { error: 'trainee not found' };
}

// ── 進捗ログの記録 ──
function logProgress(data) {
  const ss = initializeSheets();
  const logSheet = ss.getSheetByName('進捗ログ');
  const timestamp = new Date().toLocaleString('ja-JP');
  const traineeId = data.traineeId || '';
  const name = data.traineeName || '未入力';

  const events = data.events || [];
  events.forEach(ev => {
    logSheet.appendRow([
      timestamp,
      traineeId,
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
  updateSummary(ss, traineeId, name, data);
  // 詳細進捗を保存
  saveDetailedState(ss, traineeId, data);

  return { status: 'ok', logged: events.length };
}

// ── 全進捗の一括同期 ──
function syncFullProgress(data) {
  const ss = initializeSheets();
  const traineeId = data.traineeId || '';
  const name = data.traineeName || '未入力';

  // サマリー更新
  updateSummary(ss, traineeId, name, data);
  // 詳細進捗を保存
  saveDetailedState(ss, traineeId, data);

  return { status: 'ok', synced: true };
}

// ── 詳細進捗の保存（受講者マスタのJSON列） ──
function saveDetailedState(ss, traineeId, data) {
  if (!traineeId || !data.detailedState) return;

  const masterSheet = ss.getSheetByName('受講者マスタ');
  if (!masterSheet) return;

  const lastRow = masterSheet.getLastRow();
  if (lastRow <= 1) return;

  const ids = masterSheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === traineeId) {
      const jsonStr = JSON.stringify(data.detailedState);
      masterSheet.getRange(i + 2, 4).setValue(jsonStr);
      return;
    }
  }
}

// ── 受講者サマリーの更新 ──
function updateSummary(ss, traineeId, name, data) {
  const sheet = ss.getSheetByName('受講者一覧');
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  let rowIndex = -1;

  // 既存行を検索（IDで）
  if (lastRow > 1) {
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (ids[i][0] === traineeId) {
        rowIndex = i + 2;
        break;
      }
    }
  }

  const progress = data.progress || {};
  const row = [
    traineeId,
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
    sheet.getRange(rowIndex, 1, 1, 13).setValues([row]);
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

  const summaryData = summarySheet.getRange(2, 1, summaryLastRow - 1, 13).getValues();

  // 全体統計
  dashSheet.getRange('A4').setValue('全体統計').setFontWeight('bold').setFontSize(12).setFontColor('#5B0E2D');
  dashSheet.getRange('A5').setValue('受講者数');
  dashSheet.getRange('B5').setValue(summaryData.length + '名');
  dashSheet.getRange('A6').setValue('全員完了');

  const allComplete = summaryData.filter(r => r[5] === '100%').length;
  dashSheet.getRange('B6').setValue(allComplete + '名 / ' + summaryData.length + '名');

  // 受講者別進捗
  dashSheet.getRange('A8').setValue('受講者別進捗').setFontWeight('bold').setFontSize(12).setFontColor('#5B0E2D');
  dashSheet.getRange('A9:F9').setValues([['受講者名', '進捗率', 'STEP1', 'STEP2', 'STEP3', '最終アクセス']])
    .setFontWeight('bold').setBackground('#F3EDE7');

  summaryData.forEach((row, i) => {
    dashSheet.getRange(10 + i, 1, 1, 6).setValues([[
      row[1], row[5], row[10], row[11], row[12], row[2]
    ]]);

    // 進捗率に応じて色分け
    const pct = parseInt(row[5]) || 0;
    const color = pct === 100 ? '#E8F5E9' : pct >= 50 ? '#FFF8E1' : '#FFF3F0';
    dashSheet.getRange(10 + i, 2).setBackground(color);
  });
}

// ── ダッシュボードデータ取得（フロントエンド用） ──
function getDashboardData(data) {
  const ss = initializeSheets();
  const summarySheet = ss.getSheetByName('受講者一覧');
  const logSheet = ss.getSheetByName('進捗ログ');
  const masterSheet = ss.getSheetByName('受講者マスタ');

  const result = { trainees: [], recentLogs: [] };

  // 受講者マスタからID・名前を取得
  const idMap = {};
  if (masterSheet && masterSheet.getLastRow() > 1) {
    const masterData = masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, 3).getValues();
    masterData.forEach(r => { idMap[r[0]] = { name: r[1], createdAt: r[2] }; });
  }

  // 受講者一覧
  if (summarySheet && summarySheet.getLastRow() > 1) {
    const rows = summarySheet.getRange(2, 1, summarySheet.getLastRow() - 1, 13).getValues();
    result.trainees = rows.map(r => ({
      id: r[0],
      name: r[1],
      lastAccess: r[2],
      completedLessons: r[3],
      totalLessons: r[4],
      progressRate: r[5],
      completedQuizzes: r[6],
      totalQuizzes: r[7],
      quizRate: r[8],
      writtenCount: r[9],
      day1: r[10],
      day2: r[11],
      day3: r[12]
    }));
  }

  // 直近ログ（最新50件）
  if (logSheet && logSheet.getLastRow() > 1) {
    const lastRow = logSheet.getLastRow();
    const startRow = Math.max(2, lastRow - 49);
    const rows = logSheet.getRange(startRow, 1, lastRow - startRow + 1, 14).getValues();
    result.recentLogs = rows.reverse().map(r => ({
      timestamp: r[0],
      traineeId: r[1],
      name: r[2],
      type: r[3],
      day: r[4],
      moduleId: r[5],
      moduleName: r[6],
      lessonId: r[7],
      lessonName: r[8],
      quizCorrect: r[9],
      quizTotal: r[10],
      quizPassed: r[11],
      writtenAnswer: r[12],
      writtenQuestion: r[13]
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

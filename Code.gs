const SHEET_ROSTER = 'QR用名簿';
const SHEET_LOG = '提出履歴';
const SHEET_SUMMARY = '児童集計';
const SHEET_ASSIGNMENTS = '提出回数管理';
const VALID_TYPES = ['A', 'B', 'C', 'D'];


// ==================================================
// GitHub Pages から POST を受け取る
// ==================================================

function doPost(e) {

  try {

    const data = JSON.parse(
      e.postData.contents || '{}'
    );


    const correctKey =
      PropertiesService
        .getScriptProperties()
        .getProperty('API_KEY');


    if (
      !correctKey ||
      data.apiKey !== correctKey
    ) {

      return jsonResponse({
        ok: false,
        unauthorized: true,
        message: '管理キーが違います。'
      });

    }


    if (data.action === 'startSession') {
      return jsonResponse(
        startSession(data.className, data.submissionType, data.sessionId)
      );
    }


    return jsonResponse(
      registerSubmission(
        data.studentId,
        data.submissionType,
        data.className,
        data.isLateSubmission
      )
    );

  }

  catch (error) {

    return jsonResponse({
      ok: false,
      message: String(error.message || error)
    });

  }

}


// ==================================================
// API URLをブラウザで開いたときの確認用
// ==================================================

function doGet() {

  return ContentService
    .createTextOutput(
      'QR submission API is running.'
    );

}


// ==================================================
// JSONレスポンス
// ==================================================

function jsonResponse(data) {

  return ContentService
    .createTextOutput(
      JSON.stringify(data)
    )
    .setMimeType(
      ContentService.MimeType.JSON
    );

}


// ==================================================
// 受付開始：提出機会を1回記録する
// ==================================================

function startSession(className, submissionType, sessionId) {

  className = String(className || '').trim();
  submissionType = String(submissionType || '').trim().toUpperCase();
  sessionId = String(sessionId || '').trim();

  if (!className || !VALID_TYPES.includes(submissionType) || !sessionId) {
    return { ok: false, message: '受付開始の情報が不足しています。' };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_ASSIGNMENTS);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_ASSIGNMENTS);
      sheet.appendRow(['タイムスタンプ', '日付', 'クラス', '提出物', '受付ID']);
    }

    const now = new Date();
    const tz = ss.getSpreadsheetTimeZone() || 'America/Los_Angeles';
    const today = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
    if (!['3-1', '3-2'].includes(className)) {
      return { ok: false, message: 'クラスを選択してください。' };
    }
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const sessions = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
      if (sessions.some(row => String(row[4]).trim() === sessionId ||
        (dateToKey(row[1], tz) === today && String(row[2]).trim() === className &&
         String(row[3]).trim() === submissionType))) {
        return { ok: true, alreadyStarted: true };
      }
    }

    sheet.appendRow([now, now, className, submissionType, sessionId]);
    return { ok: true };

  } finally {
    lock.releaseLock();
  }

}


// ==================================================
// 提出登録
//
// 通常の登録では、提出履歴全体を読みません。
// 「児童集計」の該当する児童1人の行だけを確認します。
// ==================================================

function registerSubmission(
  studentId,
  submissionType,
  expectedClassName,
  isLateSubmission
) {

  isLateSubmission = isLateSubmission === true;

  studentId =
    String(studentId || '').trim();

  submissionType =
    String(submissionType || '')
      .trim()
      .toUpperCase();


  if (!studentId) {

    return {
      ok: false,
      message: '生徒IDがありません。'
    };

  }


  if (!VALID_TYPES.includes(submissionType)) {

    return {
      ok: false,
      message: '提出物A〜Dを選んでください。'
    };

  }


  const lock =
    LockService.getScriptLock();

  lock.waitLock(10000);


  try {

    const ss =
      SpreadsheetApp.getActiveSpreadsheet();

    const rosterSheet =
      ss.getSheetByName(SHEET_ROSTER);

    const logSheet =
      ss.getSheetByName(SHEET_LOG);

    const summarySheet =
      ss.getSheetByName(SHEET_SUMMARY);


    if (!rosterSheet) {

      return {
        ok: false,
        message: 'QR用名簿シートが見つかりません。'
      };

    }


    if (!logSheet) {

      return {
        ok: false,
        message: '提出履歴シートが見つかりません。'
      };

    }


    if (!summarySheet) {

      return {
        ok: false,
        message:
          '児童集計シートが見つかりません。'
      };

    }


    const rosterRow =
      findStudentRow(
        rosterSheet,
        studentId
      );

    const summaryRow =
      findStudentRow(
        summarySheet,
        studentId
      );


    if (!rosterRow) {

      return {
        ok: false,
        message:
          `名簿にないIDです：${studentId}`
      };

    }


    if (!summaryRow) {

      return {
        ok: false,
        message:
          '児童集計にこの児童がいません。'
      };

    }


    const student =
      rosterSheet
        .getRange(
          rosterRow,
          1,
          1,
          4
        )
        .getValues()[0];

    const id = student[0];
    const className = student[1];
    const number = student[2];
    const name = student[3];

    if (
      expectedClassName &&
      String(className).trim() !== String(expectedClassName).trim()
    ) {
      return {
        ok: false,
        message: `${className}の児童です。選択したクラスを確認してください。`
      };
    }


    const summary =
      summarySheet
        .getRange(
          summaryRow,
          1,
          1,
          12
        )
        .getValues()[0];

    const columns =
      summaryColumns(submissionType);

    const now = new Date();

    const tz =
      ss.getSpreadsheetTimeZone()
      || 'America/Los_Angeles';

    const todayKey =
      Utilities.formatDate(
        now,
        tz,
        'yyyy-MM-dd'
      );

    const lastDate =
      dateToKey(
        summary[columns.lastDateIndex],
        tz
      );


    // 通常提出は同日重複を防ぐ。
    // 「遅れて提出」を選んだ場合のみ追加登録を許可する。
    if (lastDate === todayKey && !isLateSubmission) {

      return {

        ok: false,

        duplicate: true,

        name: name,

        className: className,

        number: number,

        submissionType:
          submissionType,

        message:
          `${name}さんは今日すでに提出物${submissionType}を登録済みです。`

      };

    }


    // 遅れて提出かどうかを履歴のG列に残す。
    if (logSheet.getRange(1, 7).getValue() !== '提出区分') {
      logSheet.getRange(1, 7).setValue('提出区分');
    }

    // 正式な提出履歴は、従来どおり全件残す
    logSheet.appendRow([

      now,
      now,
      id,
      name,
      className,
      submissionType,
      isLateSubmission ? '遅れ提出' : '通常提出'

    ]);


    // 受付用の小さな集計表だけを更新する
    const oldCount =
      Number(summary[columns.countIndex])
      || 0;

    const newCount =
      oldCount + 1;

    summarySheet
      .getRange(
        summaryRow,
        columns.countIndex + 1,
        1,
        2
      )
      .setValues([
        [
          newCount,
          isLateSubmission ? summary[columns.lastDateIndex] : todayKey
        ]
      ]);


    return {

      ok: true,

      studentId: id,

      name: name,

      className: className,

      number: number,

      submissionType:
        submissionType,

      count: newCount,

      date:
        Utilities.formatDate(
          now,
          tz,
          'M/d'
        )

    };

  }

  finally {

    lock.releaseLock();

  }

}


// ==================================================
// 「児童集計」の列位置
// ==================================================

function summaryColumns(submissionType) {

  const typeIndex =
    VALID_TYPES.indexOf(
      submissionType
    );

  return {

    // A〜D は名簿情報。
    // E/F, G/H, I/J, K/L が
    // 各提出物の「回数 / 最終日」。
    countIndex:
      4 + typeIndex * 2,

    lastDateIndex:
      5 + typeIndex * 2

  };

}


// ==================================================
// 指定した生徒IDの行を探す
// ==================================================

function findStudentRow(
  sheet,
  studentId
) {

  const lastRow =
    sheet.getLastRow();

  if (lastRow < 2) {

    return 0;

  }


  const cell =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        1
      )
      .createTextFinder(
        studentId
      )
      .matchEntireCell(true)
      .findNext();


  return cell
    ? cell.getRow()
    : 0;

}


// ==================================================
// 日付を yyyy-MM-dd の文字列へそろえる
// ==================================================

function dateToKey(
  value,
  timezone
) {

  if (
    value instanceof Date
    && !isNaN(value)
  ) {

    return Utilities.formatDate(
      value,
      timezone,
      'yyyy-MM-dd'
    );

  }

  return String(value || '').trim();

}


// ==================================================
// 管理者用：名簿や過去の履歴から
// 「児童集計」を作り直す
//
// 新しい児童を名簿へ追加した後などに、
// Apps Script エディタから手動で実行します。
// 通常のQR登録では実行されません。
// ==================================================

function rebuildSummarySheet() {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    return rebuildSummarySheetUnlocked();
  } finally {
    lock.releaseLock();
  }
}

function rebuildSummarySheetUnlocked() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const rosterSheet =
    ss.getSheetByName(SHEET_ROSTER);

  const logSheet =
    ss.getSheetByName(SHEET_LOG);


  if (!rosterSheet || !logSheet) {

    throw new Error(
      'QR用名簿または提出履歴が見つかりません。'
    );

  }


  let summarySheet =
    ss.getSheetByName(SHEET_SUMMARY);

  if (!summarySheet) {

    summarySheet =
      ss.insertSheet(SHEET_SUMMARY);

  }


  const rosterLastRow =
    rosterSheet.getLastRow();

  const roster =
    rosterLastRow >= 2
      ? rosterSheet
          .getRange(
            2,
            1,
            rosterLastRow - 1,
            4
          )
          .getValues()
      : [];


  const summaryById = {};

  roster.forEach(row => {

    const id =
      String(row[0] || '').trim();

    if (!id) {
      return;
    }

    summaryById[id] = {
      roster: row,
      counts: [0, 0, 0, 0],
      lastDates: ['', '', '', '']
    };

  });


  const logLastRow =
    logSheet.getLastRow();

  const logs =
    logLastRow >= 2
      ? logSheet
          .getRange(
            2,
            1,
            logLastRow - 1,
            7
          )
          .getValues()
      : [];

  const tz =
    ss.getSpreadsheetTimeZone()
    || 'America/Los_Angeles';


  logs.forEach(row => {

    const id =
      String(row[2] || '').trim();

    const type =
      String(row[5] || '')
        .trim()
        .toUpperCase();

    const typeIndex =
      VALID_TYPES.indexOf(type);

    const item =
      summaryById[id];

    if (!item || typeIndex === -1) {
      return;
    }

    item.counts[typeIndex] += 1;

    const dateKey =
      dateToKey(
        row[1],
        tz
      );

    if (
      row[6] !== '遅れ提出' && dateKey &&
      (
        !item.lastDates[typeIndex] ||
        dateKey > item.lastDates[typeIndex]
      )
    ) {

      item.lastDates[typeIndex] =
        dateKey;

    }

  });


  const header = [
    '生徒ID',
    'クラス',
    '出席番号',
    '氏名',
    '視写 回数',
    '視写 最終日',
    '音読 回数',
    '音読 最終日',
    'ぴったりドリル 回数',
    'ぴったりドリル 最終日',
    '算数ドリル 回数',
    '算数ドリル 最終日'
  ];

  const rows = [header];

  roster.forEach(row => {

    const id =
      String(row[0] || '').trim();

    const item =
      summaryById[id];

    if (!item) return;

    rows.push([
      row[0],
      row[1],
      row[2],
      row[3],
      item.counts[0],
      item.lastDates[0],
      item.counts[1],
      item.lastDates[1],
      item.counts[2],
      item.lastDates[2],
      item.counts[3],
      item.lastDates[3]
    ]);

  });


  summarySheet.clearContents();

  summarySheet
    .getRange(
      1,
      1,
      rows.length,
      header.length
    )
    .setValues(rows);

  summarySheet.setFrozenRows(1);

}

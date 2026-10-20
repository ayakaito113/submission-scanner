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

    const lastDateCell = summarySheet.getRange(
      summaryRow, columns.lastDateIndex + 1
    );
    const lastDate = summaryDateKey(
      summary[columns.lastDateIndex], lastDateCell.getDisplayValue(), tz
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

    // 日付キーは文字列として保存し、シートの自動日付変換を防ぐ。
    lastDateCell.setNumberFormat('@');
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
          isLateSubmission ? lastDate : todayKey
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

    // 次の読み取りを許可する前に集計表への書き込みを確定する。
    try {
      SpreadsheetApp.flush();
    } finally {
      lock.releaseLock();
    }

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

// 既存の「最終日」が日付セルになっていても、表示上の日付を優先する。
// 通常の日時（提出履歴など）は従来どおりタイムゾーンで変換する。
function summaryDateKey(value, displayedValue, timezone) {
  const displayed = String(displayedValue || '').trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(displayed)) return displayed;
  const key = dateToKey(value, timezone);
  if (!key || /^\d{4}-\d{2}-\d{2}$/.test(key)) return key;
  throw new Error('児童集計の最終日を読み取れません。日付の表示形式を yyyy-mm-dd にしてください。');
}


// ==================================================
// 管理者用：名簿や過去の履歴から
// 「児童集計」を作り直す
//
// 履歴からの復旧時にだけ手動で実行します。
// 転入生の追加には「児童管理」メニューを使ってください。
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

  // 再集計後も最終日が日付セルへ自動変換されないようにする。
  if (rows.length > 1) {
    [6, 8, 10, 12].forEach(column => {
      summarySheet.getRange(2, column, rows.length - 1, 1).setNumberFormat('@');
    });
  }

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


// ==================================================
// スプレッドシート専用：転入生の追加（Web APIからは呼び出さない）
// ==================================================
function onOpen() {
  SpreadsheetApp.getUi().createMenu('児童管理')
    .addItem('転入生を追加', 'showTransferStudentPrompt')
    .addToUi();
}

function showTransferStudentPrompt() {
  const ui = SpreadsheetApp.getUi();
  const answers = [];
  for (const question of ['クラス（3-1 または 3-2）', '出席番号（1〜99）', '児童の名前']) {
    const result = ui.prompt('転入生を追加', question, ui.ButtonSet.OK_CANCEL);
    if (result.getSelectedButton() !== ui.Button.OK) return;
    answers.push(result.getResponseText());
  }
  try {
    const student = validateTransferStudent_(answers[0], answers[1], answers[2]);
    const message = student.className + '　' + student.number + '番\n' + student.name +
      '\n生徒ID：' + student.id +
      '\n\n名簿・QRコード・集計表に追加します。\n既存の提出回数は変更しません。よろしいですか？';
    if (ui.alert('登録内容の確認', message, ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
    // prompt/alert はロックを維持できないので、確認が終わってから取得する。
    const result = addTransferStudent_(student.className, student.number, student.name);
    ui.alert('追加しました', result.id +
      ' を追加しました。QR用名簿の新しい行にQRコードがあります。', ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('追加できませんでした', String(error.message || error), ui.ButtonSet.OK);
  }
}

function validateTransferStudent_(className, number, name) {
  className = String(className || '').normalize('NFKC').trim().replace(/[−ー–—]/g, '-');
  const digits = String(number == null ? '' : number).normalize('NFKC').trim();
  name = String(name || '').trim();
  if (!['3-1', '3-2'].includes(className)) throw new Error('クラスは 3-1 または 3-2 を入力してください。');
  if (!/^\d{1,2}$/.test(digits) || Number(digits) < 1)
    throw new Error('出席番号は1〜99の整数を入力してください。');
  if (!name || name.length > 100 || /[\r\n\t]/.test(name))
    throw new Error('名前は改行せず、1〜100文字で入力してください。');
  // 数式として解釈される文字を拒否。氏名や入力内容を数式に埋め込まない。
  if (/^[=+@'\-]/.test(name)) throw new Error('名前の先頭に =、+、-、@、引用符は使えません。');
  return { className: className, number: Number(digits), name: name,
    id: className.replace('-', '') + '-' + String(Number(digits)).padStart(2, '0') };
}

function addTransferStudent_(className, number, name) {
  const student = validateTransferStudent_(className, number, name);
  // QRはGAS内だけで生成する。外部のQR作成サイトへ児童情報を送らない。
  const qrBlob = studentQrPng_(student.id);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const roster = ss.getSheetByName(SHEET_ROSTER);
    const summary = ss.getSheetByName(SHEET_SUMMARY);
    const grade = ss.getSheetByName(student.className + ' 成績用集計');
    if (!roster || !summary || !grade || !ss.getSheetByName(SHEET_ASSIGNMENTS))
      throw new Error('名簿・児童集計・クラス別成績用集計・提出回数管理のシートを確認してください。');
    // 旧形式のIDでも「同じクラス・出席番号」があれば追加しない。
    for (const sheet of [roster, summary]) {
      const count = sheet.getLastRow() - 1;
      if (count < 1) continue;
      const range = sheet.getRange(2, 1, count, 4);
      const rows = range.getValues();
      const displayed = range.getDisplayValues();
      if (rows.some((row, i) => {
        const storedClass = row[1] instanceof Date
          ? Utilities.formatDate(row[1], ss.getSpreadsheetTimeZone(), 'M-d')
          : String(row[1]).trim();
        return String(row[0]).trim() === student.id ||
          ((storedClass === student.className || displayed[i][1] === student.className) &&
           Number(row[2]) === student.number);
      })) throw new Error('同じ生徒ID、または同じクラス・出席番号がすでに登録されています。追加は行いません。');
    }
    // 並べ替えや全体再集計をせず、末尾の空き行だけに追加する。
    const plans = [[roster, 5], [summary, 12], [grade, 8]].map(item => {
      const sheet = item[0], width = item[1], row = Math.max(2, sheet.getLastRow() + 1);
      if (row > sheet.getMaxRows())
        throw new Error(sheet.getName() + ' の末尾に空白行を追加してから、もう一度お試しください。');
      const range = sheet.getRange(row, 1, 1, width);
      if (!range.canEdit()) throw new Error(sheet.getName() + ' の追加先を編集する権限がありません。');
      return { sheet: sheet, row: row, range: range };
    });
    const written = [];
    let image = null;
    try {
      for (const plan of plans) {
        written.push(plan);
        if (plan.row > 2)
          plan.sheet.getRange(plan.row - 1, 1, 1, plan.range.getNumColumns())
            .copyTo(plan.range, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      }
      const r = plans[0].row, s = plans[1].row, g = plans[2].row;
      for (const plan of plans.slice(0, 2)) {
        plan.sheet.getRange(plan.row, 1, 1, 2).setNumberFormat('@');
        plan.sheet.getRange(plan.row, 3).setNumberFormat('0');
        plan.sheet.getRange(plan.row, 4).setNumberFormat('@');
      }
      roster.getRange(r, 1, 1, 4).setValues([[student.id, student.className, student.number, student.name]]);
      [6, 8, 10, 12].forEach(c => summary.getRange(s, c).setNumberFormat('@'));
      summary.getRange(s, 1, 1, 12).setValues([[
        student.id, student.className, student.number, student.name, 0, '', 0, '', 0, '', 0, ''
      ]]);
      grade.getRange(g, 1, 1, 8).setFormulas([transferGradeFormulas_(student, g)]);
      grade.getRange(g, 6).setNumberFormat('0%');
      grade.getRange(g, 8).setNumberFormat('0%');
      image = roster.insertImage(qrBlob, 5, r);
      image.setAltTextTitle(student.id).setAltTextDescription('生徒ID：' + student.id);
      image.setWidth(116).setHeight(116);
      roster.setRowHeight(r, 120);
      SpreadsheetApp.flush();
      return { id: student.id };
    } catch (error) {
      // 新規行のみ取り消す。既存の児童・提出履歴・提出機会は触らない。
      let rollbackFailed = false;
      if (image) { try { image.remove(); } catch (_) { rollbackFailed = true; } }
      for (const plan of written) {
        try { plan.range.clearContent(); } catch (_) { rollbackFailed = true; }
      }
      try { SpreadsheetApp.flush(); } catch (_) { rollbackFailed = true; }
      if (rollbackFailed)
        throw new Error('追加途中で停止しました。再実行せず、名簿と集計表の末尾を確認してください。ID：' + student.id);
      throw new Error('追加を取り消しました。既存の記録は変更していません。原因：' + String(error.message || error));
    }
  } finally {
    lock.releaseLock();
  }
}

function transferGradeFormulas_(student, row) {
  const lookup = col => '=IFERROR(INDEX(\'児童集計\'!' + col + ':' + col +
    ',MATCH("' + student.id + '",\'児童集計\'!A:A,0)),"")';
  const opportunities = type => 'COUNTIFS(\'提出回数管理\'!C:C,"' + student.className +
    '",\'提出回数管理\'!D:D,"' + type + '")';
  return [lookup('C'), lookup('D'), lookup('E'), lookup('G'), lookup('I'),
    '=IF(A' + row + '="","",IFERROR((C' + row + '+D' + row + '+E' + row + ')/(' +
      ['A', 'B', 'C'].map(opportunities).join('+') + '),""))',
    lookup('K'),
    '=IF(A' + row + '="","",IFERROR(G' + row + '/' + opportunities('D') + ',""))'];
}

// QR行列をPNGに変換（白黒8bit、無圧縮zlib）。ネットワークアクセス不要。
function studentQrPng_(id) {
  const qr = qrcodegen.QrCode.encodeText(id, qrcodegen.QrCode.Ecc.MEDIUM);
  const scale = 4, border = 4, size = (qr.size + border * 2) * scale;
  const raw = [];
  for (let y = 0; y < size; y++) {
    raw.push(0);
    for (let x = 0; x < size; x++)
      raw.push(qr.getModule(Math.floor(x / scale) - border, Math.floor(y / scale) - border) ? 0 : 255);
  }
  const be32 = value => [value >>> 24, (value >>> 16) & 255, (value >>> 8) & 255, value & 255];
  const chunk = (type, data) => {
    const payload = type.split('').map(c => c.charCodeAt(0)).concat(data);
    let crc = 0xffffffff;
    for (const byte of payload) {
      crc ^= byte;
      for (let i = 0; i < 8; i++) crc = (crc >>> 1) ^ ((crc & 1) ? 0xedb88320 : 0);
    }
    return be32(data.length).concat(payload, be32((crc ^ 0xffffffff) >>> 0));
  };
  const zlib = [0x78, 0x01];
  for (let offset = 0; offset < raw.length; offset += 65535) {
    const block = raw.slice(offset, offset + 65535), n = block.length;
    zlib.push(offset + n === raw.length ? 1 : 0, n & 255, n >>> 8, (~n) & 255, ((~n) >>> 8) & 255);
    for (const byte of block) zlib.push(byte);
  }
  let a = 1, b = 0;
  for (const byte of raw) { a = (a + byte) % 65521; b = (b + a) % 65521; }
  zlib.push(...be32(((b << 16) | a) >>> 0));
  const png = [137, 80, 78, 71, 13, 10, 26, 10]
    .concat(chunk('IHDR', be32(size).concat(be32(size), [8, 0, 0, 0, 0])),
      chunk('IDAT', zlib), chunk('IEND', []));
  // GASのByte[]は符号付き。
  return Utilities.newBlob(png.map(byte => byte > 127 ? byte - 256 : byte), 'image/png', id + '.png');
}

// Vendored QR encoder: Project Nayuki, MIT. No external requests.
/* 
 * QR Code generator library (TypeScript)
 * 
 * Copyright (c) Project Nayuki. (MIT License)
 * https://www.nayuki.io/page/qr-code-generator-library
 * 
 * Permission is hereby granted, free of charge, to any person obtaining a copy of
 * this software and associated documentation files (the "Software"), to deal in
 * the Software without restriction, including without limitation the rights to
 * use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of
 * the Software, and to permit persons to whom the Software is furnished to do so,
 * subject to the following conditions:
 * - The above copyright notice and this permission notice shall be included in
 *   all copies or substantial portions of the Software.
 * - The Software is provided "as is", without warranty of any kind, express or
 *   implied, including but not limited to the warranties of merchantability,
 *   fitness for a particular purpose and noninfringement. In no event shall the
 *   authors or copyright holders be liable for any claim, damages or other
 *   liability, whether in an action of contract, tort or otherwise, arising from,
 *   out of or in connection with the Software or the use or other dealings in the
 *   Software.
 */

"use strict";
(function(qrcodegen) {
    class QrCode {


        static encodeText(text, ecl) {
            const segs = qrcodegen.QrSegment.makeSegments(text);
            return QrCode.encodeSegments(segs, ecl);
        }
        static encodeBinary(data, ecl) {
            const seg = qrcodegen.QrSegment.makeBytes(data);
            return QrCode.encodeSegments([
                seg
            ], ecl);
        }
        static encodeSegments(segs, ecl, minVersion = 1, maxVersion = 40, mask = -1, boostEcl = true) {
            if (!(QrCode.MIN_VERSION <= minVersion && minVersion <= maxVersion && maxVersion <= QrCode.MAX_VERSION) || mask < -1 || mask > 7) throw new RangeError("Invalid value");
            let version;
            let dataUsedBits;
            for(version = minVersion;; version++){
                const dataCapacityBits = QrCode.getNumDataCodewords(version, ecl) * 8;
                const usedBits = QrSegment.getTotalBits(segs, version);
                if (usedBits <= dataCapacityBits) {
                    dataUsedBits = usedBits;
                    break;
                }
                if (version >= maxVersion) throw new RangeError("Data too long");
            }
            for (const newEcl of [
                QrCode.Ecc.MEDIUM,
                QrCode.Ecc.QUARTILE,
                QrCode.Ecc.HIGH
            ]){
                if (boostEcl && dataUsedBits <= QrCode.getNumDataCodewords(version, newEcl) * 8) ecl = newEcl;
            }
            let bb = [];
            for (const seg of segs){
                appendBits(seg.mode.modeBits, 4, bb);
                appendBits(seg.numChars, seg.mode.numCharCountBits(version), bb);
                for (const b of seg.getData())bb.push(b);
            }
            assert(bb.length == dataUsedBits);
            const dataCapacityBits = QrCode.getNumDataCodewords(version, ecl) * 8;
            assert(bb.length <= dataCapacityBits);
            appendBits(0, Math.min(4, dataCapacityBits - bb.length), bb);
            appendBits(0, (8 - bb.length % 8) % 8, bb);
            assert(bb.length % 8 == 0);
            for(let padByte = 0xEC; bb.length < dataCapacityBits; padByte ^= 0xEC ^ 0x11)appendBits(padByte, 8, bb);
            let dataCodewords = [];
            while(dataCodewords.length * 8 < bb.length)dataCodewords.push(0);
            bb.forEach((b, i)=>dataCodewords[i >>> 3] |= b << 7 - (i & 7));
            return new QrCode(version, ecl, dataCodewords, mask);
        }

        constructor(version, errorCorrectionLevel, dataCodewords, msk){
            this.modules = [];
            this.isFunction = [];
            this.version = version;
            this.errorCorrectionLevel = errorCorrectionLevel;
            if (version < QrCode.MIN_VERSION || version > QrCode.MAX_VERSION) throw new RangeError("Version value out of range");
            if (msk < -1 || msk > 7) throw new RangeError("Mask value out of range");
            this.size = version * 4 + 17;
            let row = [];
            for(let i = 0; i < this.size; i++)row.push(false);
            for(let i = 0; i < this.size; i++){
                this.modules.push(row.slice());
                this.isFunction.push(row.slice());
            }
            this.drawFunctionPatterns();
            const allCodewords = this.addEccAndInterleave(dataCodewords);
            this.drawCodewords(allCodewords);
            if (msk == -1) {
                let minPenalty = 1000000000;
                for(let i = 0; i < 8; i++){
                    this.applyMask(i);
                    this.drawFormatBits(i);
                    const penalty = this.getPenaltyScore();
                    if (penalty < minPenalty) {
                        msk = i;
                        minPenalty = penalty;
                    }
                    this.applyMask(i);
                }
            }
            assert(0 <= msk && msk <= 7);
            this.mask = msk;
            this.applyMask(msk);
            this.drawFormatBits(msk);
            this.isFunction = [];
        }
        getModule(x, y) {
            return 0 <= x && x < this.size && 0 <= y && y < this.size && this.modules[y][x];
        }
        drawFunctionPatterns() {
            for(let i = 0; i < this.size; i++){
                this.setFunctionModule(6, i, i % 2 == 0);
                this.setFunctionModule(i, 6, i % 2 == 0);
            }
            this.drawFinderPattern(3, 3);
            this.drawFinderPattern(this.size - 4, 3);
            this.drawFinderPattern(3, this.size - 4);
            const alignPatPos = this.getAlignmentPatternPositions();
            const numAlign = alignPatPos.length;
            for(let i = 0; i < numAlign; i++){
                for(let j = 0; j < numAlign; j++){
                    if (!(i == 0 && j == 0 || i == 0 && j == numAlign - 1 || i == numAlign - 1 && j == 0)) this.drawAlignmentPattern(alignPatPos[i], alignPatPos[j]);
                }
            }
            this.drawFormatBits(0);
            this.drawVersion();
        }
        drawFormatBits(mask) {
            const data = this.errorCorrectionLevel.formatBits << 3 | mask;
            let rem = data;
            for(let i = 0; i < 10; i++)rem = rem << 1 ^ (rem >>> 9) * 0x537;
            const bits = (data << 10 | rem) ^ 0x5412;
            assert(bits >>> 15 == 0);
            for(let i = 0; i <= 5; i++)this.setFunctionModule(8, i, getBit(bits, i));
            this.setFunctionModule(8, 7, getBit(bits, 6));
            this.setFunctionModule(8, 8, getBit(bits, 7));
            this.setFunctionModule(7, 8, getBit(bits, 8));
            for(let i = 9; i < 15; i++)this.setFunctionModule(14 - i, 8, getBit(bits, i));
            for(let i = 0; i < 8; i++)this.setFunctionModule(this.size - 1 - i, 8, getBit(bits, i));
            for(let i = 8; i < 15; i++)this.setFunctionModule(8, this.size - 15 + i, getBit(bits, i));
            this.setFunctionModule(8, this.size - 8, true);
        }
        drawVersion() {
            if (this.version < 7) return;
            let rem = this.version;
            for(let i = 0; i < 12; i++)rem = rem << 1 ^ (rem >>> 11) * 0x1F25;
            const bits = this.version << 12 | rem;
            assert(bits >>> 18 == 0);
            for(let i = 0; i < 18; i++){
                const color = getBit(bits, i);
                const a = this.size - 11 + i % 3;
                const b = Math.floor(i / 3);
                this.setFunctionModule(a, b, color);
                this.setFunctionModule(b, a, color);
            }
        }
        drawFinderPattern(x, y) {
            for(let dy = -4; dy <= 4; dy++){
                for(let dx = -4; dx <= 4; dx++){
                    const dist = Math.max(Math.abs(dx), Math.abs(dy));
                    const xx = x + dx;
                    const yy = y + dy;
                    if (0 <= xx && xx < this.size && 0 <= yy && yy < this.size) this.setFunctionModule(xx, yy, dist != 2 && dist != 4);
                }
            }
        }
        drawAlignmentPattern(x, y) {
            for(let dy = -2; dy <= 2; dy++){
                for(let dx = -2; dx <= 2; dx++)this.setFunctionModule(x + dx, y + dy, Math.max(Math.abs(dx), Math.abs(dy)) != 1);
            }
        }
        setFunctionModule(x, y, isDark) {
            this.modules[y][x] = isDark;
            this.isFunction[y][x] = true;
        }
        addEccAndInterleave(data) {
            const ver = this.version;
            const ecl = this.errorCorrectionLevel;
            if (data.length != QrCode.getNumDataCodewords(ver, ecl)) throw new RangeError("Invalid argument");
            const numBlocks = QrCode.NUM_ERROR_CORRECTION_BLOCKS[ecl.ordinal][ver];
            const blockEccLen = QrCode.ECC_CODEWORDS_PER_BLOCK[ecl.ordinal][ver];
            const rawCodewords = Math.floor(QrCode.getNumRawDataModules(ver) / 8);
            const numShortBlocks = numBlocks - rawCodewords % numBlocks;
            const shortBlockLen = Math.floor(rawCodewords / numBlocks);
            let blocks = [];
            const rsDiv = QrCode.reedSolomonComputeDivisor(blockEccLen);
            for(let i = 0, k = 0; i < numBlocks; i++){
                let dat = data.slice(k, k + shortBlockLen - blockEccLen + (i < numShortBlocks ? 0 : 1));
                k += dat.length;
                const ecc = QrCode.reedSolomonComputeRemainder(dat, rsDiv);
                if (i < numShortBlocks) dat.push(0);
                blocks.push(dat.concat(ecc));
            }
            let result = [];
            for(let i = 0; i < blocks[0].length; i++){
                blocks.forEach((block, j)=>{
                    if (i != shortBlockLen - blockEccLen || j >= numShortBlocks) result.push(block[i]);
                });
            }
            assert(result.length == rawCodewords);
            return result;
        }
        drawCodewords(data) {
            if (data.length != Math.floor(QrCode.getNumRawDataModules(this.version) / 8)) throw new RangeError("Invalid argument");
            let i = 0;
            for(let right = this.size - 1; right >= 1; right -= 2){
                if (right == 6) right = 5;
                for(let vert = 0; vert < this.size; vert++){
                    for(let j = 0; j < 2; j++){
                        const x = right - j;
                        const upward = (right + 1 & 2) == 0;
                        const y = upward ? this.size - 1 - vert : vert;
                        if (!this.isFunction[y][x] && i < data.length * 8) {
                            this.modules[y][x] = getBit(data[i >>> 3], 7 - (i & 7));
                            i++;
                        }
                    }
                }
            }
            assert(i == data.length * 8);
        }
        applyMask(mask) {
            if (mask < 0 || mask > 7) throw new RangeError("Mask value out of range");
            for(let y = 0; y < this.size; y++){
                for(let x = 0; x < this.size; x++){
                    let invert;
                    switch(mask){
                        case 0:
                            invert = (x + y) % 2 == 0;
                            break;
                        case 1:
                            invert = y % 2 == 0;
                            break;
                        case 2:
                            invert = x % 3 == 0;
                            break;
                        case 3:
                            invert = (x + y) % 3 == 0;
                            break;
                        case 4:
                            invert = (Math.floor(x / 3) + Math.floor(y / 2)) % 2 == 0;
                            break;
                        case 5:
                            invert = x * y % 2 + x * y % 3 == 0;
                            break;
                        case 6:
                            invert = (x * y % 2 + x * y % 3) % 2 == 0;
                            break;
                        case 7:
                            invert = ((x + y) % 2 + x * y % 3) % 2 == 0;
                            break;
                        default:
                            throw new Error("Unreachable");
                    }
                    if (!this.isFunction[y][x] && invert) this.modules[y][x] = !this.modules[y][x];
                }
            }
        }
        getPenaltyScore() {
            let result = 0;
            for(let y = 0; y < this.size; y++){
                let runColor = false;
                let runX = 0;
                let runHistory = [
                    0,
                    0,
                    0,
                    0,
                    0,
                    0,
                    0
                ];
                for(let x = 0; x < this.size; x++){
                    if (this.modules[y][x] == runColor) {
                        runX++;
                        if (runX == 5) result += QrCode.PENALTY_N1;
                        else if (runX > 5) result++;
                    } else {
                        this.finderPenaltyAddHistory(runX, runHistory);
                        if (!runColor) result += this.finderPenaltyCountPatterns(runHistory) * QrCode.PENALTY_N3;
                        runColor = this.modules[y][x];
                        runX = 1;
                    }
                }
                result += this.finderPenaltyTerminateAndCount(runColor, runX, runHistory) * QrCode.PENALTY_N3;
            }
            for(let x = 0; x < this.size; x++){
                let runColor = false;
                let runY = 0;
                let runHistory = [
                    0,
                    0,
                    0,
                    0,
                    0,
                    0,
                    0
                ];
                for(let y = 0; y < this.size; y++){
                    if (this.modules[y][x] == runColor) {
                        runY++;
                        if (runY == 5) result += QrCode.PENALTY_N1;
                        else if (runY > 5) result++;
                    } else {
                        this.finderPenaltyAddHistory(runY, runHistory);
                        if (!runColor) result += this.finderPenaltyCountPatterns(runHistory) * QrCode.PENALTY_N3;
                        runColor = this.modules[y][x];
                        runY = 1;
                    }
                }
                result += this.finderPenaltyTerminateAndCount(runColor, runY, runHistory) * QrCode.PENALTY_N3;
            }
            for(let y = 0; y < this.size - 1; y++){
                for(let x = 0; x < this.size - 1; x++){
                    const color = this.modules[y][x];
                    if (color == this.modules[y][x + 1] && color == this.modules[y + 1][x] && color == this.modules[y + 1][x + 1]) result += QrCode.PENALTY_N2;
                }
            }
            let dark = 0;
            for (const row of this.modules)dark = row.reduce((sum, color)=>sum + (color ? 1 : 0), dark);
            const total = this.size * this.size;
            const k = Math.ceil(Math.abs(dark * 20 - total * 10) / total) - 1;
            assert(0 <= k && k <= 9);
            result += k * QrCode.PENALTY_N4;
            assert(0 <= result && result <= 2568888);
            return result;
        }
        getAlignmentPatternPositions() {
            if (this.version == 1) return [];
            else {
                const numAlign = Math.floor(this.version / 7) + 2;
                const step = Math.floor((this.version * 8 + numAlign * 3 + 5) / (numAlign * 4 - 4)) * 2;
                let result = [
                    6
                ];
                for(let pos = this.size - 7; result.length < numAlign; pos -= step)result.splice(1, 0, pos);
                return result;
            }
        }
        static getNumRawDataModules(ver) {
            if (ver < QrCode.MIN_VERSION || ver > QrCode.MAX_VERSION) throw new RangeError("Version number out of range");
            let result = (16 * ver + 128) * ver + 64;
            if (ver >= 2) {
                const numAlign = Math.floor(ver / 7) + 2;
                result -= (25 * numAlign - 10) * numAlign - 55;
                if (ver >= 7) result -= 36;
            }
            assert(208 <= result && result <= 29648);
            return result;
        }
        static getNumDataCodewords(ver, ecl) {
            return Math.floor(QrCode.getNumRawDataModules(ver) / 8) - QrCode.ECC_CODEWORDS_PER_BLOCK[ecl.ordinal][ver] * QrCode.NUM_ERROR_CORRECTION_BLOCKS[ecl.ordinal][ver];
        }
        static reedSolomonComputeDivisor(degree) {
            if (degree < 1 || degree > 255) throw new RangeError("Degree out of range");
            let result = [];
            for(let i = 0; i < degree - 1; i++)result.push(0);
            result.push(1);
            let root = 1;
            for(let i = 0; i < degree; i++){
                for(let j = 0; j < result.length; j++){
                    result[j] = QrCode.reedSolomonMultiply(result[j], root);
                    if (j + 1 < result.length) result[j] ^= result[j + 1];
                }
                root = QrCode.reedSolomonMultiply(root, 0x02);
            }
            return result;
        }
        static reedSolomonComputeRemainder(data, divisor) {
            let result = divisor.map((_)=>0);
            for (const b of data){
                const factor = b ^ result.shift();
                result.push(0);
                divisor.forEach((coef, i)=>result[i] ^= QrCode.reedSolomonMultiply(coef, factor));
            }
            return result;
        }
        static reedSolomonMultiply(x, y) {
            if (x >>> 8 != 0 || y >>> 8 != 0) throw new RangeError("Byte out of range");
            let z = 0;
            for(let i = 7; i >= 0; i--){
                z = z << 1 ^ (z >>> 7) * 0x11D;
                z ^= (y >>> i & 1) * x;
            }
            assert(z >>> 8 == 0);
            return z;
        }
        finderPenaltyCountPatterns(runHistory) {
            const n = runHistory[1];
            assert(n <= this.size * 3);
            const core = n > 0 && runHistory[2] == n && runHistory[3] == n * 3 && runHistory[4] == n && runHistory[5] == n;
            return (core && runHistory[0] >= n * 4 && runHistory[6] >= n ? 1 : 0) + (core && runHistory[6] >= n * 4 && runHistory[0] >= n ? 1 : 0);
        }
        finderPenaltyTerminateAndCount(currentRunColor, currentRunLength, runHistory) {
            if (currentRunColor) {
                this.finderPenaltyAddHistory(currentRunLength, runHistory);
                currentRunLength = 0;
            }
            currentRunLength += this.size;
            this.finderPenaltyAddHistory(currentRunLength, runHistory);
            return this.finderPenaltyCountPatterns(runHistory);
        }
        finderPenaltyAddHistory(currentRunLength, runHistory) {
            if (runHistory[0] == 0) currentRunLength += this.size;
            runHistory.pop();
            runHistory.unshift(currentRunLength);
        }








    }
    QrCode.MIN_VERSION = 1;
QrCode.MAX_VERSION = 40;
QrCode.PENALTY_N1 = 3;
QrCode.PENALTY_N2 = 3;
QrCode.PENALTY_N3 = 40;
QrCode.PENALTY_N4 = 10;
QrCode.ECC_CODEWORDS_PER_BLOCK = [
            [
                -1,
                7,
                10,
                15,
                20,
                26,
                18,
                20,
                24,
                30,
                18,
                20,
                24,
                26,
                30,
                22,
                24,
                28,
                30,
                28,
                28,
                28,
                28,
                30,
                30,
                26,
                28,
                30,
                30,
                30,
                30,
                30,
                30,
                30,
                30,
                30,
                30,
                30,
                30,
                30,
                30
            ],
            [
                -1,
                10,
                16,
                26,
                18,
                24,
                16,
                18,
                22,
                22,
                26,
                30,
                22,
                22,
                24,
                24,
                28,
                28,
                26,
                26,
                26,
                26,
                28,
                28,
                28,
                28,
                28,
                28,
                28,
                28,
                28,
                28,
                28,
                28,
                28,
                28,
                28,
                28,
                28,
                28,
                28
            ],
            [
                -1,
                13,
                22,
                18,
                26,
                18,
                24,
                18,
                22,
                20,
                24,
                28,
                26,
                24,
                20,
                30,
                24,
                28,
                28,
                26,
                30,
                28,
                30,
                30,
                30,
                30,
                28,
                30,
                30,
                30,
                30,
                30,
                30,
                30,
                30,
                30,
                30,
                30,
                30,
                30,
                30
            ],
            [
                -1,
                17,
                28,
                22,
                16,
                22,
                28,
                26,
                26,
                24,
                28,
                24,
                28,
                22,
                24,
                24,
                30,
                28,
                28,
                26,
                28,
                30,
                24,
                30,
                30,
                30,
                30,
                30,
                30,
                30,
                30,
                30,
                30,
                30,
                30,
                30,
                30,
                30,
                30,
                30,
                30
            ]
        ];
QrCode.NUM_ERROR_CORRECTION_BLOCKS = [
            [
                -1,
                1,
                1,
                1,
                1,
                1,
                2,
                2,
                2,
                2,
                4,
                4,
                4,
                4,
                4,
                6,
                6,
                6,
                6,
                7,
                8,
                8,
                9,
                9,
                10,
                12,
                12,
                12,
                13,
                14,
                15,
                16,
                17,
                18,
                19,
                19,
                20,
                21,
                22,
                24,
                25
            ],
            [
                -1,
                1,
                1,
                1,
                2,
                2,
                4,
                4,
                4,
                5,
                5,
                5,
                8,
                9,
                9,
                10,
                10,
                11,
                13,
                14,
                16,
                17,
                17,
                18,
                20,
                21,
                23,
                25,
                26,
                28,
                29,
                31,
                33,
                35,
                37,
                38,
                40,
                43,
                45,
                47,
                49
            ],
            [
                -1,
                1,
                1,
                2,
                2,
                4,
                4,
                6,
                6,
                8,
                8,
                8,
                10,
                12,
                16,
                12,
                17,
                16,
                18,
                21,
                20,
                23,
                23,
                25,
                27,
                29,
                34,
                34,
                35,
                38,
                40,
                43,
                45,
                48,
                51,
                53,
                56,
                59,
                62,
                65,
                68
            ],
            [
                -1,
                1,
                1,
                2,
                4,
                4,
                4,
                5,
                6,
                8,
                8,
                11,
                11,
                16,
                16,
                18,
                16,
                19,
                21,
                25,
                25,
                25,
                34,
                30,
                32,
                35,
                37,
                40,
                42,
                45,
                48,
                51,
                54,
                57,
                60,
                63,
                66,
                70,
                74,
                77,
                81
            ]
        ];
    qrcodegen.QrCode = QrCode;
    function appendBits(val, len, bb) {
        if (len < 0 || len > 31 || val >>> len != 0) throw new RangeError("Value out of range");
        for(let i = len - 1; i >= 0; i--)bb.push(val >>> i & 1);
    }
    function getBit(x, i) {
        return (x >>> i & 1) != 0;
    }
    function assert(cond) {
        if (!cond) throw new Error("Assertion error");
    }
    class QrSegment {



        static makeBytes(data) {
            let bb = [];
            for (const b of data)appendBits(b, 8, bb);
            return new QrSegment(QrSegment.Mode.BYTE, data.length, bb);
        }
        static makeNumeric(digits) {
            if (!QrSegment.isNumeric(digits)) throw new RangeError("String contains non-numeric characters");
            let bb = [];
            for(let i = 0; i < digits.length;){
                const n = Math.min(digits.length - i, 3);
                appendBits(parseInt(digits.substring(i, i + n), 10), n * 3 + 1, bb);
                i += n;
            }
            return new QrSegment(QrSegment.Mode.NUMERIC, digits.length, bb);
        }
        static makeAlphanumeric(text) {
            if (!QrSegment.isAlphanumeric(text)) throw new RangeError("String contains unencodable characters in alphanumeric mode");
            let bb = [];
            let i;
            for(i = 0; i + 2 <= text.length; i += 2){
                let temp = QrSegment.ALPHANUMERIC_CHARSET.indexOf(text.charAt(i)) * 45;
                temp += QrSegment.ALPHANUMERIC_CHARSET.indexOf(text.charAt(i + 1));
                appendBits(temp, 11, bb);
            }
            if (i < text.length) appendBits(QrSegment.ALPHANUMERIC_CHARSET.indexOf(text.charAt(i)), 6, bb);
            return new QrSegment(QrSegment.Mode.ALPHANUMERIC, text.length, bb);
        }
        static makeSegments(text) {
            if (text == "") return [];
            else if (QrSegment.isNumeric(text)) return [
                QrSegment.makeNumeric(text)
            ];
            else if (QrSegment.isAlphanumeric(text)) return [
                QrSegment.makeAlphanumeric(text)
            ];
            else return [
                QrSegment.makeBytes(QrSegment.toUtf8ByteArray(text))
            ];
        }
        static makeEci(assignVal) {
            let bb = [];
            if (assignVal < 0) throw new RangeError("ECI assignment value out of range");
            else if (assignVal < 1 << 7) appendBits(assignVal, 8, bb);
            else if (assignVal < 1 << 14) {
                appendBits(0b10, 2, bb);
                appendBits(assignVal, 14, bb);
            } else if (assignVal < 1000000) {
                appendBits(0b110, 3, bb);
                appendBits(assignVal, 21, bb);
            } else throw new RangeError("ECI assignment value out of range");
            return new QrSegment(QrSegment.Mode.ECI, 0, bb);
        }
        static isNumeric(text) {
            return QrSegment.NUMERIC_REGEX.test(text);
        }
        static isAlphanumeric(text) {
            return QrSegment.ALPHANUMERIC_REGEX.test(text);
        }
        constructor(mode, numChars, bitData){
            this.mode = mode;
            this.numChars = numChars;
            this.bitData = bitData;
            if (numChars < 0) throw new RangeError("Invalid argument");
            this.bitData = bitData.slice();
        }
        getData() {
            return this.bitData.slice();
        }
        static getTotalBits(segs, version) {
            let result = 0;
            for (const seg of segs){
                const ccbits = seg.mode.numCharCountBits(version);
                if (seg.numChars >= 1 << ccbits) return Infinity;
                result += 4 + ccbits + seg.bitData.length;
            }
            return result;
        }
        static toUtf8ByteArray(str) {
            str = encodeURI(str);
            let result = [];
            for(let i = 0; i < str.length; i++){
                if (str.charAt(i) != "%") result.push(str.charCodeAt(i));
                else {
                    result.push(parseInt(str.substring(i + 1, i + 3), 16));
                    i += 2;
                }
            }
            return result;
        }



    }
    QrSegment.NUMERIC_REGEX = /^[0-9]*$/;
QrSegment.ALPHANUMERIC_REGEX = /^[A-Z0-9 $%*+.\/:-]*$/;
QrSegment.ALPHANUMERIC_CHARSET = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ $%*+-./:";
    qrcodegen.QrSegment = QrSegment;
})(qrcodegen || (qrcodegen = {}));
(function(qrcodegen) {
    (function(QrCode) {
        class Ecc {






            constructor(ordinal, formatBits){
                this.ordinal = ordinal;
                this.formatBits = formatBits;
            }
        }
        Ecc.LOW = new Ecc(0, 1);
Ecc.MEDIUM = new Ecc(1, 0);
Ecc.QUARTILE = new Ecc(2, 3);
Ecc.HIGH = new Ecc(3, 2);
    QrCode.Ecc = Ecc;
    })(qrcodegen.QrCode || (qrcodegen.QrCode = {}));
})(qrcodegen || (qrcodegen = {}));
(function(qrcodegen) {
    (function(QrSegment) {
        class Mode {







            constructor(modeBits, numBitsCharCount){
                this.modeBits = modeBits;
                this.numBitsCharCount = numBitsCharCount;
            }
            numCharCountBits(ver) {
                return this.numBitsCharCount[Math.floor((ver + 7) / 17)];
            }
        }
        Mode.NUMERIC = new Mode(0x1, [
                10,
                12,
                14
            ]);
Mode.ALPHANUMERIC = new Mode(0x2, [
                9,
                11,
                13
            ]);
Mode.BYTE = new Mode(0x4, [
                8,
                16,
                16
            ]);
Mode.KANJI = new Mode(0x8, [
                8,
                10,
                12
            ]);
Mode.ECI = new Mode(0x7, [
                0,
                0,
                0
            ]);
    QrSegment.Mode = Mode;
    })(qrcodegen.QrSegment || (qrcodegen.QrSegment = {}));
})(qrcodegen || (qrcodegen = {}));
var qrcodegen;

const SHEET_ROSTER = 'QR用名簿';
const SHEET_LOG = '提出履歴';
const VALID_TYPES = ['A', 'B', 'C', 'D'];


// ==================================================
// GitHub Pages から POST を受け取る
// ==================================================

function doPost(e) {

  try {

    const data = JSON.parse(
      e.postData.contents || '{}'
    );


    // ---------- 管理キー確認 ----------

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


    // ---------- 提出登録 ----------

    const result = registerSubmission(
      data.studentId,
      data.submissionType
    );

    return jsonResponse(result);

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
// 提出登録
// ==================================================

function registerSubmission(
  studentId,
  submissionType
) {

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


  if (
    !VALID_TYPES.includes(submissionType)
  ) {

    return {
      ok: false,
      message: '提出物A〜Dを選んでください。'
    };

  }


  // 先生2人がほぼ同時に登録しても
  // 二重記録されにくくする
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


    if (!rosterSheet) {

      return {
        ok: false,
        message:
          'QR用名簿シートが見つかりません。'
      };

    }


    if (!logSheet) {

      return {
        ok: false,
        message:
          '提出履歴シートが見つかりません。'
      };

    }


    // ==================================================
    // 名簿
    //
    // A 生徒ID
    // B クラス
    // C 出席番号
    // D 氏名
    // ==================================================

    const lastRosterRow =
      rosterSheet.getLastRow();


    const roster =
      rosterSheet
        .getRange(
          2,
          1,
          lastRosterRow - 1,
          4
        )
        .getValues();


    const student =
      roster.find(row =>

        String(row[0]).trim()
          === studentId

      );


    if (!student) {

      return {
        ok: false,
        message:
          `名簿にないIDです：${studentId}`
      };

    }


    const id = student[0];
    const className = student[1];
    const number = student[2];
    const name = student[3];


    // ==================================================
    // 日付
    // ==================================================

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


    // ==================================================
    // 同じ日＋同じ生徒＋同じ提出物
    // の二重登録を防ぐ
    // ==================================================

    const lastLogRow =
      logSheet.getLastRow();


    if (lastLogRow >= 2) {

      const logs =
        logSheet
          .getRange(
            2,
            1,
            lastLogRow - 1,
            6
          )
          .getValues();


      const duplicated =
        logs.some(row => {

          const rowDate = row[1];

          const rowId =
            String(row[2]).trim();

          const rowType =
            String(row[5])
              .trim()
              .toUpperCase();


          let rowDateKey = '';


          if (
            rowDate instanceof Date
            && !isNaN(rowDate)
          ) {

            rowDateKey =
              Utilities.formatDate(
                rowDate,
                tz,
                'yyyy-MM-dd'
              );

          }


          return (
            rowDateKey === todayKey
            &&
            rowId === studentId
            &&
            rowType === submissionType
          );

        });


      if (duplicated) {

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

    }


    // ==================================================
    // 提出履歴へ追加
    // ==================================================

    logSheet.appendRow([

      now,              // タイムスタンプ
      now,              // 日付
      id,               // 生徒ID
      name,             // 氏名
      className,        // クラス
      submissionType    // A〜D

    ]);


    // ==================================================
    // この生徒の、この提出物の累計回数
    // ==================================================

    const newLastRow =
      logSheet.getLastRow();


    const allLogs =
      logSheet
        .getRange(
          2,
          1,
          newLastRow - 1,
          6
        )
        .getValues();


    const count =
      allLogs.filter(row =>

        String(row[2]).trim()
          === studentId

        &&

        String(row[5])
          .trim()
          .toUpperCase()
          === submissionType

      ).length;


    return {

      ok: true,

      studentId: id,

      name: name,

      className: className,

      number: number,

      submissionType:
        submissionType,

      count: count,

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

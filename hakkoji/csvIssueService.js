/**
 * csvIssueService.js
 * setting準拠CSVの取込→即時発行
 * 既存の performCouponIssueSettingRow / loginAndUpdateUser / authenticateUser_ を再利用。
 * 推測でRakuten APIペイロードを組まない。
 */

var CSV_MAX_ROWS = 20; // GAS6分制限 + 楽天API逐次コール上限

/**
 * フロントから google.script.run.issueFromSettingCsv(userId, password, rows) で呼ぶ。
 * @param {string} userId   ログイン中ユーザーID（CSVのuserIdは無視・なりすまし防止）
 * @param {string} password 平文パスワード
 * @param {Array}  rows     CSVパース済み行（オブジェクト配列）
 * @return {Object} {success, data:{issued, total, results}}
 */
function issueFromSettingCsv(userId, password, rows) {

  // 1) 認証（既存関数を使う。trueなら成功、文字列ならエラー）
  var authResult = authenticateUser_(userId, password);
  if (authResult !== true) {
    return { success: false, error: typeof authResult === 'string' ? authResult : '認証失敗: IDまたはパスワードが正しくありません' };
  }

  if (!rows || rows.length === 0) {
    return { success: false, error: '行がありません' };
  }
  if (rows.length > CSV_MAX_ROWS) {
    return { success: false, error: '一度に発行できるのは' + CSV_MAX_ROWS + '件までです（' + rows.length + '件）' };
  }

  var results = [];
  var okCount = 0;

  for (var i = 0; i < rows.length; i++) {
    var r = rows[i] || {};
    var rowNo = i + 1;

    // 2) 必須列チェック
    var errMsg = validateCsvRow_(r);
    if (errMsg) {
      results.push({ row: rowNo, status: 'error', message: errMsg });
      continue;
    }

    try {
      // 3) settingタブへ書き込み（既存ヘルパーを使う）
      //    couponStartDate/couponEndDate は loginAndUpdateUser が書かない列なので渡さない
      //    issueRemaining=0, lastDiscount='' にリセットされる（仕様通り）
      var slot = parseInt(r.slot, 10) || 1;
      loginAndUpdateUser(
        userId,
        password,
        r.discountValues || '',
        r.discountType   || '1',
        parseInt(r.issueCount,          10) || 0,
        parseInt(r.memberAvailMaxCount, 10) || 0,
        parseInt(r.combineFlag,         10) || 0,
        r.conditionTypeCode || '',
        r.startValue        || '',
        parseInt(r.startHour,   10) || 0,
        parseInt(r.startMinute, 10) || 0,
        parseInt(r.endHour,     10) || 23,
        parseInt(r.endMinute,   10) || 59,
        r.itemCodeList || '',
        slot
      );

      // 4) settingを読み直して row 配列を作る
      //    performCouponIssueSettingRow はシートから読むのではなく Array を受け取る。
      //    loginAndUpdateUser 直後のシート状態を取得する。
      //    MAIN_SPREADSHEET_ID は forallusers.js で定義済みのグローバル定数を使う。
      var settingSheet = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID).getSheetByName('setting');
      var settingData = settingSheet.getDataRange().getValues();
      var settingRow  = null;
      for (var s = 1; s < settingData.length; s++) {
        if (String(settingData[s][0]) === String(userId) &&
            String(settingData[s][1]) === String(slot)) {
          settingRow = settingData[s];
          break;
        }
      }
      if (!settingRow) throw new Error('setting行の取得に失敗しました (userId=' + userId + ', slot=' + slot + ')');

      // 5) couponDay = couponStartDate を Date に変換（performCouponIssueSettingRow 第2引数）
      var couponDay = new Date(r.couponStartDate);
      if (isNaN(couponDay.getTime())) throw new Error('couponStartDate の日付形式が不正です: ' + r.couponStartDate);

      // 6) 発行（既存関数をそのまま呼ぶ。楽天APIペイロードはここで組まれる）
      var res = performCouponIssueSettingRow(settingRow, couponDay);

      // 7) lastDiscount を setting に書き戻す（通常は issueCouponsForAllUsersBatch が担う処理）
      for (var s2 = 1; s2 < settingData.length; s2++) {
        if (String(settingData[s2][0]) === String(userId) &&
            String(settingData[s2][1]) === String(slot)) {
          settingSheet.getRange(s2 + 1, 4).setValue(res.newLastDiscount);
          break;
        }
      }

      results.push({
        row:           rowNo,
        status:        'issued',
        couponName:    res.finalCouponName    || '',
        couponStart:   res.couponStart        || '',
        couponEnd:     res.couponEnd          || '',
        displayFactor: res.displayFactor,
        discountUnit:  res.discountUnit       || '',
        download:      res.download
      });
      okCount++;

    } catch (e) {
      results.push({ row: rowNo, status: 'error', message: e.message });
    }
  }

  return {
    success: true,
    data: { issued: okCount, total: rows.length, results: results }
  };
}

/** 必須列チェック */
function validateCsvRow_(r) {
  if (!r.discountValues || String(r.discountValues).trim() === '')
    return 'discountValues（値引き候補）は必須です';
  if (!r.discountType || String(r.discountType).trim() === '')
    return 'discountType（1=定額/2=定率/4=送料無料）は必須です';
  if (!r.couponStartDate || String(r.couponStartDate).trim() === '')
    return 'couponStartDate（発行日 YYYY-MM-DD）は必須です';
  var d = new Date(r.couponStartDate);
  if (isNaN(d.getTime())) return 'couponStartDate の形式が不正です（例: 2025-07-01）';
  return '';
}

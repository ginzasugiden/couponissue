/**
 * csvIssueService.js
 * setting準拠CSVの取込→即時発行
 * 既存の performCouponIssueSettingRow / loginAndUpdateUser / authenticateUser_ を再利用。
 *
 * CSV列仕様（改善後）:
 *   discountValues  - 単一数値のみ（例: 15）複数値不可
 *   discountType    - 定額/定率/送料無料 または 1/2/3
 *   couponStartDate - 発行開始日 YYYY/M/D（必須・本日以降）
 *   couponEndDate   - 発行終了日 YYYY/M/D（省略時は couponStartDate と同日）
 *   issueCount, memberAvailMaxCount, combineFlag - 任意（デフォルト0/1/1）
 *   startHour/startMinute/endHour/endMinute - 任意（デフォルト0/0/23/59）
 *   conditionTypeCode, startValue, itemCodeList - 任意
 *   ※ slot は廃止（GAS側で自動採番: 既存最大+1）
 */

var CSV_MAX_ROWS = 20;

/**
 * フロントから callApi('issueFromSettingCsv', {rows}) で呼ぶ。
 */
function issueFromSettingCsv(userId, password, rows) {

  // 認証
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

    var errMsg = validateCsvRow_(r);
    if (errMsg) {
      results.push({ row: rowNo, status: 'error', message: errMsg });
      continue;
    }

    try {
      // discountType を正規化: 定額/1→'1', 定率/2→'2', 送料無料/3/4→'4'
      var discountType = normalizeDiscountType_(r.discountType);

      // slot 自動採番（対象ユーザーの既存最大スロット+1）
      var settingSheet = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID).getSheetByName('setting');
      var settingData  = settingSheet.getDataRange().getValues();
      var maxSlot = 0;
      for (var s = 1; s < settingData.length; s++) {
        if (String(settingData[s][0]).trim() === String(userId)) {
          var thisSlot = parseInt(settingData[s][1], 10) || 0;
          if (thisSlot > maxSlot) maxSlot = thisSlot;
        }
      }
      var slot = maxSlot + 1;

      // setting に書き込み（loginAndUpdateUser は lastDiscount='' にリセットする）
      loginAndUpdateUser(
        userId,
        password,
        r.discountValues || '',
        discountType,
        parseInt(r.issueCount,          10) || 0,
        parseInt(r.memberAvailMaxCount, 10) || 1,
        parseInt(r.combineFlag,         10) || 1,
        r.conditionTypeCode || '',
        r.startValue        || '',
        parseInt(r.startHour,   10) || 0,
        parseInt(r.startMinute, 10) || 0,
        parseInt(r.endHour,     10) || 23,
        parseInt(r.endMinute,   10) || 59,
        r.itemCodeList || '',
        slot
      );

      // setting を読み直して row 配列を取得
      settingData = settingSheet.getDataRange().getValues();
      var settingRow = null;
      for (var s2 = 1; s2 < settingData.length; s2++) {
        if (String(settingData[s2][0]).trim() === String(userId) &&
            String(settingData[s2][1]) === String(slot)) {
          settingRow = settingData[s2];
          break;
        }
      }
      if (!settingRow) throw new Error('setting行の取得に失敗しました (userId=' + userId + ', slot=' + slot + ')');

      // 日付パース
      var couponDay    = parseDateGAS_(r.couponStartDate);
      var couponEndDay = (r.couponEndDate && r.couponEndDate.trim())
                         ? parseDateGAS_(r.couponEndDate)
                         : null;

      // 発行（楽天APIへのPOSTは performCouponIssueSettingRow が担う）
      var res = performCouponIssueSettingRow(settingRow, couponDay, couponEndDay);

      // lastDiscount を書き戻す
      settingData = settingSheet.getDataRange().getValues();
      for (var s3 = 1; s3 < settingData.length; s3++) {
        if (String(settingData[s3][0]).trim() === String(userId) &&
            String(settingData[s3][1]) === String(slot)) {
          settingSheet.getRange(s3 + 1, 4).setValue(res.newLastDiscount);
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

/** discountType 正規化 */
function normalizeDiscountType_(val) {
  var v = String(val || '').trim();
  if (v === '定額'   || v === '1') return '1';
  if (v === '定率'   || v === '2') return '2';
  if (v === '送料無料' || v === '3' || v === '4') return '4';
  return '1';
}

/** YYYY/M/D または YYYY-M-D を Date に変換（GAS用） */
function parseDateGAS_(str) {
  str = String(str || '').trim().replace(/\//g, '-');
  var parts = str.split('-');
  if (parts.length !== 3) throw new Error('日付形式が不正です: ' + str + '（例: 2025/7/1）');
  var y = parseInt(parts[0], 10), m = parseInt(parts[1], 10) - 1, d = parseInt(parts[2], 10);
  if (isNaN(y) || isNaN(m + 1) || isNaN(d)) throw new Error('日付が不正です: ' + str);
  return new Date(y, m, d, 0, 0, 0);
}

/** CSV行バリデーション */
function validateCsvRow_(r) {
  // discountValues: 単一数値のみ
  var dv = String(r.discountValues || '').trim();
  if (!dv) return 'discountValues（値引き額）は必須です';
  if (dv.indexOf(',') >= 0) return 'discountValues は単一値のみ可（例: 15）複数値は不可';
  if (isNaN(parseInt(dv, 10))) return 'discountValues は数値で入力してください（例: 15）';

  // discountType
  var dt = String(r.discountType || '').trim();
  if (!dt) return 'discountType は必須です（定額/定率/送料無料 または 1/2/3）';
  if (!['1','2','3','4','定額','定率','送料無料'].includes(dt))
    return 'discountType の値が不正です（定額/定率/送料無料 または 1/2/3）';

  // couponStartDate: 必須・本日以降
  var sd_str = String(r.couponStartDate || '').trim();
  if (!sd_str) return 'couponStartDate（発行開始日 YYYY/M/D）は必須です';
  var startDate;
  try {
    startDate = parseDateGAS_(sd_str);
    if (isNaN(startDate.getTime())) return 'couponStartDate の形式が不正です（例: 2025/7/1）';
  } catch (e) {
    return 'couponStartDate の形式が不正です（例: 2025/7/1）';
  }
  var today = new Date(); today.setHours(0, 0, 0, 0);
  if (startDate < today) return 'couponStartDate は本日以降の日付を指定してください';

  // couponEndDate: 省略可・startDate以降
  var ed_str = String(r.couponEndDate || '').trim();
  if (ed_str) {
    try {
      var endDate = parseDateGAS_(ed_str);
      if (isNaN(endDate.getTime())) return 'couponEndDate の形式が不正です（例: 2025/7/3）';
      if (endDate < startDate) return 'couponEndDate は couponStartDate 以降の日付を指定してください';
    } catch (e) {
      return 'couponEndDate の形式が不正です（例: 2025/7/3）';
    }
  }

  return '';
}

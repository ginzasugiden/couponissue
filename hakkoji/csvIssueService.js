/**
 * csvIssueService.js
 * setting準拠CSVの取込→即時発行
 * 既存の performCouponIssueSettingRow / loginAndUpdateUser / authenticateUser_ を再利用。
 *
 * フロントから受け取るオブジェクトキー（HEADER_MAP変換済み）:
 *   discountType    - 定額/定率/送料無料 または 1/2/3/4
 *   discountFactor  - 値引き額or率（単一数値。送料無料は内部で1に固定）
 *   couponStartDate - 発行開始日 YYYY/M/D（必須・本日以降）
 *   couponEndDate   - 発行終了日 YYYY/M/D（省略時は couponStartDate と同日）
 *   startHour/startMinute/endHour/endMinute - 省略可（デフォルト0/0/23/59）
 *   issueCount/memberAvailMaxCount/combineFlag - 省略可（デフォルト0/0/0）
 *   conditionTypeCode/startValue/itemCodeList  - 省略可
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

      // discountFactor: 全角→半角変換、送料無料は強制で 1
      var discountFactor;
      if (discountType === '4') {
        discountFactor = '1';
      } else {
        discountFactor = String(parseInt(toHankakuGAS_(String(r.discountFactor || '')), 10) || 0);
      }

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

      // conditionTypeCode 正規化: CSV値 1→RS003, 2→RS004
      var rawCtc    = String(r.conditionTypeCode || '').trim();
      var ctcForApi = rawCtc === '1' ? 'RS003' : rawCtc === '2' ? 'RS004' : rawCtc;

      // クーポン名（空欄=自動生成）
      var couponName = String(r.couponName || '').trim();
      if (!couponName) couponName = buildCouponName_(r.couponStartDate, r.couponEndDate);

      // setting に書き込み（loginAndUpdateUser は lastDiscount='' にリセットする）
      loginAndUpdateUser(
        userId,
        password,
        discountFactor,       // ← r.discountValues から変更（単一値・正規化済み）
        discountType,
        parseInt(toHankakuGAS_(String(r.issueCount          || '0')), 10) || 0,
        parseInt(toHankakuGAS_(String(r.memberAvailMaxCount || '0')), 10) || 0,
        parseInt(toHankakuGAS_(String(r.combineFlag         || '0')), 10) || 0,
        ctcForApi,
        r.startValue        || '',
        parseInt(toHankakuGAS_(String(r.startHour   || '0')),  10) || 0,
        parseInt(toHankakuGAS_(String(r.startMinute || '0')),  10) || 0,
        parseInt(toHankakuGAS_(String(r.endHour     || '23')), 10) || 23,
        parseInt(toHankakuGAS_(String(r.endMinute   || '59')), 10) || 59,
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
      var extra = {
        couponName:         couponName,
        couponCaption:      String(r.couponCaption       || '').trim(),
        displayFlag:        String(r.displayFlag !== undefined && r.displayFlag !== '' ? r.displayFlag : '1'),
        couponImageUrl:     String(r.couponImageUrl      || '').trim(),
        salesTypeCondition: String(r.salesTypeCondition !== undefined && r.salesTypeCondition !== '' ? r.salesTypeCondition : '0')
      };
      var res = performCouponIssueSettingRow(settingRow, couponDay, couponEndDay, extra);

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

/** 全角数字→半角数字 */
function toHankakuGAS_(str) {
  return String(str || '').replace(/[０-９]/g, function(c) {
    return String.fromCharCode(c.charCodeAt(0) - 0xFEE0);
  });
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
  if (parts.length !== 3) throw new Error('日付形式が不正です: ' + str + '（例: 2026/7/1）');
  var y = parseInt(parts[0], 10), m = parseInt(parts[1], 10) - 1, d = parseInt(parts[2], 10);
  if (isNaN(y) || isNaN(m + 1) || isNaN(d)) throw new Error('日付が不正です: ' + str);
  return new Date(y, m, d, 0, 0, 0);
}

/** CSV行バリデーション */
function validateCsvRow_(r) {
  // discountType（先に確定させてから factor を検証）
  var dt = String(r.discountType || '').trim();
  if (!dt) return '種別は必須です（定額/定率/送料無料）';
  if (['1','2','3','4','定額','定率','送料無料'].indexOf(dt) === -1)
    return '種別が不正です（定額/定率/送料無料）';
  var discountType = normalizeDiscountType_(dt);

  // discountFactor（送料無料は検証不要）
  if (discountType !== '4') {
    var fv  = toHankakuGAS_(String(r.discountFactor || '').trim());
    var f   = parseInt(fv, 10);
    if (isNaN(f)) return discountType === '2' ? '値引き率は整数で入力してください' : '値引き額は整数で入力してください';
    if (discountType === '1' && (f < 1 || f > 999999999)) return '値引き額は1〜999999999の整数で入力してください';
    if (discountType === '2' && (f < 1 || f > 99))        return '値引き率は1〜99の整数で入力してください';
  }

  // couponStartDate: 必須・本日以降
  var sd_str = String(r.couponStartDate || '').trim();
  if (!sd_str) return 'couponStartDate（発行開始日 YYYY/M/D）は必須です';
  var startDate;
  try {
    startDate = parseDateGAS_(sd_str);
    if (isNaN(startDate.getTime())) return 'couponStartDate の形式が不正です（例: 2026/7/1）';
  } catch (e) {
    return 'couponStartDate の形式が不正です（例: 2026/7/1）';
  }
  var today = new Date(); today.setHours(0, 0, 0, 0);
  if (startDate < today) return 'couponStartDate は本日以降の日付を指定してください';

  // couponEndDate: 省略可・startDate以降
  var ed_str = String(r.couponEndDate || '').trim();
  if (ed_str) {
    try {
      var endDate = parseDateGAS_(ed_str);
      if (isNaN(endDate.getTime())) return 'couponEndDate の形式が不正です（例: 2026/7/3）';
      if (endDate < startDate) return 'couponEndDate は couponStartDate 以降の日付を指定してください';
    } catch (e) {
      return 'couponEndDate の形式が不正です（例: 2026/7/3）';
    }
  }

  return '';
}

/** クーポン名自動生成: 同日=「M/D SALE！クーポン」, 別日=「M/D-M/D SALE！クーポン」 */
function buildCouponName_(startStr, endStr) {
  var sParts   = String(startStr || '').trim().replace(/-/g, '/').split('/');
  var eTrimmed = String(endStr   || '').trim();
  var eParts   = eTrimmed ? eTrimmed.replace(/-/g, '/').split('/') : sParts;
  var sMonth = parseInt(sParts[1], 10), sDay = parseInt(sParts[2], 10);
  var eMonth = parseInt(eParts[1], 10), eDay = parseInt(eParts[2], 10);
  if (sMonth === eMonth && sDay === eDay) return sMonth + '/' + sDay + ' SALE！クーポン';
  return sMonth + '/' + sDay + '-' + eMonth + '/' + eDay + ' SALE！クーポン';
}

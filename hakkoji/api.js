/**
 * api.js
 * GitHub Pages フロントエンド用 JSON API レイヤー。
 * 既存サーバー関数を一切書き換えず、ここから呼ぶだけ。
 * CORS preflight を避けるため Content-Type は付けない（text/plain 送信）。
 */

function doPost(e) {
  var result;
  try {
    var req = JSON.parse(e.postData.contents);
    var action = req.action || '';
    var userId = req.userId || '';
    var pw     = req.pw     || '';

    if (action === 'getHistory') {
      // 認証不要：更新履歴は公開情報
      result = { ok: true, data: getHistory() };

    } else if (action === 'getActiveCoupons') {
      // 認証不要：有効クーポン一覧は公開情報
      result = { ok: true, data: getActiveCouponsData_() };

    } else if (action === 'getUserSettingsList') {
      // getUserSettingsList は内部で authenticateUser_ を呼ぶ
      // 認証失敗時は文字列を返すのでそれで判別する
      var data = getUserSettingsList(userId, pw);
      if (typeof data === 'string') {
        result = { ok: false, error: data };
      } else {
        result = { ok: true, data: data };
      }

    } else if (action === 'getUserSettings') {
      var authResult = authenticateUser_(userId, pw);
      if (authResult !== true) {
        result = { ok: false, error: typeof authResult === 'string' ? authResult : '認証失敗' };
      } else {
        // data: object (設定あり) or '設定が見つかりませんでした' (文字列)
        var data = getUserSettings(userId, pw, req.slot);
        result = { ok: true, data: data };
      }

    } else if (action === 'loginAndUpdateUser') {
      // 先に認証してから loginAndUpdateUser を呼ぶ
      var authResult = authenticateUser_(userId, pw);
      if (authResult !== true) {
        result = { ok: false, error: typeof authResult === 'string' ? authResult : '認証失敗' };
      } else {
        var ret = loginAndUpdateUser(
          userId, pw,
          req.discountValues, req.discountType,
          req.issueCount,     req.memberAvailMaxCount,
          req.combineFlag,    req.conditionTypeCode,
          req.startValue,     req.startHour, req.startMinute,
          req.endHour,        req.endMinute,
          req.itemCodeList,   req.slot
        );
        result = { ok: true, data: ret };
      }

    } else if (action === 'issueFromSettingCsv') {
      // issueFromSettingCsv は内部で認証する
      var ret = issueFromSettingCsv(userId, pw, req.rows);
      if (!ret.success) {
        result = { ok: false, error: ret.error };
      } else {
        result = { ok: true, data: ret.data };
      }

    } else if (action === 'shippingCouponIssue') {
      // 元は認証なしだったが、楽天API呼出のため認証を追加
      var authResult = authenticateUser_(userId, pw);
      if (authResult !== true) {
        result = { ok: false, error: typeof authResult === 'string' ? authResult : '認証失敗' };
      } else {
        var ret = shippingCouponIssue(
          req.couponStartDate, req.couponEndDate,
          req.issueCount,      req.memberAvailMaxCount,
          req.combineFlag,     req.couponImage,
          req.discountType,    req.discountFactor, req.itemType
        );
        result = { ok: true, data: ret };
      }

    } else {
      result = { ok: false, error: '不明なaction: ' + action };
    }

  } catch (err) {
    result = { ok: false, error: String(err) };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

/** ActiveCoupons シートの 2D array を返す（doGet テンプレート注入の代替） */
function getActiveCouponsData_() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('ActiveCoupons');
  if (!sheet) return [];
  return sheet.getDataRange().getValues();
}

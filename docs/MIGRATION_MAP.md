# 発行侍 google.script.run → JSON API 移行対応表

| HTML内の google.script.run 呼び出し | 既存サーバー関数 | 引数 | 戻り値 | 機密(licenseKey/serviceSecret)を含むか | 移行後 action名 |
|---|---|---|---|---|---|
| mainichi.html: `getUserSettingsList(uid, pw)` | `設定確認.js: getUserSettingsList()` | userId, password | Array of setting objects（空配列も可）or 認証エラー文字列 | なし（コイン枚数は含むが非機密） | `getUserSettingsList` |
| mainichi.html: `getUserSettings(uid, pw, slot)` | `設定確認.js: getUserSettings()` | userId, password, slotOpt | setting object or `'設定が見つかりませんでした'`（文字列） | なし | `getUserSettings` |
| mainichi.html: `loginAndUpdateUser(uid, pw, ...13params)` | `CP設定登録.js: loginAndUpdateUser()` | userId, password, discountValues, discountType, issueCount, memberAvailMaxCount, combineFlag, conditionTypeCode, startValue, startHour, startMinute, endHour, endMinute, itemCodeList, slotOpt | string `'認証成功: 登録の内容を更新しました'` or 認証エラー文字列 | なし | `loginAndUpdateUser` |
| mainichi.html: `getHistory()` | `history.js: getHistory()` | なし | HTML文字列 `<ul><li>...</li></ul>` | なし | `getHistory` |
| mainichi.html: `issueFromSettingCsv(uid, pw, rows)` | `csvIssueService.js: issueFromSettingCsv()` | userId, password, rows[] | `{success, data:{issued, total, results}}` | なし（楽天APIとの通信はGAS内で完結） | `issueFromSettingCsv` |
| freeship.html: `shippingCouponIssue(...9params)` | `freeshipCP.js: shippingCouponIssue()` | couponStartDate, couponEndDate, issueCount, memberAvailMaxCount, combineFlag, couponImage, discountType, discountFactor, itemType | string（成功/エラーメッセージ） | なし（serviceSecret/licenseKeyはGAS PropertiesServiceのみ） | `shippingCouponIssue` |
| index.html: `<? ?>` テンプレートでデータ注入 | doGet内で `SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ActiveCoupons")` | - | 2D array（ヘッダー行+データ行） | なし | `getActiveCoupons` |

## 機密判定サマリー
- `licenseKey` / `serviceSecret` は `freeshipCP.js` 内で `PropertiesService.getScriptProperties()` から取得し、APIレスポンスに含めない ✓
- `forallusers.js` / `設定確認.js` での認証情報も同様にシート内のみで処理 ✓
- GSDコイン枚数（download）は `getUserSettings` / `getUserSettingsList` のレスポンスに含まれるが非機密扱い ✓

## 認証方針
- `getHistory` / `getActiveCoupons`：認証不要（公開情報）
- 上記以外の全action：`req.userId` / `req.pw` で `authenticateUser_()` または既存関数の内部認証を通す
- フロントは `sessionStorage.getItem('auth')` → `{userId, pw}` を全callApiに自動付与

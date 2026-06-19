// ───────── 設定 ───────── forallusers.gs
const MAIN_SPREADSHEET_ID = '1xKADeobhVYiBs7CGwVtoVRXjkRG4oG7AtRxBjpAwc48';  // user／setting を含む 
const API_KEY_SHEET_ID    = '1iYeV2SbOVoRH8Qjm2d1w5tWmhlE_zcc-yO1tDSLN7Rk';  // api_key シート 
//const REPORT_EMAIL        = 'tokyoflowercoltd+coupon@gmail.com';          // 管理者宛メール
// ──────────────────────────

function issueCouponsForAllUsersBatch() {
  const ss       = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
  const shtUser  = ss.getSheetByName('user');
  const shtSet   = ss.getSheetByName('setting');
  const allUser  = shtUser.getDataRange().getValues();
  const allSet   = shtSet.getDataRange().getValues();
  const ssApi    = SpreadsheetApp.openById(API_KEY_SHEET_ID);
  const apiData  = ssApi.getSheetByName('api_key').getDataRange().getValues();

  // 明日の日付文字列
  const today     = new Date();
  const tomorrow  = new Date(today);
  tomorrow.setDate(today.getDate() + 1);
  const tomorrowStr = Utilities.formatDate(tomorrow, 'Asia/Tokyo', 'yyyy-MM-dd');

  // userId→残り issue
  const userMap = {};
  for (let i = 1; i < allUser.length; i++) {
    const uid = String(allUser[i][0]||'').trim();
    if (!uid) continue;
    userMap[uid] = parseInt(allUser[i][1], 10) || 0;
  }

  // 管理者向けまとめログ
  const globalLogs = [
    'お世話になっております。',
    `本日(${tomorrowStr}分)のクーポン一括発行結果を以下にご報告いたします。`,
    ''
  ];

  // ユーザー向けログ保管と、GSDコイン消費情報
  const userLogs     = {};                  // userId→[ログ行]
  const toppedUpInfo = {};                  // userId→{ coins:残コイン数 }
  
  // ★★ 追加：各ユーザーのissue値を既に減算したかを記録 ★★
  const userIssueProcessed = {};            // userId→true（既に処理済み）

  // 設定シート2行目以降をループ
  for (let i = 1; i < allSet.length; i++) {
    const row    = allSet[i];
    const userId = String(row[0]||'').trim();
    if (!userId) continue;

    // ユーザーログ初期化（ヘッダ行）
    if (!userLogs[userId]) {
      userLogs[userId] = [
        'お世話になっております。',
        `本日(${tomorrowStr}分)のクーポン発行結果を以下にご報告いたします。`,
        ''
      ];
    }

    let remaining = userMap[userId] || 0;
    let didTopUp  = false;
    let newCoinCount = null;

    // ── GSDコイン補填 ──
    if (remaining === 0) {
      for (let j = 1; j < apiData.length; j++) {
        if (String(apiData[j][0]).trim() === userId) {
          const download = parseInt(apiData[j][4], 10) || 0;
          if (download > 0) {
            // コインを1枚消費
            ssApi.getSheetByName('api_key')
                 .getRange(j+1, 5)
                 .setValue(download - 1);
            remaining = 2;
            didTopUp  = true;
            newCoinCount = download - 1;  // 消費後のコイン数
          }
          break;
        }
      }
    }

    // コイン不足
    if (remaining === 0) {
      const api = getApiRow(userId);
      sendCoinPurchasePromptMail(api.email);
      globalLogs.push(`・${userId}：GSDコイン不足により発行せず`);
      Utilities.sleep(1000);
      continue;
    }

    // ── クーポン発行 ──
    try {
      const res = performCouponIssueSettingRow(row, tomorrow);

      // ★★ 修正：ユーザーごとに初回のみ残回数を-1 ★★
      if (!userIssueProcessed[userId]) {
        remaining--;
        const uRow = allUser.findIndex(r => String(r[0]).trim() === userId) + 1;
        if (uRow > 1) shtUser.getRange(uRow, 2).setValue(remaining);
        userMap[userId] = remaining;
        userIssueProcessed[userId] = true;  // 処理済みフラグを立てる
      }

      // ② settingシート：lastDiscount更新
      shtSet.getRange(i+1, 4).setValue(res.newLastDiscount);

      // 管理者ログ
      const msg = `・${res.finalCouponName} を発行 (${res.displayFactor}${res.discountUnit})`;
      globalLogs.push(msg);

      // ユーザーログ：クーポン行
      userLogs[userId].push(msg);

      // トップアップ発生なら記録
      if (didTopUp) {
        toppedUpInfo[userId] = {
          coins: newCoinCount
        };
      }

    } catch (e) {
      const errMsg = `・${userId}：エラー→${e.message}`;
      globalLogs.push(errMsg);
      userLogs[userId].push(errMsg);
    }

    Utilities.sleep(1000);
  }

  // ■ 管理者向けメール
  globalLogs.push('', '以上、よろしくお願いいたします。');
  MailApp.sendEmail({
    to:      REPORT_EMAIL,
    subject: '【通知】クーポン一括発行結果',
    body:    globalLogs.join('\n')
  });

  // ■ 各ユーザー向けメール
  for (const userId in userLogs) {
    const api     = getApiRow(userId);
    const lines   = userLogs[userId];

    // ここで「GSDコイン消費案内」を挿入
    if (toppedUpInfo[userId]) {
      const info = toppedUpInfo[userId];
      lines.push(
        '',  // 空行
        '※今回、GSDコインを1枚消費してクーポンを発行しました。',
        `   発行後の残り発行可能回数は ${userMap[userId]} 回です。`,
        `GSDコイン数：${info.coins} 枚`,
        '',
        'GSDコインのご購入はこちら： https://x.gd/hOMBS',
        ''
      );
    }

    // 最後に締めの挨拶
    lines.push(
      '毎日最新のクーポンでお客様に笑顔をお届けできるよう努めております。',
      'ご不明点などございましたらお気軽にお問い合わせください。'
    );

    MailApp.sendEmail({
      to:      api.email,
      cc:      REPORT_EMAIL,
      subject: '【通知】クーポン発行結果',
      body:    lines.join('\n')
    });
  }
}

/**
 * １行分の設定データからクーポンを発行し、
 * newLastDiscount などを返す
 * @param {Array} row        setting シートの１行
 * @param {Date}  couponDay  発行有効日の Date オブジェクト
 */
function performCouponIssueSettingRow(row, couponDay, couponEndDay) {
  const userId       = row[0];
  const apiIssueCnt  = parseInt(row[6], 10) || 0;
  const lastDisc     = parseInt(row[3], 10) || 0;
  const optsRaw      = String(row[4] || '');
  const discountMap  = {'1':'定額値引き','2':'定率値引き','4':'送料無料'};
  const discountType = discountMap[String(row[5])] || '定額値引き';
  const memberAvail  = parseInt(row[7], 10) || 0;
  const combineFlag  = row[8];
  const condCode     = row[9];
  const startValue   = row[10];
  const sh = parseInt(row[11],10), sm = parseInt(row[12],10);
  const eh = parseInt(row[13],10), em = parseInt(row[14],10);

  // --- 有効期間の開始/終了日時を計算 ---
  const base    = couponDay;
  const endBase = couponEndDay || base;
  const startDt = new Date(base.getFullYear(), base.getMonth(), base.getDate(), sh, sm, 0);
  let   endDt   = new Date(endBase.getFullYear(), endBase.getMonth(), endBase.getDate(), eh, em, 59);
  if (!couponEndDay && endDt <= startDt) endDt.setDate(endDt.getDate()+1);
  const couponStart = Utilities.formatDate(startDt, 'Asia/Tokyo', "yyyy-MM-dd'T'HH:mm:ss'+09:00'");
  const couponEnd   = Utilities.formatDate(endDt,   'Asia/Tokyo', "yyyy-MM-dd'T'HH:mm:ss'+09:00'");

  // --- 認証情報取得 ---
  const api     = getApiRow(userId);
  const license = decode(api.license);
  const secret  = decode(api.secret);
  const auth    = `ESA ${Utilities.base64Encode(secret + ':' + license)}`;
  const download= parseInt(api.download, 10) || 0;

  // --- 割引額／タイプ ---
  let discountNum, itemType;
  if (discountType === '送料無料') {
    discountNum = 1; itemType = 5;
  } else {
    const opts = optsRaw.split(',')
      .map(n=>parseInt(n.trim(),10))
      .filter(n=>!isNaN(n)&&n!== lastDisc);
    discountNum = opts[Math.floor(Math.random()*opts.length)]||0;
    itemType    = 4;
  }
  const discountUnit  = (discountType==='定率値引き') ? '%OFF' : '円OFF';
  const displayFactor = (discountType==='送料無料') ? '' : discountNum;

  // --- 商品指定の有無によって名前を切り替え ---
  const rawList   = String(row[17]||'');
  const itemCodes = rawList.split(/\s*,\s*/).filter(s=>s);
  // --- クーポン名生成部分の修正 ---
  const mm = base.getMonth() + 1;
  const dd = base.getDate();
  // 「商品限定」か「店内全品」を先に決める
  const typeLabel = itemCodes.length > 0 ? '商品限定' : '店内全品';
  // ご要望のフォーマット："7月16日SALE！ 商品限定 250円OFF"
  const finalCouponName = `${mm}月${dd}日SALE！ ${typeLabel} ${displayFactor}${discountUnit}`;

  // --- items XML ---
  const itemTypeFinal = itemCodes.length===0
    ? itemType
    : itemCodes.length===1 ? 1 : 3;
  const itemsXml = itemCodes.length
    ? '<items>' + itemCodes.map(c=>`<item><itemUrl>${c}</itemUrl></item>`).join('') + '</items>'
    : '<items/>';

  // --- otherConditions XML ---
  let otherXml = '<otherConditions/>';
  if (condCode && startValue) {
    otherXml =
      `<otherConditions><otherCondition>` +
        `<conditionTypeCode>${condCode}</conditionTypeCode>` +
        `<startValue>${startValue}</startValue>` +
      `</otherCondition></otherConditions>`;
  }

  // --- API POST ---
  const xml =
    `<?xml version="1.0" encoding="UTF-8"?>` +
    `<request><couponIssueRequest><coupon>` +
      `<couponName>${finalCouponName}</couponName>` +
      `<couponCaption>${mm}月${dd}日SALEクーポン！</couponCaption>` +
      `<couponStartDate>${couponStart}</couponStartDate>` +
      `<couponEndDate>${couponEnd}</couponEndDate>` +
      `<issueCount>${apiIssueCnt}</issueCount>` +
      `<itemType>${itemTypeFinal}</itemType>` +
      `<discountType>${discountType==='定率値引き'?2:1}</discountType>` +
      `<discountFactor>${discountNum}</discountFactor>` +
      `<memberAvailMaxCount>${memberAvail}</memberAvailMaxCount>` +
      `<purchaseHistoryCond><type>0</type></purchaseHistoryCond>` +
      `<multiRankCond><rankCond>0</rankCond></multiRankCond>` +
      `<genderCond>NONE</genderCond>` +
      `<ageRangeCond><lowerBound>0</lowerBound><upperBound>0</upperBound></ageRangeCond>` +
      `<birthmonthCond>0</birthmonthCond>` +
      `<multiPrefectureCond><prefectureCond>NONE</prefectureCond></multiPrefectureCond>` +
      `<combineFlag>${combineFlag}</combineFlag>` +
      `<displayFlag>1</displayFlag>` +
      itemsXml +
      otherXml +
    `</coupon></couponIssueRequest></request>`;

  const res = UrlFetchApp.fetch('https://api.rms.rakuten.co.jp/es/1.0/coupon/issue', {
    method:      'post',
    contentType: 'text/xml',
    payload:     xml,
    headers:     { Authorization: auth },
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) {
    throw new Error(`HTTP ${res.getResponseCode()}`);
  }

  return {
    newLastDiscount: discountNum,
    download,
    finalCouponName,
    displayFactor,
    discountUnit,
    couponStart,
    couponEnd,
    itemCodes
  };
}

/**
 * api_key シートからユーザー行を取得
 */
function getApiRow(userId) {
  const rows = SpreadsheetApp
    .openById(API_KEY_SHEET_ID)
    .getSheetByName('api_key')
    .getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === userId) {
      return {
        license:  rows[i][2],
        secret:   rows[i][3],
        download: rows[i][4],
        email:    rows[i][8]
      };
    }
  }
  throw new Error('API認証情報が見つかりません: ' + userId);
}

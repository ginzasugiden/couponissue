/**
 * debug.js（一時ファイル）
 * 販売方法条件（RS002）追加のドライラン検証用。
 * buildConditionXmlBlocks_(row) は API 呼び出しを含まない純粋関数なので、
 * この関数を実行しても実際のクーポン発行は一切発生しない。
 * 検証完了後、本ファイルは削除して push すること。
 */
function dryRunSalesMethod() {
  function mkRow(overrides) {
    const row = new Array(29).fill('');
    Object.keys(overrides).forEach(k => { row[k] = overrides[k]; });
    return row;
  }

  const cases = [
    { label: 'row[28]=undefined（毎日発行経路想定）', row: mkRow({}) },
    { label: "salesMethod='0'（通常購入のみ）", row: mkRow({ 28: '0' }) },
    { label: "salesMethod='1'（定期購入のみ）+ gender/age/birth 指定あり", row: mkRow({ 23: '男性', 24: '20', 25: '40', 26: '5', 28: '1' }) },
    { label: "salesMethod='99'（指定しない）", row: mkRow({ 28: '99' }) },
    { label: "salesMethod='0' + RS003(金額条件1000円)", row: mkRow({ 9: 'RS003', 10: '1000', 28: '0' }) },
    { label: "salesMethod='99' + RS003(金額条件1000円)", row: mkRow({ 9: 'RS003', 10: '1000', 28: '99' }) },
    { label: '会員ランク=レギュラー + salesMethod=1 + gender/age指定あり', row: mkRow({ 18: 'レギュラー', 23: '女性', 24: '30', 25: '50', 28: '1' }) },
    { label: '会員ランク=シルバー + salesMethod=0 + gender指定あり', row: mkRow({ 18: 'シルバー', 23: '男性', 28: '0' }) }
  ];

  cases.forEach(c => {
    const r = buildConditionXmlBlocks_(c.row);
    const combined = r.otherXml + r.rankXml + r.purchaseHistoryXml + r.genderXml + r.ageXml + r.birthXml + r.prefXml;
    Logger.log('=== ' + c.label + ' ===');
    Logger.log('otherXml: ' + r.otherXml);
    Logger.log('rankXml: ' + r.rankXml);
    Logger.log('purchaseHistoryXml: ' + r.purchaseHistoryXml);
    Logger.log('genderXml: ' + r.genderXml);
    Logger.log('ageXml: ' + r.ageXml);
    Logger.log('birthXml: ' + r.birthXml);
    Logger.log('prefXml: ' + r.prefXml);
    Logger.log('undefined混入チェック: ' + (/undefined/.test(combined) ? 'NG（undefinedが混入）' : 'OK'));
    Logger.log('');
  });

  Logger.log('dryRunSalesMethod 完了（API呼び出しなし・クーポン未発行）');
}

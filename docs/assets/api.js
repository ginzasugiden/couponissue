/**
 * GAS JSON API 共通ヘルパー
 * Content-Type を付けない（text/plain 送信）→ CORS preflight が発生しない
 * sessionStorage の {userId, pw} を全リクエストに自動付与する
 */
async function callApi(action, params) {
  params = params || {};
  const auth = JSON.parse(sessionStorage.getItem('auth') || '{}');
  const body = Object.assign({ action }, params, auth);
  const res = await fetch(GAS_EXEC_URL, {
    method: 'POST',
    body: JSON.stringify(body)
    // Content-Type は設定しない（preflight 回避）
  });
  const json = await res.json();
  if (!json.ok) throw new Error(json.error || 'API error');
  return json.data;
}

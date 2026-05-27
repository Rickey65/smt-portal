/* 통장 거래내역 파싱 + 자동 분개 매칭 (2026-05-27)
   12컬럼 표준 (기업은행 인터넷뱅킹 다운로드 양식)
*/
(function(global){
  'use strict';

  let DICT = null;       // bank_keyword_accounts.json
  let CLIENTS = null;    // 거래처 마스터
  let ALIASES = null;    // 별칭 사전

  async function loadDict(){
    if (DICT) return DICT;
    try {
      const r = await fetch('data/shared/bank_keyword_accounts.json?t='+Date.now());
      if (r.ok) DICT = await r.json();
    } catch(e){ console.warn('bank_keyword_accounts.json load fail', e); }
    return DICT;
  }

  // XLSX 통장 파일 파싱 → 거래 배열
  // 컬럼: [거래일시, 출금, 입금, 잔액, 메모, 계좌, 상대은행, _, 거래종류, _, _, 실제이름]
  function parseFromWorkbook(wb){
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const range = XLSX.utils.decode_range(sheet['!ref']);
    const txs = [];
    for (let r = range.s.r; r <= range.e.r; r++){
      const row = [];
      for (let c = 0; c < 12; c++){
        const cell = sheet[XLSX.utils.encode_cell({r, c})];
        row.push(cell ? cell.v : '');
      }
      if (!row[0]) continue;
      // 헤더 스킵 (첫 행이 "거래일시" 같은 텍스트)
      if (typeof row[0] === 'string' && (row[0].includes('거래') || row[0].includes('일시'))) continue;
      const dt = String(row[0]);
      const dateMatch = dt.match(/(\d{4})[\-\/\.](\d{1,2})[\-\/\.](\d{1,2})/);
      const date = dateMatch ? `${dateMatch[1]}-${dateMatch[2].padStart(2,'0')}-${dateMatch[3].padStart(2,'0')}` : dt.slice(0,10);
      txs.push({
        일시: dt,
        일자: date,
        출금: +row[1] || 0,
        입금: +row[2] || 0,
        잔액: +row[3] || 0,
        메모: String(row[4] || '').trim(),
        상대은행: String(row[6] || '').trim(),
        거래종류: String(row[8] || '').trim(),
        실제이름: String(row[11] || '').trim(),
      });
    }
    return txs;
  }

  // 키워드 매칭 — 메모 우선, 실제이름 폴백
  function matchKeyword(tx){
    if (!DICT) return null;
    const memoKey = (tx.메모 || '').slice(0,20);
    const realKey = (tx.실제이름 || '').slice(0,20);
    return (DICT.by_memo && DICT.by_memo[memoKey])
        || (DICT.by_real_name && DICT.by_real_name[realKey])
        || null;
  }

  // 분개 자동 생성
  function generateEntry(tx, opts = {}){
    opts = Object.assign({clients:[], receivable:[]}, opts);
    const match = matchKeyword(tx);
    const isIn = tx.입금 > 0;
    const amt = isIn ? tx.입금 : tx.출금;

    // 신뢰도 결정
    let confidence = 'low'; // 🔴
    if (match) {
      const cnt = match.건수 || 0;
      if (cnt >= 5) confidence = 'high';        // 🟢
      else if (cnt >= 2) confidence = 'mid';    // 🟡
      else confidence = 'low';                   // 🔴
    }

    // 추천 계정 (Top 2)
    let recAccts = match ? (match.추천계정 || []).slice(0,3).map(([a,c]) => ({계정:a, 빈도:c})) : [];
    let recCli = match ? (match.추천거래처 || []).slice(0,2).map(([c,n]) => ({거래처:c, 빈도:n})) : [];
    let recMemo = match ? (match.추천적요 || []).slice(0,2).map(([m,n]) => ({적요:m, 빈도:n})) : [];

    // 외상매출금 입금 매칭 (입금 + 거래처 매칭 + 미수채권 잔액)
    let receivableMatch = null;
    if (isIn && opts.receivable && recCli.length > 0) {
      const cli = recCli[0].거래처;
      const r = opts.receivable.find(x => (x.거래처명 || '') === cli);
      if (r) {
        const balance = +(r.채권_잔액 || r.미수금 || 0);
        receivableMatch = {거래처: cli, 잔액: balance};
      }
    }

    // 자동 분개 — 차변/대변 결정
    let chae, dae, memo, mainAcct;
    if (isIn) {
      // 입금 — 차변 보통예금 / 대변 외상매출금 (또는 추천 매출/기타)
      chae = [{계정:'보통예금', 코드:'103', 금액:amt}];
      // 대변 = 외상매출금 (수금분개) 또는 추천 매출계정
      if (receivableMatch || (recAccts[0] && recAccts[0].계정.includes('외상매출'))) {
        dae = [{계정:'외상매출금', 코드:'108', 금액:amt}];
      } else if (recAccts.length) {
        dae = [{계정:recAccts[0].계정, 코드:'', 금액:amt}];
      } else {
        dae = [{계정:'외상매출금', 코드:'108', 금액:amt}];
      }
      mainAcct = dae[0].계정;
    } else {
      // 출금 — 차변 비용 / 대변 보통예금
      const acct = recAccts.length ? recAccts[0].계정 : '미지급비용';
      chae = [{계정:acct, 코드:'', 금액:amt}];
      dae = [{계정:'보통예금', 코드:'103', 금액:amt}];
      mainAcct = chae[0].계정;
    }
    memo = recMemo.length ? recMemo[0].적요 : (tx.메모 || tx.실제이름 || '');

    return {
      _tx: tx,
      신뢰도: confidence,
      차변: chae,
      대변: dae,
      적요: memo,
      거래처: recCli.length ? recCli[0].거래처 : (tx.실제이름 || tx.메모 || ''),
      주_계정: mainAcct,
      추천계정: recAccts,
      추천거래처: recCli,
      추천적요: recMemo,
      미수매칭: receivableMatch,
      _match건수: match ? match.건수 : 0
    };
  }

  global.BankParser = {
    loadDict, parseFromWorkbook, matchKeyword, generateEntry
  };
})(window);

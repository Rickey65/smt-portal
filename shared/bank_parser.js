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

  // 키워드 매칭 — 카테고리 룰(명확) → 정확 → 부분
  function matchKeyword(tx){
    if (!DICT) return null;
    const memoKey = (tx.메모 || '').slice(0,20);
    const realKey = (tx.실제이름 || '').slice(0,20);
    const fullMemo = (tx.메모 || '') + ' ' + (tx.실제이름 || '');

    // 1순위: 카테고리 룰 (주유소·충전소·통신·식당 등 명확)
    const rules = DICT._category_rules || {};
    for (const [acct, kws] of Object.entries(rules)) {
      for (const kw of kws) {
        if (fullMemo.includes(kw)) {
          return {
            건수: 1,
            추천계정: [[acct, 999]],  // 빈도 999 = 룰 우선
            추천거래처: tx.실제이름 ? [[tx.실제이름, 1]] : (tx.메모 ? [[tx.메모, 1]] : []),
            추천적요: [[`${tx.메모 || tx.실제이름}`, 1]],
            _ruleMatch: kw, _ruleCategory: acct
          };
        }
      }
    }

    // 2순위: 정확 매칭 (학습 데이터)
    let match = (DICT.by_memo && DICT.by_memo[memoKey])
             || (DICT.by_real_name && DICT.by_real_name[realKey]);
    if (match) return match;

    // 3순위: 부분 매칭 — 단, 키가 5자 이상이어야 (오매칭 방지)
    if (DICT.by_memo) {
      for (const k of Object.keys(DICT.by_memo)) {
        if (k.length >= 5 && (memoKey.includes(k) || (k.includes(memoKey) && memoKey.length >= 4))) {
          return DICT.by_memo[k];
        }
      }
    }
    if (DICT.by_real_name) {
      for (const k of Object.keys(DICT.by_real_name)) {
        if (k.length >= 5 && (realKey.includes(k) || (k.includes(realKey) && realKey.length >= 4))) {
          return DICT.by_real_name[k];
        }
      }
    }
    return null;
  }

  // 체크카드 번호 → 사원 매핑 (학습된 사원 이름 + 카드 사용자 추정)
  function getEmployeeFromTx(tx){
    if (!DICT) return null;
    const empNames = DICT._employee_names || {};
    // 메모/실제이름에 사원 이름 직접 포함
    const fullText = (tx.메모 || '') + ' ' + (tx.실제이름 || '');
    for (const nm of Object.keys(empNames)) {
      if (fullText.includes(nm)) return nm;
    }
    return null;
  }

  // 통장 계좌 정보 (대변 거래처용) — 위하고 표준양식 일치 (계좌번호는 코드 098000으로 분리)
  const BANK_ACCOUNT_INFO = {
    SMT: { 거래처: '기업은행발안산단', 코드: '098000' },
    GW:  { 거래처: '기업은행',         코드: '098001' },  // GW 마스터 확정 시 갱신
  };
  // 분개 자동 생성 — 차변·대변 거래처 분리
  function generateEntry(tx, opts = {}){
    opts = Object.assign({clients:[], receivable:[]}, opts);
    const match = matchKeyword(tx);
    const isIn = tx.입금 > 0;
    const amt = isIn ? tx.입금 : tx.출금;

    // 신뢰도 결정
    let confidence = 'low'; // 🔴
    if (match) {
      const cnt = match.건수 || 0;
      if (match._ruleMatch) confidence = 'high';  // 카테고리 룰 = 명확 신뢰도 높음
      else if (cnt >= 5) confidence = 'high';     // 🟢
      else if (cnt >= 2) confidence = 'mid';      // 🟡
      else confidence = 'low';                     // 🔴
    }

    // 추천 계정 (Top 2)
    let recAccts = match ? (match.추천계정 || []).slice(0,3).map(([a,c]) => ({계정:a, 빈도:c})) : [];
    let recCli = match ? (match.추천거래처 || []).slice(0,2).map(([c,n]) => ({거래처:c, 빈도:n})) : [];
    let recMemo = match ? (match.추천적요 || []).slice(0,2).map(([m,n]) => ({적요:m, 빈도:n})) : [];

    // 사원 이름 처리:
    // 1) 학습된 적요에서 (이름) 패턴이 있으면 → 현재 거래자와 무관하므로 제거
    // 2) 메모·실제이름에 사원 이름이 명시된 경우만 추가
    recMemo = recMemo.map(m => ({
      ...m,
      적요: m.적요.replace(/\s*[(（][가-힣]{2,4}[)）]\s*/g, '').trim()
    }));
    const emp = getEmployeeFromTx(tx);
    if (emp) {
      // 메모/실제이름에 정말 그 사원이 등장하는 경우만 추가
      const fullText = (tx.메모 || '') + ' ' + (tx.실제이름 || '');
      if (fullText.includes(emp)) {
        if (recMemo.length > 0) recMemo[0].적요 += ' (' + emp + ')';
        else recMemo.push({적요: `${tx.메모||tx.실제이름} (${emp})`, 빈도: 1});
      }
    }

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

    // 자동 분개 — 차변/대변에 거래처 각각 박기 (위하고 분개 양식)
    const company = opts.company || 'SMT';
    const bankAcct = BANK_ACCOUNT_INFO[company] || BANK_ACCOUNT_INFO.SMT;
    // 실제 가맹점·송금자 거래처 (통장 계좌 제외)
    const isAccount = (c) => /기업은행|발안산단|04-012|230-067203|국민은행|하나은행|신한은행|카카오뱅크|\d{3}-\d{6}/.test(c);
    let realPartner = '';
    const real = (recCli || []).find(c => !isAccount(c.거래처));
    if (real) realPartner = real.거래처;
    else realPartner = tx.실제이름 || tx.메모 || (recCli[0] && recCli[0].거래처) || '';

    let chae, dae, memo, mainAcct;
    if (isIn) {
      // 입금 — 차변 보통예금(통장 거래처) / 대변 외상매출금(실제 거래처)
      chae = [{계정:'보통예금', 코드:'103', 금액:amt, 거래처: bankAcct.거래처, 거래처코드: bankAcct.코드}];
      let acct, code;
      if (receivableMatch || (recAccts[0] && recAccts[0].계정.includes('외상매출'))) {
        acct = '외상매출금'; code = '108';
      } else if (recAccts.length) {
        acct = recAccts[0].계정; code = '';
      } else {
        acct = '외상매출금'; code = '108';
      }
      dae = [{계정: acct, 코드: code, 금액: amt, 거래처: realPartner, 거래처코드: ''}];
      mainAcct = dae[0].계정;
    } else {
      // 출금 — 차변 비용(실제 가맹점) / 대변 보통예금(통장 거래처)
      const acct = recAccts.length ? recAccts[0].계정 : '미지급비용';
      chae = [{계정:acct, 코드:'', 금액:amt, 거래처: realPartner, 거래처코드: ''}];
      dae = [{계정:'보통예금', 코드:'103', 금액:amt, 거래처: bankAcct.거래처, 거래처코드: bankAcct.코드}];
      mainAcct = chae[0].계정;
    }
    memo = recMemo.length ? recMemo[0].적요 : (tx.메모 || tx.실제이름 || '');

    return {
      _tx: tx,
      신뢰도: confidence,
      차변: chae,
      대변: dae,
      적요: memo,
      // 거래처: 실제 가맹점·송금자 우선 (통장 계좌번호 제외)
      거래처: realPartner,  // 카드 표시용 (실제 가맹점)
      통장거래처: bankAcct.거래처,  // 통장 측 거래처
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

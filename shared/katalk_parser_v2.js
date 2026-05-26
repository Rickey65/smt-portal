/* SMT·건웅 통합 — 카톡 파서 v2 (2026-05-26)
   영업맨 카톡 자동 파싱 시스템.
   입력: 카톡 한 덩이 (또는 한 메시지 라인 묶음)
   출력: orders[] — 각 order = {영업맨, 거래처, 품목[], 배송유형, 발신·수신, 비고, _confidence, _raw}
*/
(function(global){
  'use strict';

  // ===== 1. 메시지 분할 =====
  // 카톡 시간 헤더: [이름] [오전/오후 H:MM]
  const MSG_HEADER = /^\[([^\]]+)\]\s*\[(오전|오후)\s*(\d{1,2}):(\d{2})\]\s*$/m;

  function splitMessages(text){
    const lines = text.split(/\r?\n/);
    const msgs = [];
    let cur = null;
    for (const line of lines) {
      const m = line.match(/^\[([^\]]+)\]\s*\[(오전|오후)\s*(\d{1,2}):(\d{2})\]\s*(.*)$/);
      if (m) {
        if (cur) msgs.push(cur);
        cur = {
          sender: m[1].trim(),
          ampm: m[2],
          hh: +m[3], mm: +m[4],
          body: m[5] || ''
        };
      } else if (cur) {
        cur.body += '\n' + line;
      }
    }
    if (cur) msgs.push(cur);
    return msgs.map(m => ({...m, body: m.body.trim()}));
  }

  // ===== 2. 모델·품목 인식 =====
  // 양사 발전기·엔진·양수기 prefix
  const MODEL_PATTERNS = [
    /\b(EU\d+[A-Za-z]*)\b/g,        // EU10i, EU30is, EU22i
    /\b(EM\d+[A-Z]*)\b/g,
    /\b(EG\d+[A-Z]*)\b/g,
    /\b(EK\d+[A-Z]*)\b/g,
    /\b(EZ\d+[A-Z]*)\b/g,
    /\b(ER\d+[A-Z]*)\b/g,
    /\b(E1[A-Z]\d*[A-Z]*)\b/g,
    /\b(SG\d+[A-Z]*[\-A-Z]*)\b/g,   // SG5000SX, SG7500SX, SG10000EX
    /\b(SP\d+[A-Z]*)\b/g,
    /\b(SH[A-Z0-9\-]+)\b/g,
    /\b(KG\d+[A-Z]*)\b/g,           // KG7500EX
    /\b(KS\d+[A-Z]*)\b/g,           // KS435
    /\b(XE\d+[A-Z]*)\b/g,
    /\b(EP\d+[A-Z]*)\b/g,
    /\b(GW\d+[A-Za-z]*)\b/g,        // TORO GW
    /\b(GPW\d+[A-Z]*)\b/g,
    /\b(GX\d+[A-Z]*[\-A-Z0-9]*)\b/g,
    /\b(GP\d+[A-Z]*)\b/g,
    /\b(GK\d+[A-Z]*[\-A-Z0-9]*)\b/g,
    /\b(WB\d+[A-Z]*)\b/g,           // WB30XT 양수기
    /\b(BS\d+[A-Z]*)\b/g,           // 보조연료탱크
    /\b(EGS\s*\d+[xX×]\d+)\b/gi,    // EGS 10x100
  ];

  function extractItems(text){
    const items = [];
    // 표준화: 공백·하이픈 정규화
    const t = text.replace(/[Ⅹx×]/g, 'x');
    const seen = new Set();
    for (const re of MODEL_PATTERNS) {
      let m;
      const re2 = new RegExp(re.source, re.flags);
      while ((m = re2.exec(t)) !== null) {
        const code = m[1].toUpperCase().replace(/\s+/g,'');
        if (seen.has(code)) continue;
        seen.add(code);
        // 모델 뒤 N대 / N개 추출
        const after = t.slice(m.index + m[0].length, m.index + m[0].length + 30);
        let qty = 1;
        const qm = after.match(/(\d+)\s*[대개ea]/i);
        if (qm) qty = +qm[1];
        // 단가 추출 — "99만원" / "99만" / "1,450,000"
        let price = 0;
        const pm1 = after.match(/(\d+(?:[\.,]\d+)?)\s*만\s*원?/);
        const pm2 = after.match(/([\d,]{4,})\s*원/);
        if (pm1) price = Math.round(parseFloat(pm1[1].replace(',','.')) * 10000);
        else if (pm2) price = parseInt(pm2[1].replace(/,/g,''));
        items.push({모델: code, 수량: qty, 단가: price, _raw: m[0]});
      }
    }
    return items;
  }

  // ===== 3. 거래처 추출 =====
  // 지역명 사전 (시도·시군구 일부)
  const REGIONS = ['서울','부산','대구','인천','광주','대전','울산','세종',
    '경기','강원','충북','충남','전북','전남','경북','경남','제주',
    '수원','성남','용인','안양','과천','광명','부천','안산','시흥','평택','오산','이천','여주','화성','파주','김포',
    '청주','충주','제천','천안','아산','공주','보령','전주','익산','군산','정읍','남원','목포','여수','순천','광양',
    '포항','경주','안동','김천','구미','상주','문경','창원','진주','통영','김해','양산','거제','밀양','원주','강릉','춘천'];
  const REGION_RE = new RegExp('(' + REGIONS.join('|') + ')', 'g');

  // 결제·배송 키워드
  const PAYMENT_KW = ['현불','선불','착불','외상','현금','대납','월말','익월말','결제','청구'];
  const SHIP_KW_NORMAL = ['정배','정기배송','오늘발송','내일발송','내일정배','대신화물','대신택배','경동화물','경동택배'];

  function extractClient(body, knownClients, aliases){
    // 1) 직직배송 양식 '보내는분:' — 발신측이 곧 청구 거래처
    const direct2 = body.match(/보내는분\s*:\s*([^\n]+)/);
    if (direct2) {
      const nm = direct2[1].trim();
      // CLIENTS와 매칭 시도
      if (knownClients) {
        const exact = knownClients.find(c => (c.거래처명||'') === nm);
        if (exact) return { name: exact.거래처명, source: 'direct2_exact' };
        const partial = knownClients.find(c => (c.거래처명||'').includes(nm) || nm.includes(c.거래처명||''));
        if (partial) return { name: partial.거래처명, source: 'direct2_partial', raw: nm };
      }
      return { name: nm, source: 'direct2_raw' };
    }

    // 2) 첫 줄에서 거래처 키워드 추출
    const firstLine = body.split('\n')[0].trim();
    // 모델·숫자·결제·배송 키워드 제거 → 후보 키워드 추출
    let cleaned = firstLine
      .replace(/^(내일|오늘|이번주|다음주|정배에?|정기배송|호남|영남|충청)/g, '').trim()
      .replace(/(EU|EM|EG|EK|EZ|ER|E1|SG|SP|SH|KG|KS|XE|EP|GW|GPW|GX|GP|GK|WB|BS|EGS)\s*\d+[A-Za-z0-9\-x×]*/gi,' ')
      .replace(/\d+\s*[대개]/g,' ')
      .replace(/[\d,]+\s*(원|만원|만)/g,' ');
    PAYMENT_KW.forEach(kw => { cleaned = cleaned.replace(new RegExp(kw,'g'),' '); });
    SHIP_KW_NORMAL.forEach(kw => { cleaned = cleaned.replace(new RegExp(kw,'g'),' '); });
    cleaned = cleaned.replace(/[\s,.]+/g,' ').trim();

    // 3) 별칭 사전 매칭 (가장 정확)
    if (aliases) {
      for (const alias of Object.keys(aliases)) {
        if (cleaned.includes(alias)) {
          const real = aliases[alias];
          if (knownClients) {
            const c = knownClients.find(c => (c.거래처명||'') === real);
            if (c) return { name: c.거래처명, source: 'alias', token: alias };
          }
          return { name: real, source: 'alias_raw', token: alias };
        }
      }
    }

    // 4) CLIENTS 사전 매칭 — 토큰별 부분일치, 가장 긴 매칭 우선
    if (knownClients && cleaned) {
      const tokens = cleaned.split(/\s+/).filter(t => t.length >= 2);
      let best = null, bestLen = 0;
      for (const c of knownClients) {
        const cn = (c.거래처명 || '').replace(/\s+/g,'');
        if (!cn || (c.폐업일자 && String(c.폐업일자).trim())) continue;
        for (const tok of tokens) {
          if (cn.includes(tok) && tok.length > bestLen) {
            best = c; bestLen = tok.length;
            break;
          }
        }
      }
      if (best) return { name: best.거래처명, source: 'client_partial', token: cleaned };
    }
    return cleaned ? { name: cleaned, source: 'guess', raw: firstLine } : null;
  }

  // ===== 4. 배송유형 인식 =====
  function detectShipType(body){
    // 직직배송 양식 — 보내는분·받는분 둘 다
    if (/보내는분\s*:/.test(body) && /받는분\s*:/.test(body)) return '직직배송';
    // 직배송 키워드 — "현불대납 청구" + 수신자 주소
    if (/대납\s*청구/.test(body) && /\d{5}.*[가-힣]/.test(body)) return '직배송요청';
    // 직배송 — 받는 사람 주소·전화 들어있고 거래처 외 
    if (/받는분\s*:/.test(body)) return '직배송요청';
    return '일반';
  }

  // ===== 5. 결제·운송비 인식 =====
  function detectShipFee(body){
    if (/착불/.test(body)) return '착불';
    if (/선불/.test(body)) return '선불';
    if (/현불/.test(body)) return '현불';
    return '';
  }

  function detectPayment(body){
    if (/현금|현결|현/.test(body) && !/현불/.test(body)) return '현금';
    if (/익월말/.test(body)) return '익월말';
    if (/월말/.test(body)) return '월말';
    if (/외상/.test(body)) return '외상';
    return '';
  }

  // ===== 6. 발신/수신 주소·연락처 =====
  const PHONE_RE = /(0\d{1,2}[\-\s]?\d{3,4}[\-\s]?\d{4})|(\d{10,11})/g;
  const ZIP_RE = /\(?(\d{5})\)?/;

  function extractContacts(body){
    const sender = {}, recv = {};
    // 표준 양식
    const m1 = body.match(/보내는분\s*:\s*([^\n]+)/); if (m1) sender.name = m1[1].trim();
    const m2 = body.match(/연락처\s*:\s*([^\n]+)/); if (m2) sender.phone = m2[1].trim();
    const m3 = body.match(/받는분\s*:\s*([^\n]+)/);
    if (m3) {
      const r = m3[1].trim();
      const ph = r.match(PHONE_RE);
      if (ph) { recv.phone = ph[0]; recv.name = r.replace(PHONE_RE,'').trim(); }
      else recv.name = r;
    }
    const m4 = body.match(/주소\s*:\s*\n?\s*([^\n]+(?:\n[^\n:]+)?)/);
    if (m4) recv.addr = m4[1].replace(/\s+/g,' ').trim();
    const m5 = body.match(/택배\s*메모\s*:\s*([^\n]+)/); if (m5) recv.memo = m5[1].trim();

    // 자유 양식 — 김성원 11:25 같은 케이스: 멀티라인 주소·전화·이름
    if (!recv.name && !recv.addr) {
      const lines = body.split('\n').map(l=>l.trim()).filter(Boolean);
      // 주소 라인 (지역명 + 도로명/번지)
      for (const ln of lines) {
        if (REGION_RE.test(ln) && /(\d+|로|길|동)/.test(ln) && ln.length > 8) {
          if (!recv.addr) recv.addr = ln.replace(PHONE_RE,'').trim();
        }
      }
      // 전화·이름 라인 (010-... 최명호)
      for (const ln of lines) {
        const ph = ln.match(PHONE_RE);
        if (ph) {
          if (!recv.phone) recv.phone = ph[0];
          const rest = ln.replace(PHONE_RE,'').trim();
          const nm = rest.match(/[가-힣]{2,4}\s*$/);
          if (nm && !recv.name) recv.name = nm[0].trim();
        }
      }
    }
    return { sender, recv };
  }

  // ===== 7. 컨텍스트 페어링 (같은 영업맨 5분 이내 합치기) =====
  function timeMinutes(m){
    let h = m.hh;
    if (m.ampm === '오후' && h < 12) h += 12;
    if (m.ampm === '오전' && h === 12) h = 0;
    return h*60 + m.mm;
  }

  function pairMessages(msgs){
    const grouped = [];
    for (const m of msgs) {
      const last = grouped[grouped.length-1];
      if (last && last.sender === m.sender && Math.abs(timeMinutes(m) - timeMinutes(last.lastTime)) <= 5) {
        last.body += '\n' + m.body;
        last.lastTime = m;
        last._merged = (last._merged || 1) + 1;
      } else {
        grouped.push({ sender: m.sender, body: m.body, time: m, lastTime: m });
      }
    }
    return grouped;
  }

  // ===== 8. 메인 파서 =====
  function parseAll(text, knownClients, aliases){
    const rawMsgs = splitMessages(text);
    const merged = pairMessages(rawMsgs);
    const orders = [];

    // 헤더 적용 (예: "내일 호남정기배송" → 후속 메시지에 배송일·정기배송)
    let pendingHeader = null;

    for (const grp of merged) {
      const body = grp.body.trim();
      if (!body) continue;
      // 사진·시스템 메시지 무시
      if (/^사진$|^이모티콘|^동영상/.test(body)) continue;

      // 헤더성 메시지 (모델·수량 없음) → pendingHeader
      const itemsCheck = extractItems(body);
      if (itemsCheck.length === 0 && body.length < 30 && /정배|정기배송|배송/.test(body)) {
        pendingHeader = body;
        continue;
      }
      // 헤더 단독 전화번호 — 다음 거래에 페어링
      if (itemsCheck.length === 0 && /^[\d\s\-]+$/.test(body)) {
        pendingHeader = (pendingHeader || '') + ' ' + body;
        continue;
      }

      // 실제 주문 메시지
      const shipType = detectShipType(body);
      const items = extractItems(body);
      const client = extractClient(body, knownClients, aliases);
      const contacts = extractContacts(body);

      // 모델 없으면 신뢰도 낮음 (수정·취소 메시지일 수도)
      const confidence = items.length > 0 ? 'high' : 'low';

      // 수정·취소 메시지 감지
      const isModify = /다음주로|미루|취소|연기|변경|추가|빼고/.test(body);

      orders.push({
        영업맨: grp.sender,
        시간: `${grp.time.ampm} ${grp.time.hh}:${String(grp.time.mm).padStart(2,'0')}`,
        거래처: client ? client.name : null,
        _거래처매칭: client ? client.source : null,
        품목: items,
        배송유형: shipType,
        화물운송비: detectShipFee(body),
        결제조건: detectPayment(body),
        발신: contacts.sender,
        수신: contacts.recv,
        헤더: pendingHeader,
        비고: body.length > 200 ? body.slice(0,200)+'...' : body,
        _confidence: confidence,
        _수정여부: isModify,
        _raw: body
      });
      pendingHeader = null;  // 사용 후 클리어
    }
    return orders;
  }

  global.KatalkParserV2 = { splitMessages, extractItems, extractClient, detectShipType, parseAll };
})(window);

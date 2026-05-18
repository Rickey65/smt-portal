/* 카톡 텍스트 자동 파싱 모듈 v1
   영업맨 카톡 → 거래처/품목/수량/가격/배송 자동 추출
   사용: const result = KatalkParser.parse(rawText, { clients, inventory });
*/
(function(global) {
  'use strict';

  // ===== 품목 별칭 사전 =====
  const ITEM_ALIASES = {
    '혼다발전기|혼다': ['EU','EM','EG','EP','EZ','HONDA','GENERATOR'],
    '재커리|jackery': ['JACKERY','EXPLORER','솔라','파워뱅크'],
    '블루에티|bluetti': ['BLUETTI','AC','EB'],
    '예초기': ['BCM','EBCM','SG','LS','TBC','예초'],
    '가스예초기': ['SG','SGS','SGE','가스예초'],
    '펌프|양수기|살수기': ['WP','WT','WB','펌프','양수','살수'],
    '송풍기|블로워': ['BLX','HHB','EBZ','송풍','블로워'],
    '엔진': ['GX','GP','GC','EX','엔진'],
    '바이브레이터|진동기': ['EH','진동','바이브'],
    '콤팩터|평판다짐기': ['EP','HONDAJP','평판','콤팩터'],
    '용접기': ['MMA','MIG','용접'],
    '인버터발전기': ['EU','이그제큐티브','인버터'],
    '온열매트': ['EM','온열'],
    '폴딩카트': ['HC','폴딩'],
    '커팅기': ['HHE','커팅'],
    '리코일|스타팅': ['리코일','스타팅','스타터'],
  };

  // ===== 배송 키워드 =====
  const SHIP_TYPE_KW = {
    직직배송: ['직송건', '직배송건', '직직배송'],
    직배송요청: ['직배송요청', '직배송'],
    일반: []
  };
  
  const SHIP_METHOD_KW = {
    대신화물: ['대신화물', '대신택배', '대신'],
    일반택배: ['택배', '우체국', 'CJ', '한진', '로젠'],
    고객수령: ['고객수령', '직접수령', '인수'],
    정규배송: ['정규배송', '정배']
  };

  const TODAY_KW = {
    오늘: ['오늘', '오늘발송', '오늘배송', '금일'],
    내일: ['내일', '내일배송', '내일발송'],
    모레: ['모레', '모레배송'],
    이번주금: ['금요', '금요일', '이번주 금', '이번주금'],
    월: ['월요', '월요일'],
    화: ['화요', '화요일'],
    수: ['수요', '수요일'],
    목: ['목요', '목요일'],
  };

  const PAYMENT_KW = {
    선불: ['선불', '현불', '결제완료', '입금'],
    착불: ['착불', '도착불']
  };

  // ===== 헬퍼 =====
  function normalize(text) {
    return text.replace(/\s+/g, ' ').trim();
  }
  
  function detectInList(text, dict) {
    const lower = text.toLowerCase();
    for (const [key, kws] of Object.entries(dict)) {
      for (const kw of kws) {
        if (lower.includes(kw.toLowerCase())) return key;
      }
    }
    return null;
  }

  // ===== 1. 거래처 매칭 =====
  function matchClient(text, allClients) {
    if (!allClients || allClients.length === 0) return null;
    
    // 라인별로 후보 추출
    const lines = text.split('\n').filter(l => l.trim());
    const candidates = [];
    
    // 거래처 마스터 정규화 (검색용)
    const normalized = allClients.map(c => ({
      ...c,
      _searchName: (c.거래처명 || c.상호 || '').toLowerCase().replace(/[()㈜주식회사 ]/g, '')
    }));
    
    for (const line of lines.slice(0, 6)) {  // 상위 6줄에서 찾기
      const cleanLine = line.toLowerCase().replace(/[()㈜주식회사 ]/g, '');
      for (const c of normalized) {
        if (!c._searchName || c._searchName.length < 3) continue;
        if (cleanLine.includes(c._searchName)) {
          candidates.push({ client: c, score: c._searchName.length, line: line });
        }
      }
    }
    
    // 가장 긴(가장 구체적) 매칭 선택
    candidates.sort((a, b) => b.score - a.score);
    return candidates[0] ? candidates[0].client : null;
  }

  // ===== 2. 품목 매칭 =====
  function expandAliases(query) {
    const out = new Set([query.toLowerCase()]);
    for (const [pat, kws] of Object.entries(ITEM_ALIASES)) {
      if (new RegExp(pat, 'i').test(query)) {
        kws.forEach(k => out.add(k.toLowerCase()));
      }
    }
    return Array.from(out);
  }
  
  function matchItem(text, allInventory) {
    if (!allInventory || allInventory.length === 0) return [];
    const text_lower = text.toLowerCase();
    const results = [];
    
    for (const inv of allInventory) {
      const code = (inv.품목코드 || '').toLowerCase();
      const name = (inv.품명 || inv.품목명 || '').toLowerCase();
      const spec = (inv.규격 || '').toLowerCase();
      
      // 품목코드 직접 매칭 (가장 정확)
      if (code && code.length >= 4 && text_lower.includes(code)) {
        results.push({ item: inv, score: code.length + 10, type: 'code' });
        continue;
      }
      // 품목명 매칭
      if (name && name.length >= 4 && text_lower.includes(name)) {
        results.push({ item: inv, score: name.length, type: 'name' });
        continue;
      }
    }
    
    results.sort((a, b) => b.score - a.score);
    // 중복 제거 (같은 코드 하나만)
    const seen = new Set();
    const unique = [];
    for (const r of results) {
      const c = r.item.품목코드;
      if (!seen.has(c)) { seen.add(c); unique.push(r); }
    }
    return unique.slice(0, 10);
  }

  // ===== 3. 수량·단가 추출 =====
  // 패턴: "30a 8.5×50 2대", "MCC13×70 1대", "SG7600 SXE 1대", "30a 밸브 50개"
  // "1,400,000원", "149만원", "51만 별도", "29만 별도"
  function extractItemsFromLines(lines, allInventory) {
    const itemMatches = [];
    
    for (const line of lines) {
      // 수량+단위 패턴
      const qtyMatch = line.match(/(\d+)\s*(대|개|박스|EA|ea|짝|세트|롤)/);
      if (!qtyMatch) continue;
      
      const qty = parseInt(qtyMatch[1]);
      const lineBeforeQty = line.substring(0, qtyMatch.index).trim();
      
      // 같은 라인 또는 인접한 라인에서 단가 찾기
      let price = 0;
      const priceMatch = line.match(/(\d{1,3}(,\d{3})*|\d+)\s*만/);  // "149만", "51만"
      if (priceMatch) {
        price = parseInt(priceMatch[1].replace(/,/g,'')) * 10000;
      } else {
        const priceMatch2 = line.match(/(\d{1,3}(,\d{3})+)\s*원?/);  // "1,400,000원"
        if (priceMatch2) {
          price = parseInt(priceMatch2[1].replace(/,/g,''));
        }
      }
      
      // 품목 매칭 (라인 텍스트 전체로)
      const matches = matchItem(line, allInventory);
      const matched = matches[0] ? matches[0].item : null;
      
      itemMatches.push({
        rawText: line,
        품목코드: matched ? matched.품목코드 : '',
        품목명: matched ? (matched.품명 || matched.품목명) : lineBeforeQty,
        규격: matched ? matched.규격 : '',
        단위: matched ? (matched.단위 || '') : '',
        수량: qty,
        단가: price,
        매칭상태: matched ? 'matched' : 'unmatched',
        매칭후보: matches.slice(0, 3).map(m => ({ code: m.item.품목코드, name: m.item.품명 || m.item.품목명 }))
      });
    }
    
    return itemMatches;
  }

  // ===== 4. 배송 정보 추출 =====
  function extractShipInfo(text) {
    return {
      배송유형: detectInList(text, SHIP_TYPE_KW) || '일반',
      배송방법: detectInList(text, SHIP_METHOD_KW) || '대신화물',
      배송일: detectInList(text, TODAY_KW) || '오늘',
      결제: detectInList(text, PAYMENT_KW) || (text.match(/별도/) ? '착불' : '선불')
    };
  }

  // ===== 5. 연락처·주소·이름 추출 =====
  function extractContacts(text) {
    const phones = [];
    const phoneRegex = /(\d{2,4}[-\s.]\d{3,4}[-\s.]\d{4})/g;
    let m;
    while ((m = phoneRegex.exec(text)) !== null) {
      phones.push(m[1].replace(/\s/g, '-').replace(/\./g, '-'));
    }
    
    // 주소 추출 (서울/경기/...로 시작하는 라인)
    const addressMatch = text.match(/((서울|경기|부산|대전|대구|광주|인천|울산|강원|충북|충남|전북|전남|경북|경남|제주)[^\n]*)/);
    const address = addressMatch ? addressMatch[1].trim() : '';
    
    // 사람 이름 추출 (한글 2~4글자 + 슬래시·줄바꿈)
    const nameMatch = text.match(/([가-힣]{2,4})\s*[\/]/);
    const name = nameMatch ? nameMatch[1] : '';
    
    return { phones, address, name };
  }

  // ===== 6. 메인 파서 =====
  global.KatalkParser = {
    
    parse(rawText, options = {}) {
      const { clients = [], inventory = [], salesperson = '' } = options;
      
      const text = normalize(rawText);
      const lines = rawText.split('\n').map(l => l.trim()).filter(l => l);
      
      const result = {
        영업맨: salesperson,
        rawText: rawText,
        파싱시간: new Date().toISOString(),
        경고: []
      };
      
      // 1. 거래처 매칭
      const clientMatch = matchClient(text, clients);
      result.거래처 = clientMatch ? {
        코드: clientMatch.거래처코드 || clientMatch.코드 || clientMatch.code || '',
        명: clientMatch.거래처명 || clientMatch.상호 || '',
        대표: clientMatch.대표자명 || clientMatch.대표자 || '',
        주소: clientMatch.사업장주소 || clientMatch.주소 || '',
        매칭상태: 'matched'
      } : {
        코드: '', 명: lines[0] || '', 매칭상태: 'unmatched'
      };
      if (!clientMatch) result.경고.push('거래처 매칭 실패 — 마스터에서 확인');
      
      // 2. 품목 추출
      result.품목들 = extractItemsFromLines(lines, inventory);
      const unmatched = result.품목들.filter(i => i.매칭상태 === 'unmatched');
      if (unmatched.length > 0) {
        result.경고.push(`품목 ${unmatched.length}개 매칭 실패 — 수동 선택 필요`);
      }
      if (result.품목들.length === 0) {
        result.경고.push('품목을 인식하지 못함 — 카톡에 "1대", "2개" 같은 수량 표기 필요');
      }
      
      // 3. 배송 정보
      result.배송 = extractShipInfo(text);
      
      // 4. 연락처·주소
      result.연락 = extractContacts(text);
      
      // 5. 직송 유형 자동 판정
      const hasDirectSend = /직송건|직배송/i.test(text);
      const hasOtherAddress = result.연락.address && clientMatch && 
                              clientMatch.주소 && !text.includes(clientMatch.주소.slice(0, 8));
      
      if (hasDirectSend && hasOtherAddress) {
        result.배송.배송유형 = '직직배송';
      } else if (hasDirectSend || (hasOtherAddress && result.연락.phones.length > 0)) {
        result.배송.배송유형 = '직배송요청';
      }
      
      // 6. 비고 추출 (남은 텍스트 중 의미 있는 것)
      const memos = [];
      if (text.match(/대납청구/)) memos.push('대납청구');
      if (text.match(/계산서.*처리|계산서\s*\S+/)) memos.push(text.match(/계산서[^,\n]+/)[0]);
      if (text.match(/현불|선결제/)) memos.push('현불');
      if (text.match(/기존가/)) memos.push('기존가 적용');
      if (text.match(/별도/)) memos.push('VAT 별도');
      result.비고 = memos.join(', ');
      
      return result;
    }
  };

})(window);

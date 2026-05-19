/* SMT·건웅 통합 시스템 — 공통 데이터 로더 v2 */
(function(global) {
  'use strict';

  // ===== 회사 토글 상태 =====
  const TOGGLE_KEY = 'smt_company_filter';
  
  global.Company = {
    get() { return sessionStorage.getItem(TOGGLE_KEY) || 'ALL'; },
    set(v) { 
      sessionStorage.setItem(TOGGLE_KEY, v);
      window.dispatchEvent(new CustomEvent('company-changed', { detail: v }));
    },
    label() {
      const v = this.get();
      return v === 'SMT' ? 'SMT서울기연' : v === 'GW' ? '주식회사건웅' : '양사 통합';
    },
    color() {
      const v = this.get();
      return v === 'SMT' ? '#3B7EE5' : v === 'GW' ? '#8845D5' : '#1D9E75';
    }
  };

  // ===== 데이터 캐시 (한 페이지 세션 내) =====
  const CACHE = {};
  
  async function fetchJson(path) {
    if (CACHE[path]) return CACHE[path];
    try {
      const res = await fetch(path);
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const data = await res.json();
      CACHE[path] = data;
      return data;
    } catch (e) {
      console.warn(`[DataLoader] ${path} 로딩 실패:`, e.message);
      return null;
    }
  }

  // ===== 양사 데이터 통합 헬퍼 =====
  global.Data = {
    
    // 판매: 회사 토글에 따라 양사 또는 한쪽 합치기
    async sales(years = [2024, 2025, 2026]) {
      const company = Company.get();
      const all = [];
      const companies = company === 'ALL' ? ['smt', 'gw'] : [company.toLowerCase()];
      
      for (const c of companies) {
        for (const y of years) {
          const d = await fetchJson(`data/per_company/sales_${c}_${y}.json`);
          if (d && d.rows) all.push(...d.rows);
        }
      }
      return all;
    },
    
    async inventory() {
      const company = Company.get();
      const all = [];
      const companies = company === 'ALL' ? ['smt', 'gw'] : [company.toLowerCase()];
      for (const c of companies) {
        const d = await fetchJson(`data/per_company/inventory_${c}.json`);
        if (d && d.rows) all.push(...d.rows);
      }
      return all;
    },
    
    async receivable() {
      const company = Company.get();
      const all = [];
      const companies = company === 'ALL' ? ['smt', 'gw'] : [company.toLowerCase()];
      for (const c of companies) {
        const d = await fetchJson(`data/per_company/receivable_${c}.json`);
        if (d && d.clients) all.push(...d.clients);
      }
      return all;
    },
    
    async clients() {
      // 거래처는 통합본 우선
      const merged = await fetchJson('data/shared/clients.json');
      if (!merged) return [];
      const company = Company.get();
      if (company === 'ALL') return merged.rows;
      return merged.rows.filter(c => c._companies && c._companies.includes(company));
    },
    
    async itemOwner() {
      const d = await fetchJson('data/shared/item_owner.json');
      return d ? d.items : {};
    },
    
    async employees() {
      const d = await fetchJson('data/shared/employees.json');
      if (!d || !d.rows) return [];
      const company = Company.get();
      if (company === 'ALL') return d.rows;
      return d.rows.filter(e => e._company === company);
    },
    
    async shippingFee() {
      const d = await fetchJson('data/shared/shipping_fee.json');
      return d || { default: {}, items: {} };
    },
    
    async clientAliases() {
      const d = await fetchJson('data/shared/client_aliases.json');
      return (d && d.aliases) ? d.aliases : {};
    },
    
    // 회사 정보 (거래명세표 등)
    companyInfo: {
      SMT: {
        등록번호: '134-81-77776',
        상호: '에스엠티서울기연주식회사',
        대표: '임창수',
        주소: '경기 화성시 만세구 향남읍 발안공단로1길 81 2층',
        업태: '도매및소매업',
        종목: '전기용기계ㆍ장비및관련기자재도매업'
      },
      GW: {
        등록번호: '222-86-01511',
        상호: '주식회사건웅',
        대표: '김태종',
        주소: '서울 금천구 가산디지털1로 33-33 804호 (가산동, 대륭테크노타운2차)',
        업태: '도매및소매업',
        종목: '전자상거래소매업'
      }
    },
    
    // 증빙번호 생성 (STX + YYYYMMDD + 90001~99999)
    // 위하고 자체 일련번호(보통 00001~50000)와 겹치지 않도록 90000번대 사용
    generateChunkNo(prefix = 'STX') {
      const today = new Date();
      const yyyymmdd = today.getFullYear() + 
        String(today.getMonth() + 1).padStart(2, '0') +
        String(today.getDate()).padStart(2, '0');
      const seqKey = `${prefix}_${yyyymmdd}_seq`;
      let seq = parseInt(sessionStorage.getItem(seqKey) || '90000');
      seq++;
      if (seq > 99999) seq = 90001;
      sessionStorage.setItem(seqKey, seq);
      return `${prefix}${yyyymmdd}${String(seq).padStart(5, '0')}`;
    }
  };
})(window);

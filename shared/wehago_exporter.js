/* 위하고 SmartA10 일괄 업로드 엑셀 4종 생성기 v1 
   입고전표·출고전표·일반전표·세금계산서 (SheetJS 기반)
   사용처: module_order, module_acct
*/
(function(global) {
  'use strict';
  
  // SheetJS는 외부 CDN 또는 호출 페이지에서 로딩
  // <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
  
  if (!global.XLSX) {
    console.warn('[WehagoExporter] XLSX(SheetJS) 라이브러리 미로딩. CDN 추가 필요.');
  }

  // ===== 계정과목 코드 매핑 =====
  const ACCT = {
    보통예금: { code: '103', 구분차: 3, 구분대: 4 },
    외상매출금: { code: '108', 구분차: 3, 구분대: 4 },
    상품매출: { code: '401', 구분차: 3, 구분대: 4 },
    부가세예수금: { code: '255', 구분차: 3, 구분대: 4 },
    매입: { code: '146', 구분차: 3, 구분대: 4 },
    외상매입금: { code: '251', 구분차: 3, 구분대: 4 },
    부가세대급금: { code: '135', 구분차: 3, 구분대: 4 },
    복리후생비: { code: '811', 구분차: 1, 구분대: 2 },
    여비교통비: { code: '812', 구분차: 1, 구분대: 2 },
    차량유지비: { code: '822', 구분차: 1, 구분대: 2 },
    지급수수료: { code: '831', 구분차: 1, 구분대: 2 }
  };

  // ===== 회사 정보 =====
  const COMPANIES = {
    SMT: { 명: '에스엠티서울기연주식회사', 등록번호: '134-81-77776', 대표: '임창수' },
    GW:  { 명: '주식회사건웅',           등록번호: '222-86-01511', 대표: '김태종' }
  };

  // 헬퍼: 일자 분리
  function splitDate(d) {
    if (!d) d = new Date().toISOString().slice(0,10);
    const [y, m, day] = d.split('-');
    return { y: parseInt(y), m: parseInt(m), d: parseInt(day) };
  }


  // ===== 위하고 표준 2시트 구조 헬퍼 =====
  // 사이트가 받는 위하고 양식:
  //  시트1 = "출고처리 거래처정보(상단)" — 그룹번호·일자·거래처코드·부서코드·사원코드·관리항목코드·VAT여부값·과세구분·비고1·2·3
  //  시트2 = "출고처리 폼목정보(하단)" — 그룹번호·품목코드·규격·납기일자·수량·단가·공급가액·부가세액·창고코드·프로젝트코드·품목비고·입고단가
  //  그룹번호로 상단·하단 연결 (1주문 = 1그룹, 품목 N건이면 하단 N행)


  // 거래처명 → 거래처코드 폴백 (영업맨 던지기에서 코드 누락 시)
  let _CLIENTS_CACHE = null;
  async function _loadClients(){
    if (_CLIENTS_CACHE) return _CLIENTS_CACHE;
    try {
      const r = await fetch('data/shared/clients.json?t='+Date.now());
      if (r.ok) {
        const d = await r.json();
        _CLIENTS_CACHE = d.rows || [];
      }
    } catch(e){ _CLIENTS_CACHE = []; }
    return _CLIENTS_CACHE || [];
  }
  function _findClientCode(name){
    if (!name) return '';
    const list = window._wehagoClients || _CLIENTS_CACHE || [];
    const exact = list.find(c => (c.거래처명 || c.상호) === name);
    if (exact) return exact.거래처코드 || exact.코드 || exact.code || '';
    const partial = list.find(c => ((c.거래처명 || c.상호 || '').includes(name)) || (name.includes(c.거래처명 || '')));
    return partial ? (partial.거래처코드 || partial.코드 || partial.code || '') : '';
  }
  // 페이지 로드 시 미리 캐시
  _loadClients().then(list => { window._wehagoClients = list; });

  // 양사 회사별 디폴트 코드 (필요시 부장님 위하고 코드로 보정)
  const WEHAGO_DEFAULTS = {
    SMT: { 부서코드:'0002', 사원코드:'122', 창고코드:'12', 부서명:'영업팀' },  // 부서10·창고화성하길리·사원임성우
    GW:  { 부서코드:'0001', 사원코드:'07',  창고코드:'1',  부서명:'영업팀' },
  };

  // 결제조건 → 관리항목코드 매핑 (부장님 확인 필요 — 디폴트)
  const PAY_TO_GWANRI = { '현금':'1', '월말':'2', '익월말':'3', '카드':'4', '기타':'9' };

  // 일자 yyyymmdd
  function ymd(d){
    if (!d) d = new Date().toISOString().slice(0,10);
    return d.replace(/-/g,'');
  }

  // ===== WehagoExporter 객체 초기화 =====
  global.WehagoExporter = global.WehagoExporter || {};

  // ===== 1. 출고전표 — 2시트 구조 =====
  global.WehagoExporter.makeChulgo = function(orders, company){
    if (!global.XLSX) return null;
    const def = WEHAGO_DEFAULTS[company] || WEHAGO_DEFAULTS.SMT;
    // 상단 헤더
    const topHeader = ['그룹번호','일자','거래처코드','부서코드','사원코드','관리항목코드','VAT여부값','과세구분','비고1','비고2','비고3'];
    const botHeader = ['그룹번호','품목코드','규격','납기일자','수량','단가','공급가액','부가세액','창고코드','프로젝트코드','품목비고','입고단가'];
    const topRows = [topHeader];
    const botRows = [botHeader];
    let groupNo = 0;
    orders.filter(o => (o.발행회사||'SMT') === company).forEach(o => {
      groupNo++;
      const gwanri = PAY_TO_GWANRI[o.결제조건] || '2';
      const op = o.업무맨_사원코드 || def.사원코드;
      const cliCode = o.거래처코드 || _findClientCode(o.거래처명) || '';
      // 상단 1행
      topRows.push([
        groupNo,                       // 그룹번호
        ymd(o.일자),                   // 일자 yyyymmdd
        cliCode,                       // 거래처코드
        def.부서코드,                  // 부서코드
        op,                            // 사원코드 (업무맨)
        gwanri,                        // 관리항목코드 (결제조건 매핑)
        1,                             // VAT여부값 (1=별도, 디폴트)
        0,                             // 과세구분 (0=과세)
        o.배송방법 || '',              // 비고1 (택배사)
        (o.화물운송비 === '선불' ? '선불' : (o.화물운송비 === '착불' ? '착불' : '')),  // 비고2
        o.비고 || ''                   // 비고3
      ]);
      // 하단 N행 (품목별)
      (o.품목 || []).forEach(item => {
        const qty = +item.수량 || 0;
        const price = +item.단가 || 0;
        const supply = qty * price;
        const vat = Math.round(supply * 0.1);
        botRows.push([
          groupNo,                     // 그룹번호 (상단과 연결)
          item.품목코드 || '',          // 품목코드
          item.규격 || '',              // 규격
          ymd(o.일자),                  // 납기일자
          qty,                          // 수량
          price,                        // 단가
          supply,                       // 공급가액
          vat,                          // 부가세액
          def.창고코드,                 // 창고코드
          '',                           // 프로젝트코드
          item.품목비고 || '',          // 품목비고
          ''                            // 입고단가
        ]);
      });
    });
    const ws1 = XLSX.utils.aoa_to_sheet(topRows);
    const ws2 = XLSX.utils.aoa_to_sheet(botRows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws1, '출고처리 거래처정보(상단)');
    XLSX.utils.book_append_sheet(wb, ws2, '출고처리 폼목정보(하단)');
    return wb;
  };

  // ===== 2. 입고전표 — 2시트 구조 (양사 내부거래 받는 쪽) =====
  global.WehagoExporter.makeIpgo = function(orders, company){
    if (!global.XLSX) return null;
    const def = WEHAGO_DEFAULTS[company] || WEHAGO_DEFAULTS.SMT;
    const topHeader = ['그룹번호','일자','거래처코드','부서코드','사원코드','관리항목코드','VAT여부값','과세구분','비고1','비고2','비고3'];
    const botHeader = ['그룹번호','품목코드','납기일자','수량','단가','공급가액','부가세액','창고코드','프로젝트코드','품목비고'];
    const topRows = [topHeader];
    const botRows = [botHeader];
    let groupNo = 0;
    // 양사 내부거래에서 'company'가 입고 받는 쪽인 경우만
    orders.forEach(o => {
      const intern = o._internalTransfer;
      if (!intern || !intern.confirmed) return;
      // 발행회사 != company (입고는 받는 쪽)
      if ((o.발행회사 || 'SMT') === company) return;
      // 입고 받는 쪽 = company. 출고 발행 = 반대편
      const other = (o.발행회사 || 'SMT');  // 출고 발행
      groupNo++;
      const op = def.사원코드;  // 받는 쪽 사원코드
      topRows.push([
        groupNo,
        ymd(o.일자),
        _findClientCode(o.거래처명) || '',  // 거래처코드 (clients 사전 매칭)
        def.부서코드, op, '2', 1, 0,
        '양사내부 입고', '', o.비고 || ''
      ]);
      (o.품목 || []).forEach(item => {
        const qty = +item.수량 || 0;
        const price = +item._internalPrice || +item.단가 || 0;
        const supply = qty * price;
        const vat = Math.round(supply * 0.1);
        botRows.push([
          groupNo,
          item.품목코드 || '',
          ymd(o.일자),
          qty, price, supply, vat,
          def.창고코드, '', item.품목비고 || ''
        ]);
      });
    });
    const ws1 = XLSX.utils.aoa_to_sheet(topRows);
    const ws2 = XLSX.utils.aoa_to_sheet(botRows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws1, '입고처리 거래처정보(상단)');
    XLSX.utils.book_append_sheet(wb, ws2, '입고처리 폼목정보(하단)');
    return wb;
  };

  // ===== 3. 일반전표 엑셀 생성 (우리 표준양식) =====
  // 컬럼: 월·일·구분·계정코드·계정과목·거래처코드·거래처·적요·차변·대변
  global.WehagoExporter.makeIlban = function(orders, company) {
    if (!global.XLSX) return null;
    const rows = [];
    rows.push(['월','일','구분','Code','계정과목','Code','거래처','적요','차변(출금)','대변(입금)']);
    
    let entryNo = 0;
    orders.filter(o => o.발행회사 === company).forEach(o => {
      entryNo++;
      const { m, d } = splitDate(o.일자);
      const totalSupply = o.품목.reduce((a,i) => a + (i.수량 * i.단가), 0);
      const totalVat = Math.round(totalSupply * 0.1);
      const totalAmt = totalSupply + totalVat;
      const memo = `${o.거래처명} ${o.품목[0]?.품목명 || ''}${o.품목.length > 1 ? ` 외 ${o.품목.length-1}건` : ''}`;
      
      // 대체전표 패턴: 차변 외상매출금 + 대변 상품매출 + 대변 부가세예수금
      rows.push([m, d, 3, '108', '외상매출금', o.거래처코드, o.거래처명, memo, totalAmt, '']);
      rows.push([m, d, 4, '401', '상품매출',   o.거래처코드, o.거래처명, memo, '', totalSupply]);
      rows.push([m, d, 4, '255', '부가세예수금', o.거래처코드, o.거래처명, memo, '', totalVat]);
    });
    
    const ws = XLSX.utils.aoa_to_sheet(rows);
    
    // 컬럼 너비 설정
    ws['!cols'] = [
      {wch:5}, {wch:5}, {wch:6}, {wch:8}, {wch:15}, {wch:9}, {wch:25}, {wch:35}, {wch:14}, {wch:14}
    ];
    
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '일반전표');
    return wb;
  };

  // ===== 4. 세금계산서 엑셀 생성 =====
  // 위하고 세금계산서 일괄 등록 양식
  global.WehagoExporter.makeSegyumetan = function(orders, company) {
    if (!global.XLSX) return null;
    const rows = [];
    rows.push([
      '작성일자','거래처코드','거래처명','거래처사업자번호','대표자','업태','종목',
      '품목명','규격','수량','단가','공급가액','부가세','합계금액','비고','전자세금계산서구분'
    ]);
    
    orders.filter(o => o.발행회사 === company && o._taxIssue).forEach(o => {
      o.품목.forEach(item => {
        const supply = (item.수량 || 0) * (item.단가 || 0);
        const vat = Math.round(supply * 0.1);
        rows.push([
          o.일자,
          o.거래처코드,
          o.거래처명,
          o.거래처사업자번호 || '',
          o.거래처대표 || '',
          o.거래처업태 || '도매및소매업',
          o.거래처종목 || '',
          item.품목명,
          item.규격 || '',
          item.수량,
          item.단가,
          supply,
          vat,
          supply + vat,
          o.비고 || '',
          '01'  // 일반 발행
        ]);
      });
    });
    
    const ws = XLSX.utils.aoa_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '세금계산서');
    return wb;
  };

  // ===== 일괄 다운로드 (한 번에 회사 토글 기준 4종) =====
  global.WehagoExporter.downloadAll = function(orders, company) {
    if (!global.XLSX) {
      alert('XLSX 라이브러리가 로딩되지 않았습니다.');
      return;
    }
    
    const ts = new Date().toISOString().slice(0,16).replace(/[-:T]/g, '').slice(0,12);
    const companyName = company === 'SMT' ? 'SMT' : 'GW';
    
    // 각 회사 4종씩 다운로드
    const types = [
      { fn: 'makeIpgo',       name: '입고전표' },
      { fn: 'makeChulgo',     name: '출고전표' },
      { fn: 'makeIlban',      name: '일반전표' },
      { fn: 'makeSegyumetan', name: '세금계산서' }
    ];
    
    let downloaded = 0;
    types.forEach(t => {
      const wb = WehagoExporter[t.fn](orders, company);
      if (wb && wb.SheetNames.length > 0) {
        const ws = wb.Sheets[wb.SheetNames[0]];
        const range = XLSX.utils.decode_range(ws['!ref']);
        // 헤더만 있는 경우 (데이터 0행) 다운로드 안 함
        if (range.e.r > 0) {
          XLSX.writeFile(wb, `위하고_${t.name}_${companyName}_${ts}.xlsx`);
          downloaded++;
        }
      }
    });
    
    return downloaded;
  };

  // ===== 양사 모두 (통합 다운로드) =====
  global.WehagoExporter.downloadBoth = function(orders) {
    const smt = this.downloadAll(orders, 'SMT');
    const gw = this.downloadAll(orders, 'GW');
    return { SMT: smt, GW: gw };
  };

})(window);

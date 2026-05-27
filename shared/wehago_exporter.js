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
    SMT: { 명: '에스엠티서울기연주식회사', 등록번호: '13481777760'.replace(/-/g,''), 대표: '임창수',
           등록번호_원: '134-81-77776',
           주소: '경기 화성시 만세구 향남읍 발안공단로1길 81', 주소상세: '2층',
           업태: '도매및소매업', 종목: '전기용기계ㆍ장비및관련기자재도매업',
           담당자명: '임성우', 담당자전화: '031-1644-2090', 담당자이메일: 'smtseoul@daum.net' },
    GW:  { 명: '주식회사건웅', 등록번호: '22286015110'.replace(/-/g,''), 대표: '김태종',
           등록번호_원: '222-86-01511',
           주소: '서울 금천구 가산디지털1로 33-33', 주소상세: '804호 (가산동, 대륭테크노타운2차)',
           업태: '도매및소매업', 종목: '전자상거래소매업',
           담당자명: '전재홍', 담당자전화: '02-2113-2090', 담당자이메일: '' }
  };
  // 84컬럼 위하고 표준 세금계산서 헤더 (Row1·Row2 모두 데이터처럼 안 만들고 데이터만 작성)
  const SEGYU_HEADER_1 = Array.from({length:84}, (_,i)=>i+1);  // 1~84 숫자
  const SEGYU_HEADER_2 = ['발송시스템_고유번호','결과_고유번호','청구유형코드','거래구분코드','사용자구분코드','과세구분코드','책번호_권','책번호_호','일련번호','공급자정보','','','','','','','','','','','','','공급받는자정보','','','','','','','','','','','','','작성일자','공급가액','세액','품목1','','','','','','','','품목2','','','','','','','','품목3','','','','','','','','품목4','','','','','','','','합계금액','현금','수표','어음','외상미수금','비고1','비고2','비고3','수정세금계산서여부','수정코드','원천세금계산서 고유번호','원천세금계산서 승인번호','협력업체코드','종료'];

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
    SMT: { 부서코드:'0002', 사원코드:'122', 창고코드:'12', 부서명:'영업팀',
           // 양사간 거래처코드 (SMT 위하고에 박힌 GW 거래처)
           양사거래처코드:'005366', 양사거래처명:'주식회사 건웅' },
    GW:  { 부서코드:'0001', 사원코드:'07',  창고코드:'1',  부서명:'영업팀',
           // 양사간 거래처코드 (GW 위하고에 박힌 SMT 거래처)
           양사거래처코드:'000124', 양사거래처명:'에스엠티서울기연 주식회사' },
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



  // ===== 양사거래의 반대편 출고전표 (회계상 양사 매출/매입 짝 맞춤) =====
  // 예: SMT 발행 + GW 제품 양사거래 → GW가 SMT에 매출하는 출고전표 (GW 위하고용)
  global.WehagoExporter.makeChulgoCross = function(orders, company){
    if (!global.XLSX) return null;
    const def = WEHAGO_DEFAULTS[company] || WEHAGO_DEFAULTS.SMT;
    const topHeader = ['그룹번호','일자','거래처코드','부서코드','사원코드','관리항목코드','VAT여부값','과세구분','비고1','비고2','비고3'];
    const botHeader = ['그룹번호','품목코드','규격','납기일자','수량','단가','공급가액','부가세액','창고코드','프로젝트코드','품목비고','입고단가'];
    const topRows = [topHeader];
    const botRows = [botHeader];
    let groupNo = 0;
    // 양사거래에서 company = 매출하는 쪽 (발행회사의 반대)
    orders.forEach(o => {
      const intern = o._internalTransfer;
      if (!intern || !intern.confirmed) return;
      // 양사거래 발행회사 != company. company는 반대편 (매출 발생).
      if ((o.발행회사 || 'SMT') === company) return;
      groupNo++;
      // 양사간 거래처코드 (company 위하고에 박힌 발행회사)
      const partnerCode = def.양사거래처코드;
      topRows.push([
        groupNo, ymd(o.일자),
        partnerCode,  // 양사간 거래처코드 (부장님 확인 필요)
        def.부서코드, def.사원코드, '2', 1, 0,
        '양사내부 매출', '', '대상회사: ' + (o.발행회사 || 'SMT')
      ]);
      (o.품목 || []).forEach(item => {
        const qty = +item.수량 || 0;
        const price = +item._internalPrice || +item.단가 || 0;
        const supply = qty * price;
        const vat = Math.round(supply * 0.1);
        botRows.push([
          groupNo, item.품목코드 || '', item.규격 || '', ymd(o.일자),
          qty, price, supply, vat,
          def.창고코드, '', '양사거래', ''
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
    // 양사 내부거래 — 발행회사 = company. 발행회사가 다른 회사에서 매입 후 거래처로 출고.
    orders.forEach(o => {
      const intern = o._internalTransfer;
      if (!intern || !intern.confirmed) return;
      // 발행회사 = company만 (입고 받는 쪽 = 발행회사 자기)
      if ((o.발행회사 || 'SMT') !== company) return;
      const other = (company === 'SMT') ? 'GW' : 'SMT';  // 매입한 반대편 회사
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

  // ===== 3. 일반전표 — 매출 분개 (계정 동적 분기 + 5종 적요) =====
  // 매출 계정: EU/EM/EG/EK/EZ = 상품매출(혼다 수입) / SG/SP/SH/KG/XE = 제품매출(자체) / 그외 = 부품매출
  function _decideSalesAccount(item) {
    const name = ((item && (item.품목명 || item.모델)) || '').toUpperCase();
    if (/\b(EU|EM|EG|EK|EZ|ER|E1)\d/.test(name)) return { 계정:'상품매출', 코드:'401' };
    if (/\b(SG|SP|SH|KG|XE|EP|SHW)\d/.test(name)) return { 계정:'제품매출', 코드:'404' };
    return { 계정:'부품매출', 코드:'402' };
  }
  function _generateMemo(o){
    const items = o.품목 || [];
    if (items.length === 0) return o.거래처명 || '';
    const it = items[0];
    const qty = it.수량 || 1;
    const price = it.단가 || 0;
    const nm = it.품목명 || it.모델 || '';
    let memo = (qty === 1) ? `${nm} @${price.toLocaleString()}` : `${nm} ${qty}*@${price.toLocaleString()}`;
    if (items.length > 1) memo += ` 외 ${items.length-1}건`;
    return memo;
  }
  global.WehagoExporter.makeIlban = function(orders, company) {
    if (!global.XLSX) return null;
    const rows = [];
    rows.push(['월','일','구분','Code','계정과목','Code','거래처','적요','차변(출금)','대변(입금)']);
    orders.filter(o => (o.발행회사||'SMT') === company).forEach(o => {
      const { m, d } = splitDate(o.일자);
      // 매출 계정은 첫 품목 기준 결정 (대부분 한 분개 = 같은 계정)
      const acct = _decideSalesAccount((o.품목 || [])[0]);
      const totalSupply = (o.품목 || []).reduce((a,i) => a + ((i.수량||0)*(i.단가||0)), 0);
      const totalVat = Math.round(totalSupply * 0.1);
      const totalAmt = totalSupply + totalVat;
      const memo = _generateMemo(o);
      // 외상매출금 / 매출계정 + 부가세예수금
      rows.push([m, d, 3, '108', '외상매출금',    o.거래처코드 || '', o.거래처명, memo, totalAmt, '']);
      rows.push([m, d, 4, acct.코드, acct.계정, o.거래처코드 || '', o.거래처명, memo, '', totalSupply]);
      rows.push([m, d, 4, '255', '부가세예수금', o.거래처코드 || '', o.거래처명, memo, '', totalVat]);
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



  // ===== 매입 분개 (양사거래 시 입고측 회계 처리) =====
  // 차변 상품·부가세대급금 / 대변 외상매입금 (양사간 매입)
  global.WehagoExporter.makeIlbanMaeip = function(orders, company){
    if (!global.XLSX) return null;
    const def = WEHAGO_DEFAULTS[company] || WEHAGO_DEFAULTS.SMT;
    const rows = [['월','일','구분','Code','계정과목','Code','거래처','적요','차변(출금)','대변(입금)']];
    orders.forEach(o => {
      const intern = o._internalTransfer;
      if (!intern || !intern.confirmed) return;
      if ((o.발행회사 || 'SMT') !== company) return;  // 매입 = 발행회사 자기 입장
      const { m, d } = splitDate(o.일자);
      const totalSupply = (o.품목||[]).reduce((a,i) => a + (i.수량||0)*(i._internalPrice || i.단가||0), 0);
      const totalVat = Math.round(totalSupply * 0.1);
      const totalAmt = totalSupply + totalVat;
      const memo = `${def.양사거래처명} 양사매입 ${(o.품목||[])[0]?.품목명 || ''}${(o.품목||[]).length>1 ? ` 외 ${(o.품목||[]).length-1}건` : ''}`;
      // 차변 상품·부가세대급금 / 대변 외상매입금
      rows.push([m, d, 3, '146', '상품',       def.양사거래처코드, def.양사거래처명, memo, totalSupply, '']);
      rows.push([m, d, 3, '135', '부가세대급금', def.양사거래처코드, def.양사거래처명, memo, totalVat, '']);
      rows.push([m, d, 4, '251', '외상매입금',  def.양사거래처코드, def.양사거래처명, memo, '', totalAmt]);
    });
    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws['!cols'] = [{wch:5},{wch:5},{wch:6},{wch:8},{wch:15},{wch:9},{wch:25},{wch:35},{wch:14},{wch:14}];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '일반전표');
    return wb;
  };

  // ===== 매출 분개 (양사거래의 반대편 = 매출 발생 측, 예: GW가 SMT에 매출) =====
  global.WehagoExporter.makeIlbanCross = function(orders, company){
    if (!global.XLSX) return null;
    const def = WEHAGO_DEFAULTS[company] || WEHAGO_DEFAULTS.SMT;
    const rows = [['월','일','구분','Code','계정과목','Code','거래처','적요','차변(출금)','대변(입금)']];
    orders.forEach(o => {
      const intern = o._internalTransfer;
      if (!intern || !intern.confirmed) return;
      if ((o.발행회사 || 'SMT') === company) return;  // 매출 측 = 발행회사 반대편
      const { m, d } = splitDate(o.일자);
      const totalSupply = (o.품목||[]).reduce((a,i) => a + (i.수량||0)*(i._internalPrice || i.단가||0), 0);
      const totalVat = Math.round(totalSupply * 0.1);
      const totalAmt = totalSupply + totalVat;
      const memo = `${def.양사거래처명} 양사매출 ${(o.품목||[])[0]?.품목명 || ''}${(o.품목||[]).length>1 ? ` 외 ${(o.품목||[]).length-1}건` : ''}`;
      rows.push([m, d, 3, '108', '외상매출금',   def.양사거래처코드, def.양사거래처명, memo, totalAmt, '']);
      rows.push([m, d, 4, '401', '상품매출',      def.양사거래처코드, def.양사거래처명, memo, '', totalSupply]);
      rows.push([m, d, 4, '255', '부가세예수금', def.양사거래처코드, def.양사거래처명, memo, '', totalVat]);
    });
    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws['!cols'] = [{wch:5},{wch:5},{wch:6},{wch:8},{wch:15},{wch:9},{wch:25},{wch:35},{wch:14},{wch:14}];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '일반전표');
    return wb;
  };

  // ===== 양사 세금계산서 (양사거래의 매출 발생 측) =====
  global.WehagoExporter.makeSegyumetanCross = function(orders, company){
    if (!global.XLSX) return null;
    const def = WEHAGO_DEFAULTS[company] || WEHAGO_DEFAULTS.SMT;
    const rows = [['작성일자','거래처코드','거래처명','거래처사업자번호','대표자','업태','종목','품목명','규격','수량','단가','공급가액','부가세','합계금액','비고','전자세금계산서구분']];
    orders.forEach(o => {
      const intern = o._internalTransfer;
      if (!intern || !intern.confirmed) return;
      if ((o.발행회사 || 'SMT') === company) return;  // 매출 측 발급
      (o.품목 || []).forEach(item => {
        const supply = (item.수량 || 0) * (item._internalPrice || item.단가 || 0);
        const vat = Math.round(supply * 0.1);
        rows.push([
          o.일자, def.양사거래처코드, def.양사거래처명, '',
          '', '도매및소매업', '',
          item.품목명, item.규격 || '', item.수량,
          (item._internalPrice || item.단가), supply, vat, supply+vat,
          '양사간 매출', '01'
        ]);
      });
    });
    const ws = XLSX.utils.aoa_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '세금계산서');
    return wb;
  };

  // ===== 4. 세금계산서 (위하고 표준 84컬럼 풀 양식) =====
  // 시트1 = "엑셀예시양식" Row1=숫자 Row2=영역헤더 Row3=세부헤더 Row4=설명 Row5~=데이터
  // 우리는 Row1 (숫자) + Row2 (영역헤더) + 데이터만 작성
  global.WehagoExporter.makeSegyumetan = function(orders, company){
    if (!global.XLSX) return null;
    const co = COMPANIES[company] || COMPANIES.SMT;
    const rows = [SEGYU_HEADER_1, SEGYU_HEADER_2];
    let seq = 1;
    orders.filter(o => (o.발행회사||'SMT') === company).forEach(o => {
      // 공급받는자 정보 (거래처)
      const cli = (window._wehagoClients || []).find(c => (c.거래처명||c.상호) === o.거래처명) || {};
      const buyer = {
        등록번호: (cli['사업자(주민번호)'] || cli['사업자번호'] || '').replace(/-/g,''),
        상호: o.거래처명 || '',
        대표자명: cli['대표자명'] || cli['대표자'] || '',
        주소: cli['사업장주소'] || cli['주소'] || '',
        주소상세: '',
        종사업장번호: '',
        업태: cli['업태'] || '',
        종목: cli['종목'] || '',
        부서명: '',
        담당자이메일: cli['E-mail 주소'] || cli['E-mail'] || '',
        담당자휴대번호: (cli['담당자핸드폰'] || cli['담당자휴대전화'] || '').replace(/-/g,''),
        담당자명: cli['업체담당자명'] || cli['담당자명'] || '',
        담당자전화번호: (cli['담당자전화'] || cli['전화번호'] || cli['Fax번호'] || '').replace(/-/g,''),
      };
      const ymdStr = ymd(o.일자);
      const yyyy = ymdStr.slice(0,4), mm = ymdStr.slice(4,6);
      const seqStr = String(seq++).padStart(5,'0');
      const txId = 'TX' + ymdStr + seqStr;

      const items = (o.품목 || []).slice(0, 4);  // 최대 4종 (84컬럼 양식 제한)
      const totalSupply = (o.품목 || []).reduce((a,i) => a + ((i.수량||0)*(i.단가||0)), 0);
      const totalVat = Math.round(totalSupply * 0.1);
      const totalAmt = totalSupply + totalVat;
      const isOesang = ['월말','익월말','외상'].includes(o.결제조건 || '월말');
      const cheonggu = isOesang ? '1' : '2';  // 1:청구(외상) / 2:영수(즉시)

      // 84컬럼 데이터 작성
      const r = new Array(84).fill('');
      r[0] = txId;          // 발송시스템_고유번호
      r[1] = 'EXCEL';       // 결과_고유번호
      r[2] = cheonggu;      // 청구유형코드 (1:청구 / 2:영수)
      r[3] = '1';           // 거래구분코드 (1:매출)
      r[4] = '1';           // 사용자구분코드 (1:기업)
      r[5] = '1';           // 과세구분코드 (1:과세)
      r[6] = yyyy;          // 책번호_권
      r[7] = mm;            // 책번호_호
      r[8] = seqStr;        // 일련번호
      // 공급자정보 (10~22, idx 9~21)
      r[9]  = co.등록번호_원.replace(/-/g,'');  // 등록번호
      r[10] = co.명;        // 상호
      r[11] = co.대표;       // 대표자명
      r[12] = co.주소;       // 주소
      r[13] = co.주소상세 || '';
      r[14] = '';            // 종사업장번호
      r[15] = co.업태;
      r[16] = co.종목;
      r[17] = '';            // 부서명
      r[18] = co.담당자이메일 || '';
      r[19] = (co.담당자전화 || '').replace(/-/g,'');  // 담당자휴대번호 위치
      r[20] = co.담당자명 || '';
      r[21] = (co.담당자전화 || '').replace(/-/g,'');  // 담당자전화번호
      // 공급받는자정보 (23~35, idx 22~34)
      r[22] = buyer.등록번호;
      r[23] = buyer.상호;
      r[24] = buyer.대표자명;
      r[25] = buyer.주소;
      r[26] = buyer.주소상세;
      r[27] = buyer.종사업장번호;
      r[28] = buyer.업태;
      r[29] = buyer.종목;
      r[30] = buyer.부서명;
      r[31] = buyer.담당자이메일;
      r[32] = buyer.담당자휴대번호;
      r[33] = buyer.담당자명;
      r[34] = buyer.담당자전화번호;
      // 작성일자 / 공급가액 / 세액 (36~38, idx 35~37)
      r[35] = ymdStr;
      r[36] = totalSupply;
      r[37] = totalVat;
      // 품목1~4 (각 8컬럼 — idx 38~69)
      for (let k=0; k<4; k++) {
        const base = 38 + k*8;
        const it = items[k];
        if (!it) continue;
        const sup = (it.수량||0)*(it.단가||0);
        const v = Math.round(sup*0.1);
        r[base+0] = ymdStr;                  // 일자
        r[base+1] = it.품목명 || it.모델 || '';   // 품목
        r[base+2] = it.규격 || '';             // 규격
        r[base+3] = it.수량 || 0;              // 수량
        r[base+4] = it.단가 || 0;              // 단가
        r[base+5] = sup;                       // 공급가액
        r[base+6] = v;                          // 세액
        r[base+7] = it.품목비고 || '';          // 비고
      }
      // 합계금액 + 결제 (71~75, idx 70~74)
      r[70] = totalAmt;
      // 현금/수표/어음/외상미수금 — 외상이면 외상미수금에 합계
      if (isOesang) {
        r[74] = totalAmt;  // 외상미수금
      } else {
        r[71] = totalAmt;  // 현금
      }
      // 비고1·2·3 (76~78, idx 75~77)
      r[75] = o.비고 || '';
      r[76] = '';
      r[77] = '';
      // 수정세금계산서 — 디폴트 N
      r[78] = 'N';
      r[79] = '';  // 수정코드
      r[80] = '';  // 원천 고유번호
      r[81] = '';  // 원천 승인번호
      r[82] = '';  // 협력업체코드
      r[83] = 'END';  // 종료
      rows.push(r);
    });
    const ws = XLSX.utils.aoa_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '엑셀예시양식');
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

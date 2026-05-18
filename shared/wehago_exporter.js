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

  // ===== 1. 입고전표 엑셀 생성 =====
  // 양사 간 내부거래(상대방으로부터 입고) 또는 일반 매입
  // 컬럼: 일자 | 거래처코드 | 거래처명 | 품목코드 | 품목명 | 규격 | 단위 | 수량 | 단가 | 공급가액 | 부가세
  global.WehagoExporter = global.WehagoExporter || {};
  
  global.WehagoExporter.makeIpgo = function(orders, company) {
    if (!global.XLSX) return null;
    const rows = [];
    rows.push(['일자','거래처코드','거래처명','품목코드','품목명','규격','단위','수량','단가','공급가액','부가세','비고']);
    
    orders.forEach(o => {
      // 양사 간 내부거래만 입고전표 생성
      if (o._internalTransfer && o._internalTransfer.toCompany === company) {
        o.품목.forEach(item => {
          if (item._fromCompany && item._fromCompany !== company) {
            const supply = (item.수량 || 0) * (item._internalPrice || item.단가 || 0);
            const vat = Math.round(supply * 0.1);
            rows.push([
              o.일자,
              o._fromCompanyCode || '',
              COMPANIES[item._fromCompany]?.명 || '',
              item.품목코드,
              item.품목명,
              item.규격 || '',
              item.단위 || '',
              item.수량,
              item._internalPrice || item.단가,
              supply,
              vat,
              `양사간이동: ${item._fromCompany}→${company}`
            ]);
          }
        });
      }
    });
    
    const ws = XLSX.utils.aoa_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '입고전표');
    return wb;
  };

  // ===== 2. 출고전표 엑셀 생성 =====
  // 모든 매출 거래 (영업맨이 입력한 모든 주문)
  global.WehagoExporter.makeChulgo = function(orders, company) {
    if (!global.XLSX) return null;
    const rows = [];
    rows.push(['일자','거래처코드','거래처명','품목코드','품목명','규격','단위','수량','단가','공급가액','부가세','담당사원','비고']);
    
    orders.filter(o => o.발행회사 === company).forEach(o => {
      o.품목.forEach(item => {
        const supply = (item.수량 || 0) * (item.단가 || 0);
        const vat = Math.round(supply * 0.1);
        rows.push([
          o.일자,
          o.거래처코드,
          o.거래처명,
          item.품목코드 || '',
          item.품목명,
          item.규격 || '',
          item.단위 || '',
          item.수량,
          item.단가,
          supply,
          vat,
          o.영업맨,
          o.비고 || ''
        ]);
      });
    });
    
    const ws = XLSX.utils.aoa_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '출고전표');
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

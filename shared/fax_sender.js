/* SMT·건웅 — 팩스 발송 래퍼 (2026-05-26)
   현재: PDF 다운로드 + 안내 (수동 발송)
   향후: 엔팩스 API 연동 — sendByApi() 함수만 교체하면 됨
*/
(function(global){
  'use strict';

  // ===== 설정 =====
  const CONFIG = {
    // 양사 공통 발신 팩스
    senderFaxNumber: '031-353-4727',
    // 엔팩스 API 키 (가입 + 발신번호 승인 후 채움)
    enfaxApiKey: '',  // TODO: 받으면 여기에
    enfaxApiUrl: 'https://api.enfax.com/v1/send',  // 실제 엔드포인트 가입 후 확인
    enfaxSenderName: '',  // SMT 또는 GW
  };

  // 택배사 팩스번호 사전 (부장님 입력 시 채움)
  // localStorage에 저장 → UI에서 변경 가능
  function getCourierFaxMap(){
    try {
      const saved = JSON.parse(localStorage.getItem('courier_fax_map') || '{}');
      return {
        '대신화물': saved['대신화물'] || '',
        '대신택배': saved['대신택배'] || '',
        '경동화물': saved['경동화물'] || '',
        '경동택배': saved['경동택배'] || '',
        ...saved
      };
    } catch(e) { return {}; }
  }
  function setCourierFax(name, faxNumber){
    const m = getCourierFaxMap();
    m[name] = faxNumber.trim();
    localStorage.setItem('courier_fax_map', JSON.stringify(m));
  }

  // ===== 발송 로그 (localStorage) =====
  function logFax(record){
    try {
      const logs = JSON.parse(localStorage.getItem('fax_logs') || '[]');
      logs.unshift({
        _ts: Date.now(),
        ...record
      });
      localStorage.setItem('fax_logs', JSON.stringify(logs.slice(0, 200)));  // 최근 200건
    } catch(e){}
  }
  function getFaxLogs(){
    try { return JSON.parse(localStorage.getItem('fax_logs') || '[]'); } catch(e){ return []; }
  }

  // ===== 1. 현재 모드 (수동 발송 — PDF 다운로드) =====
  async function sendByDownload(pdfBlob, fileName, meta){
    // PDF 다운로드 트리거
    const url = URL.createObjectURL(pdfBlob);
    const a = document.createElement('a');
    a.href = url; a.download = fileName;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    URL.revokeObjectURL(url);
    // 로그
    logFax({...meta, method: 'manual_download', status: 'downloaded'});
    return { success: true, mode: 'manual', message: 'PDF 다운로드 완료. 모바일팩스 앱으로 발송해주세요.' };
  }

  // ===== 2. 향후 모드 (엔팩스 API 자동 발송) =====
  // 가입·API 키 받으면 활성화 — 이 함수만 채우면 자동 발송 전환
  async function sendByApi(pdfBlob, recipientFax, meta){
    if (!CONFIG.enfaxApiKey) {
      console.warn('[fax_sender] 엔팩스 API 키 미설정 — 수동 모드로 전환');
      return sendByDownload(pdfBlob, meta.fileName || 'fax.pdf', meta);
    }
    // TODO: 엔팩스 API 호출
    // FormData 또는 base64 변환 후 POST
    /*
    const fd = new FormData();
    fd.append('apikey', CONFIG.enfaxApiKey);
    fd.append('sender', CONFIG.senderFaxNumber);
    fd.append('receiver', recipientFax);
    fd.append('file', pdfBlob, meta.fileName);
    fd.append('title', meta.title || '발송통지서');
    const res = await fetch(CONFIG.enfaxApiUrl, { method:'POST', body: fd });
    const data = await res.json();
    logFax({...meta, method:'enfax_api', status: data.success ? 'sent' : 'failed', tx_id: data.tx_id, error: data.error});
    return { success: data.success, mode:'api', tx_id: data.tx_id, message: data.message };
    */
    // 키 받으면 위 주석 풀고 엔드포인트·필드명만 엔팩스 가이드대로 조정
    return { success: false, mode: 'api_not_configured', message: 'API 키 미설정' };
  }

  // ===== 통합 send 함수 — 현재는 PDF 다운로드, 향후 API 자동 전환 =====
  async function send(pdfBlob, recipientFax, meta = {}) {
    meta.fileName = meta.fileName || `발송통지서_${new Date().toISOString().slice(0,10)}.pdf`;
    meta.recipientFax = recipientFax;
    // 모드 자동 결정
    if (CONFIG.enfaxApiKey && recipientFax) {
      return sendByApi(pdfBlob, recipientFax, meta);
    } else {
      return sendByDownload(pdfBlob, meta.fileName, meta);
    }
  }

  // ===== Export =====
  global.FaxSender = {
    CONFIG, send, sendByDownload, sendByApi,
    getCourierFaxMap, setCourierFax,
    logFax, getFaxLogs
  };
})(window);

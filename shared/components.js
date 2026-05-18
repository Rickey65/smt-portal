/* SMT·건웅 통합 시스템 — 공통 UI 컴포넌트 v2 */
(function(global) {
  'use strict';

  // ===== 회사 토글 위젯 =====
  global.CompanyToggle = {
    
    /**
     * 화면 우상단에 회사 토글 위젯 삽입
     * @param {string} containerId - 토글을 삽입할 컨테이너 ID (옵션, 없으면 자동 fixed)
     * @param {Function} onChange - 토글 변경 시 콜백 (옵션, 페이지 리로드가 기본)
     */
    render(containerId, onChange) {
      const html = `
        <div class="company-toggle">
          <button data-company="ALL"  class="ct-btn">통합</button>
          <button data-company="SMT"  class="ct-btn">SMT</button>
          <button data-company="GW"   class="ct-btn">건웅</button>
        </div>
      `;
      
      let container;
      if (containerId) {
        container = document.getElementById(containerId);
        if (container) container.innerHTML = html;
      } else {
        // 우상단 fixed 위젯
        const wrap = document.createElement('div');
        wrap.id = 'company-toggle-fixed';
        wrap.innerHTML = html;
        document.body.appendChild(wrap);
        container = wrap;
      }
      
      // 활성 상태 표시
      this.updateActive();
      
      // 클릭 이벤트
      container.querySelectorAll('.ct-btn').forEach(btn => {
        btn.addEventListener('click', () => {
          const v = btn.dataset.company;
          Company.set(v);
          this.updateActive();
          if (typeof onChange === 'function') {
            onChange(v);
          } else {
            // 기본: 페이지 새로고침
            location.reload();
          }
        });
      });
    },
    
    updateActive() {
      const current = Company.get();
      document.querySelectorAll('.company-toggle .ct-btn').forEach(btn => {
        if (btn.dataset.company === current) {
          btn.classList.add('active');
        } else {
          btn.classList.remove('active');
        }
      });
    }
  };

  // ===== 숫자 포맷 헬퍼 =====
  global.Fmt = {
    won(n) {
      if (!n || isNaN(n)) return '0원';
      n = Math.round(n);
      if (Math.abs(n) >= 100000000) return (n / 100000000).toFixed(1) + '억';
      if (Math.abs(n) >= 10000) return Math.round(n / 10000).toLocaleString() + '만';
      return n.toLocaleString() + '원';
    },
    wonFull(n) {
      if (!n || isNaN(n)) return '0';
      return Math.round(n).toLocaleString();
    },
    pct(n) {
      if (!n || isNaN(n)) return '0%';
      return (n > 0 ? '+' : '') + n.toFixed(1) + '%';
    },
    date(d) {
      if (!d) return '';
      return String(d).slice(0, 10);
    }
  };

  // ===== 음성 출력 (TTS) — 어르신 친화 =====
  global.Voice = {
    enabled: true,
    
    speak(text) {
      if (!this.enabled || !text || !window.speechSynthesis) return;
      // 기존 음성 중지
      speechSynthesis.cancel();
      setTimeout(() => {
        const u = new SpeechSynthesisUtterance(text);
        u.lang = 'ko-KR';
        u.rate = 1.0;
        u.pitch = 1.0;
        // 한국어 음성 선택
        const voices = speechSynthesis.getVoices();
        const koVoice = voices.find(v => v.lang.startsWith('ko'));
        if (koVoice) u.voice = koVoice;
        speechSynthesis.speak(u);
        // Chrome 15초 무음 버그 회피
        const interval = setInterval(() => {
          if (!speechSynthesis.speaking) {
            clearInterval(interval);
          } else {
            speechSynthesis.pause();
            speechSynthesis.resume();
          }
        }, 8000);
      }, 50);
    },
    
    toggle() {
      this.enabled = !this.enabled;
      if (!this.enabled) speechSynthesis.cancel();
      return this.enabled;
    }
  };

})(window);

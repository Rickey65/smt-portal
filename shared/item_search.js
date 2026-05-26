/* 품목 한·영 검색 + 카테고리 자동 분류 — 2026-05-26 1차 */
(function(global){
  'use strict';

  let CATEGORIES = null;
  let SYNONYMS = null;
  let _itemCatCache = {};  // 품목명 → {cat, sub, brand}

  async function load(){
    if (CATEGORIES) return CATEGORIES;
    try {
      const r = await fetch('data/shared/item_categories.json?t='+Date.now());
      const d = await r.json();
      CATEGORIES = d;
      SYNONYMS = d.synonyms || {};
      return d;
    } catch(e){ console.warn('item_categories.json load fail',e); return null; }
  }

  // 검색어 확장: 한국어 → 영어 + 영어 → 한국어 양방향
  function expandQuery(q){
    if (!q) return [];
    const tokens = q.trim().toLowerCase().split(/\s+/);
    const expanded = new Set();
    tokens.forEach(t => {
      expanded.add(t);
      // 한국어 → 영어
      if (SYNONYMS && SYNONYMS[t]) SYNONYMS[t].forEach(s => expanded.add(s.toLowerCase()));
      // 한국어 키 부분일치
      if (SYNONYMS) Object.keys(SYNONYMS).forEach(k => {
        if (k.toLowerCase().includes(t)) SYNONYMS[k].forEach(s => expanded.add(s.toLowerCase()));
      });
      // 영어 → 한국어 (synonyms의 값에서 t를 찾아 키 반환)
      if (SYNONYMS) Object.entries(SYNONYMS).forEach(([k,vs]) => {
        vs.forEach(v => {
          if (v.toLowerCase() === t || v.toLowerCase().includes(t)) expanded.add(k.toLowerCase());
        });
      });
    });
    return [...expanded];
  }

  // 품목 분류: 품목명 → 카테고리/소분류
  function classify(itemName){
    if (!itemName || !CATEGORIES) return null;
    if (_itemCatCache[itemName]) return _itemCatCache[itemName];
    const upper = itemName.toUpperCase();
    for (const big of CATEGORIES.tree){
      for (const sub of (big.sub || [])){
        // prefix 매치
        if (sub.prefix){
          for (const p of sub.prefix){
            if (upper.startsWith(p)) {
              const r = {cat: big.cat, sub: sub.cat, brand: sub.brand || ''};
              _itemCatCache[itemName] = r;
              return r;
            }
          }
        }
        // name_kw 매치
        if (sub.name_kw){
          for (const kw of sub.name_kw){
            if (itemName.includes(kw) || upper.includes(kw.toUpperCase())) {
              const r = {cat: big.cat, sub: sub.cat, brand: sub.brand || ''};
              _itemCatCache[itemName] = r;
              return r;
            }
          }
        }
      }
    }
    _itemCatCache[itemName] = null;
    return null;
  }

  // 품목 매칭: 품목 객체 + 검색어 → boolean
  function matchItem(item, queryTokens){
    if (!queryTokens || queryTokens.length === 0) return true;
    const name = (item.품목명 || '').toLowerCase();
    const code = (item.품목코드 || '').toLowerCase();
    const spec = (item.규격 || '').toLowerCase();
    const cls = classify(item.품목명 || '');
    const cat  = (cls && cls.cat)  ? cls.cat.toLowerCase()  : '';
    const sub  = (cls && cls.sub)  ? cls.sub.toLowerCase()  : '';
    const brand= (cls && cls.brand)? cls.brand.toLowerCase(): '';
    const hay = name + ' ' + code + ' ' + spec + ' ' + cat + ' ' + sub + ' ' + brand;
    // 모든 토큰이 매치 (AND) — 그러나 expanded 안에서는 OR
    // 여기서는 단순 OR (어느 한 토큰이라도 hit이면 매치)
    return queryTokens.some(t => t && hay.includes(t));
  }

  // 카테고리 트리 반환 (UI에서 그릴 때 사용)
  function getTree(){ return CATEGORIES ? CATEGORIES.tree : []; }
  function getBrands(){ return CATEGORIES ? CATEGORIES.brands : {}; }

  global.ItemSearch = { load, expandQuery, classify, matchItem, getTree, getBrands };
})(window);

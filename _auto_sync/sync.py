# -*- coding: utf-8 -*-
"""
SMT서울기연·건웅 양사 위하고 데이터 자동 동기화 스크립트

매일 사용법:
  1) 위하고에서 양사 판매일보·현재고·미수채권 엑셀 다운로드
  2) _auto_sync/위하고다운로드/서울기연/ 또는 /건웅/ 폴더에 떨구기
  3) sync.bat 더블클릭

자동으로:
  - 엑셀 → JSON 변환 (옛 단일 + 새 양사분리 양식 둘 다)
  - dist/에 덮어쓰기
  - git add / commit / push
  - Vercel이 자동 배포

요구사항: Python 3.9+, openpyxl, git
"""
import os, sys, json, re, glob, shutil, subprocess
from datetime import datetime
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_DIR = SCRIPT_DIR.parent
DIST_DIR = PROJECT_DIR  # repo root
DATA_PC = PROJECT_DIR / "data" / "per_company"
# 다운로드 폴더 자동 감지:
# 1순위) 클로드 작업폴더 (부장님이 매일 떨구는 곳)
# 2순위) GitHub 폴더의 _auto_sync (기본 위치)
_CLAUDE_FOLDER = Path(r"C:\Users\lsw05\작업\클로드\01_SMT서울기연\_영업(2026)\세션 3 SMT서울기연 - 영업\_auto_sync\위하고다운로드")
if _CLAUDE_FOLDER.exists() and any(_CLAUDE_FOLDER.glob("*/*.xls*")):
    DOWNLOAD_DIR = _CLAUDE_FOLDER
    print(f"[정보] 클로드 작업폴더의 자료 사용: {DOWNLOAD_DIR}")
else:
    DOWNLOAD_DIR = SCRIPT_DIR / "위하고다운로드"
DONE_DIR = DOWNLOAD_DIR / "_처리완료"
DONE_DIR.mkdir(exist_ok=True)

TODAY = datetime.now().strftime("%Y-%m-%d")
NOW = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

print("=" * 60)
print(f"  SMT·건웅 양사 데이터 동기화  ({NOW})")
print("=" * 60)

try:
    import openpyxl
except ImportError:
    print("\n[오류] openpyxl 모듈이 없습니다.")
    print("      cmd에서 다음 명령 실행: pip install openpyxl")
    input("\n엔터를 누르면 종료...")
    sys.exit(1)


def read_xlsx(path):
    """xls/xlsx 양쪽 지원. xlsx는 openpyxl, xls는 xlrd."""
    path_str = str(path).lower()
    if path_str.endswith('.xlsx'):
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        ws = wb.active
        aoa = []
        for row in ws.iter_rows(values_only=True):
            aoa.append([("" if v is None else v) for v in row])
        wb.close()
        return aoa
    elif path_str.endswith('.xls'):
        try:
            import xlrd
        except ImportError:
            import subprocess, sys
            print("    [정보] xlrd 자동 설치 중...")
            subprocess.run([sys.executable, '-m', 'pip', 'install', 'xlrd==1.2.0', '--quiet'], check=False)
            import xlrd
        wb = xlrd.open_workbook(str(path))
        sh = wb.sheet_by_index(0)
        aoa = []
        for i in range(sh.nrows):
            row = sh.row_values(i)
            aoa.append([("" if v in (None, '') else v) for v in row])
        return aoa
    else:
        raise ValueError(f"지원하지 않는 파일 형식: {path}")


def s(v):
    return "" if v is None else str(v).strip()


def num(v):
    try:
        if v == "" or v is None:
            return 0
        return float(v)
    except Exception:
        return 0


def detect_type(aoa):
    flat_raw = " ".join(s(c) for row in aoa[:30] for c in row)
    flat = flat_raw.replace(" ", "")  # 공백 무시 (위하고 "분 개 장" 같은 띄어쓰기 대응)
    # 미수채권상세현황
    if "미수채권상세현황" in flat and "당기발생" in flat and "당기수금" in flat:
        return "receivable_detail"
    # 채권채무집계
    if ("채권" in flat and "채무" in flat and "거래처" in flat and "잔액" in flat
        and "미수채권상세현황" not in flat):
        return "receivable"
    # 분개장
    if "분개장" in flat and "차변" in flat and "대변" in flat:
        return "journal"
    # 현재고
    if "품목계정" in flat and "기말재고" in flat and "기초수량" in flat:
        return "inventory"
    # 판매일보 (신·구 모두 — "일자" + "품목" + "공급가액")
    if "일자" in flat and ("품목명" in flat or "품목코드" in flat) and "공급가액" in flat:
        return "sales"
    return None


def parse_sales(aoa, company):
    """판매일보 파서 — 신·구 양식 자동 인식. 헤더 위치 + 컬럼 인덱스 동적 추출."""
    # 헤더 행 찾기 ("일자" 컬럼이 있는 첫 행)
    hi = -1
    headers = []
    for i, row in enumerate(aoa[:50]):
        cells = [s(c) for c in row]
        if "일자" in cells and ("품목명" in cells or "품목코드" in cells):
            hi = i
            headers = cells
            break
    if hi < 0:
        raise ValueError("판매일보 헤더(일자/품목명) 못 찾음")
    # 동적 컬럼 인덱스
    def col(name, alt=None):
        if name in headers:
            return headers.index(name)
        if alt:
            for a in alt if isinstance(alt, list) else [alt]:
                if a in headers: return headers.index(a)
        return -1
    idx = {
        "일자": col("일자"),
        "품목계정": col("품목계정"),
        "품목코드": col("품목코드"),
        "품목명": col("품목명", "품명"),
        "규격": col("규격"),
        "수량": col("수량"),
        "단가": col("단가"),
        "공급가액": col("공급가액"),
        "부가세": col("부가세"),
        "합계금액": col("합계금액"),
        "거래처명": col("거래처명", "거래처"),
        "유형": col("유형"),
        "사원": col("사원", "담당자"),
        "창고명": col("창고명", "창고"),
    }
    rows = []
    for r in aoa[hi + 1:]:
        if not r or len(r) < 3: continue
        date_raw = s(r[idx["일자"]]) if idx["일자"] >= 0 and idx["일자"] < len(r) else ""
        if not date_raw or date_raw in ("일자", "합계", "소계"): continue
        if not re.match(r"\d{4}[-./]\d{2}[-./]\d{2}", date_raw): continue
        date = date_raw[:10].replace(".", "-").replace("/", "-")
        def get(k, default=""):
            i_ = idx.get(k, -1)
            return r[i_] if 0 <= i_ < len(r) else default
        rows.append({
            "일자": date, "_month": date[:7],
            "품목계정": s(get("품목계정")),
            "품목코드": s(get("품목코드")),
            "품목명": s(get("품목명")),
            "규격": s(get("규격")),
            "수량": num(get("수량", 0)),
            "단가": num(get("단가", 0)),
            "공급가액": num(get("공급가액", 0)),
            "부가세": num(get("부가세", 0)),
            "합계금액": num(get("합계금액", 0)),
            "거래처명": s(get("거래처명")),
            "유형": s(get("유형")) or "판매출고",
            "담당자": s(get("사원")),
            "창고": s(get("창고명")),
            "_company": company,
        })
    return rows


def parse_inventory(aoa, company):
    hi = -1
    for i, row in enumerate(aoa):
        if len(row) >= 2 and s(row[0]) == "품목계정" and s(row[1]) == "품목코드":
            hi = i
            break
    if hi < 0:
        raise ValueError("현재고 헤더(품목계정/품목코드)를 찾지 못함")
    rows = []
    for r in aoa[hi + 1:]:
        if not r:
            continue
        cat, code, name = s(r[0]), s(r[1]), s(r[2])
        if not code or code == "품목코드":
            continue
        rows.append({
            "품목계정": cat, "품목코드": code, "품목명": name,
            "규격": s(r[3] if len(r) > 3 else ""),
            "기초수량": num(r[4] if len(r) > 4 else 0),
            "입고": num(r[5] if len(r) > 5 else 0),
            "출고": num(r[6] if len(r) > 6 else 0),
            "기말재고": num(r[7] if len(r) > 7 else 0),
            "_company": company,
        })
    return rows



def parse_journal(aoa, company):
    """분개장 양식: 6행부터 데이터 (월/일·번호·차변금액·차변계정·대변계정·대변금액)"""
    rows = []
    current = None
    year = datetime.now().year
    # 연도 자동 감지
    for r in aoa[:10]:
        for c in r:
            m = re.search(r"(\d{4})년", s(c))
            if m:
                year = int(m.group(1)); break
        if year != datetime.now().year: break
    
    for r in aoa[5:]:
        if not r or len(r) < 6: continue
        date_s = s(r[0])
        no_s = s(r[1])
        dr_amt = num(r[2])
        dr_acct = s(r[3])
        cr_acct = s(r[4])
        cr_amt = num(r[5])
        
        if date_s and re.match(r'\d{1,2}/\d{1,2}', date_s):
            if current:
                rows.append(current)
            mm, dd = date_s.split('/')
            current = {
                '_company': company,
                '일자': f'{year}-{int(mm):02d}-{int(dd):02d}',
                '번호': no_s, '차변': [], '대변': [], '적요': '', '거래처': ''
            }
            if dr_amt > 0: current['차변'].append({'계정': dr_acct, '금액': dr_amt})
            if cr_amt > 0: current['대변'].append({'계정': cr_acct, '금액': cr_amt})
        elif current is not None:
            if dr_amt > 0 and dr_acct:
                current['차변'].append({'계정': dr_acct, '금액': dr_amt})
            if cr_amt > 0 and cr_acct:
                current['대변'].append({'계정': cr_acct, '금액': cr_amt})
            if dr_amt == 0 and cr_amt == 0:
                if dr_acct and not current['적요']: current['적요'] = dr_acct
                if cr_acct and not current['거래처']: current['거래처'] = cr_acct
    if current: rows.append(current)
    return rows


def parse_receivable_detail(aoa, company):
    """미수채권상세현황 양식 — 거래처 블록 반복"""
    clients = {}
    current = None
    for i in range(len(aoa)):
        row = aoa[i]
        c0 = s(row[0]) if row else ''
        c6 = s(row[6]) if len(row)>6 else ''
        # 거래처 헤더 검색
        m = re.search(r'거래처\s*:\s*(\d+)\[(.+?)$', c6)
        if not m and c0.startswith('회사명'):
            m = re.search(r'거래처\s*:\s*(\d+)\[(.+?)$', c6 if c6 else s(row[6] if len(row)>6 else ''))
        if m:
            code, name = m.group(1), m.group(2).rstrip(']').strip()
            current = name
            if name not in clients:
                clients[name] = {'_company':company, '거래처코드':code, '거래처명':name,
                                 '전기이월':0, '매출':[], '수금':[], '잔액추이':[]}
            continue
        if not current: continue
        c = clients[current]
        if c0.startswith('전기'):
            c['전기이월'] = num(row[1] if len(row)>1 else 0)
            continue
        if re.match(r'\d{4}-\d{2}-\d{2}', c0):
            date = c0[:10]
            bal = num(row[12] if len(row)>12 else 0)
            공급 = num(row[7] if len(row)>7 else 0)
            부가 = num(row[8] if len(row)>8 else 0)
            합계 = num(row[9] if len(row)>9 else 0)
            유형 = s(row[10] if len(row)>10 else '')
            수금금액 = num(row[11] if len(row)>11 else 0)
            품목 = s(row[3] if len(row)>3 else '')
            if 공급 != 0 and 품목 and '★입금계좌번호★' not in s(row[2] if len(row)>2 else ''):
                c['매출'].append({'일자':date, '공급가액':공급, '부가세':부가, '합계':합계 or 공급+부가, '품목':품목})
            if 유형 and 수금금액:
                c['수금'].append({'일자':date, '유형':유형, '금액':수금금액})
            if bal != 0:
                c['잔액추이'].append({'일자':date, '잔액':bal})
    return list(clients.values())


def parse_receivable(aoa, company):
    """위하고 채권채무집계 양식
    행 1: 헤더1 (No, 거래처코드, 거래처, ..., 채권, 채무)
    행 2: 헤더2 (전기이월, 당기발생, 선수금, 당기수금, 잔액, ...)
    행 3+: 데이터
    열: 0=No, 1=거래처코드, 2=거래처, 5=채권전기, 6=채권당기, 7=선수금,
        8=채권수금, 9=채권잔액(미수금), 10~=채무"""
    rows = []
    for r in aoa[2:]:
        if not r:
            continue
        no = s(r[0])
        if not no:
            continue
        # No가 숫자가 아니면 건너뜀 (소계/합계 행 등)
        try:
            int(float(no))
        except Exception:
            continue
        rows.append({
            "거래처코드": s(r[1] if len(r) > 1 else ""),
            "거래처명": s(r[2] if len(r) > 2 else ""),
            "관리항목": s(r[4] if len(r) > 4 else ""),
            "채권_전기이월": num(r[5] if len(r) > 5 else 0),
            "채권_당기발생": num(r[6] if len(r) > 6 else 0),
            "채권_선수금": num(r[7] if len(r) > 7 else 0),
            "채권_당기수금": num(r[8] if len(r) > 8 else 0),
            "채권_잔액": num(r[9] if len(r) > 9 else 0),
            "채무_전기이월": num(r[10] if len(r) > 10 else 0),
            "채무_당기발생": num(r[11] if len(r) > 11 else 0),
            "채무_당기수금": num(r[13] if len(r) > 13 else 0),
            "채무_잔액": num(r[14] if len(r) > 14 else 0),
            "미수금": num(r[9] if len(r) > 9 else 0),  # 호환성
            "_company": company,
        })
    return rows


def parse_receivable_old(aoa, company):
    """옛 SHBM 미수채권상세현황 (호환 유지)"""
    rows = []
    cur, bal = None, 0
    for r in aoa:
        if not r:
            continue
        c0 = s(r[0])
        m = re.search(r"거래처\s*:\s*\d+\s*\[(.+?)\]", c0)
        if m:
            if cur and abs(bal) > 0:
                rows.append({"거래처명": cur, "미수금": bal, "_company": company})
            cur = m.group(1).strip()
            bal = 0
            continue
        if "잔액" in c0 or "합계" in c0:
            for v in r[1:]:
                n = num(v)
                if n != 0:
                    bal = n
            continue
    if cur and abs(bal) > 0:
        rows.append({"거래처명": cur, "미수금": bal, "_company": company})
    return rows


PARSERS = {"sales": parse_sales, "inventory": parse_inventory,
           "receivable": parse_receivable, "receivable_old": parse_receivable_old,
           "receivable_detail": parse_receivable_detail, "journal": parse_journal}
TYPE_NAMES = {"sales": "판매일보", "inventory": "현재고",
              "receivable": "채권채무집계", "receivable_old": "미수채권(옛)",
              "receivable_detail": "미수채권상세현황", "journal": "분개장"}


def main():
    by_company = {"SMT": [], "GW": []}
    for company, folder in [("SMT", "서울기연"), ("GW", "건웅")]:
        path = DOWNLOAD_DIR / folder
        if not path.exists():
            print(f"  ! 폴더 없음: {path}")
            continue
        files = list(path.glob("*.xlsx")) + list(path.glob("*.xls"))
        by_company[company] = files
        print(f"\n[{folder}] {len(files)}개 파일")
        for f in files:
            print(f"    - {f.name}")

    total_files = sum(len(v) for v in by_company.values())
    if total_files == 0:
        print("\n[오류] 위하고다운로드/서울기연 또는 /건웅 폴더에 엑셀이 없습니다.")
        print(f"        경로: {DOWNLOAD_DIR}")
        input("\n엔터를 누르면 종료...")
        return

    parsed = {"SMT": {"sales": [], "inventory": [], "receivable": [], "receivable_old": [],
                      "receivable_detail": [], "journal": []},
              "GW":  {"sales": [], "inventory": [], "receivable": [], "receivable_old": [],
                      "receivable_detail": [], "journal": []}}
    processed_files = []
    for company, files in by_company.items():
        for f in files:
            try:
                print(f"\n  변환: [{company}] {f.name} ...", end=" ")
                aoa = read_xlsx(f)
                t = detect_type(aoa)
                if t is None:
                    print("양식 감지 실패 (skip)")
                    continue
                rows = PARSERS[t](aoa, company)
                # receivable_old(옛 양식)는 receivable에 병합
                key = "receivable" if t == "receivable_old" else t
                parsed[company][key].extend(rows)
                print(f"{TYPE_NAMES[t]} {len(rows)}건 OK")
                processed_files.append(f)
            except Exception as e:
                print(f"오류: {e}")

    print("\n[저장] 양사 분리 (data/per_company/) — 사이트 로더 표준 포맷")
    DATA_PC.mkdir(parents=True, exist_ok=True)
    for company in ("SMT", "GW"):
        cl = company.lower()
        # 판매: 년도별 분리 + {rows:[...]} 래핑 + 기존 파일 일자 단위 병합
        if parsed[company]["sales"]:
            by_year = {}
            for r in parsed[company]["sales"]:
                ystr = (r.get("_month") or r.get("일자") or "")[:4]
                if not ystr.isdigit():
                    continue
                by_year.setdefault(int(ystr), []).append(r)
            for y, rows_y in by_year.items():
                outp = DATA_PC / f"sales_{cl}_{y}.json"
                existing_rows = []
                if outp.exists():
                    try:
                        with open(outp, encoding="utf-8") as fp:
                            ed = json.load(fp)
                        if isinstance(ed, dict) and "rows" in ed:
                            existing_rows = ed["rows"]
                        elif isinstance(ed, list):
                            existing_rows = ed
                    except Exception:
                        existing_rows = []
                new_dates = {r.get("일자") for r in rows_y}
                merged_rows = [r for r in existing_rows if r.get("일자") not in new_dates] + rows_y
                merged_rows.sort(key=lambda x: x.get("일자", ""))
                wrapped = {
                    "_company": company, "_year": y, "_count": len(merged_rows),
                    "_generated": NOW, "rows": merged_rows,
                }
                with open(outp, "w", encoding="utf-8") as fp:
                    json.dump(wrapped, fp, ensure_ascii=False, separators=(",", ":"))
                print(f"    {outp.name} (총 {len(merged_rows)}건 / 신규 {len(rows_y)}건)")
        # 재고: {rows:[...]} 래핑 + 분류 머지
        if parsed[company]["inventory"]:
            rows_i = parsed[company]["inventory"]
            # 분류 사전 로드 (위하고에 분류 없으면 보강 + 위하고에 분류 있으면 학습)
            cls_path = ROOT / "data" / "shared" / "item_classification.json"
            cls_map = {}
            cls_doc = {"_generated": NOW, "_note": "", "_count": 0, "items": []}
            if cls_path.exists():
                try:
                    with open(cls_path, encoding="utf-8") as fp:
                        cls_doc = json.load(fp)
                    for it in cls_doc.get("items", []):
                        if it.get("품목코드"):
                            cls_map[it["품목코드"]] = it
                except Exception:
                    pass
            changed = 0; learned = 0
            for r in rows_i:
                code = (r.get("품목코드") or "").strip()
                if not code: continue
                wh_big = (r.get("대분류") or "").strip()
                wh_mid = (r.get("중분류") or "").strip()
                wh_sub = (r.get("소분류") or "").strip()
                cls_entry = cls_map.get(code)
                if wh_big:  # 위하고가 진실 — 사전에 학습
                    if not cls_entry or cls_entry.get("대분류") != wh_big or cls_entry.get("중분류") != wh_mid or cls_entry.get("소분류") != wh_sub:
                        cls_map[code] = {
                            "_company": company, "품목계정": r.get("품목계정",""),
                            "품목코드": code, "품목명": r.get("품목명",""), "규격": r.get("규격",""),
                            "대분류": wh_big, "중분류": wh_mid, "소분류": wh_sub
                        }
                        learned += 1
                elif cls_entry:  # 위하고 빈값 → 사전으로 보강
                    r["대분류"] = cls_entry.get("대분류","")
                    r["중분류"] = cls_entry.get("중분류","")
                    r["소분류"] = cls_entry.get("소분류","")
                    changed += 1
            if changed or learned:
                cls_doc["items"] = list(cls_map.values())
                cls_doc["_count"] = len(cls_doc["items"])
                cls_doc["_generated"] = NOW
                with open(cls_path, "w", encoding="utf-8") as fp:
                    json.dump(cls_doc, fp, ensure_ascii=False, indent=2)
                print(f"    분류 머지: 보강 {changed}건 / 학습 {learned}건 (item_classification.json 갱신)")
            wrapped = {
                "_company": company, "_count": len(rows_i),
                "_generated": NOW, "rows": rows_i,
            }
            outp = DATA_PC / f"inventory_{cl}.json"
            with open(outp, "w", encoding="utf-8") as fp:
                json.dump(wrapped, fp, ensure_ascii=False, separators=(",", ":"))
            print(f"    {outp.name} ({len(rows_i)}건)")
        # 미수: {clients:[...]} 래핑 (로더가 d.clients로 접근)
        if parsed[company]["receivable"]:
            clients_r = parsed[company]["receivable"]
            active = [c for c in clients_r if abs(c.get("채권_잔액", c.get("미수금", 0))) > 0]
            total_chae = sum(c.get("채권_잔액", c.get("미수금", 0)) for c in clients_r)
            total_chamu = sum(c.get("채무_잔액", 0) for c in clients_r)
            wrapped = {
                "_company": company, "_source": "위하고 채권채무집계",
                "_count": len(clients_r), "_active_count": len(active),
                "_total_chae_balance": total_chae,
                "_total_balance": total_chae - total_chamu,
                "_generated": NOW, "clients": clients_r,
            }
            outp = DATA_PC / f"receivable_{cl}.json"
            with open(outp, "w", encoding="utf-8") as fp:
                json.dump(wrapped, fp, ensure_ascii=False, separators=(",", ":"))
            print(f"    {outp.name} ({len(clients_r)}거래처, 활성 {len(active)})")

    print("\n[저장] 양사 합산 (root sales.json) — 옛 dashboard.html용")
    # 양사 모두 합산 (이전엔 SMT만 사용한 버그 — 2026-05-21 수정)
    legacy_sales = []
    for company in ("SMT", "GW"):
        for r in parsed[company]["sales"]:
            legacy_sales.append({
                "날짜": r["일자"], "_month": r["_month"],
                "품목계정": r["품목계정"], "품목코드": r["품목코드"],
                "품목명": r["품목명"], "규격": r["규격"],
                "수량": r["수량"], "단가": r["단가"],
                "공급가액": r["공급가액"], "부가세": r["부가세"],
                "합계금액": r["합계금액"], "금액": r["공급가액"],
                "거래처명": r["거래처명"], "유형": r["유형"],
                "상태": r["유형"], "담당자": r["담당자"], "창고": r["창고"],
                "_company": company,
            })
    if legacy_sales:
        existing_path = DIST_DIR / "sales.json"
        existing = []
        if existing_path.exists():
            try:
                with open(existing_path, encoding="utf-8") as fp:
                    loaded = json.load(fp)
                if isinstance(loaded, dict) and "rows" in loaded:
                    existing = loaded["rows"]
                elif isinstance(loaded, list):
                    existing = loaded
                existing = [x for x in existing if isinstance(x, dict)]
            except Exception:
                existing = []
        # 같은 (일자+회사) 조합만 교체 (다른 회사의 같은 일자 데이터 보존)
        new_keys = {(x["날짜"], x.get("_company", "")) for x in legacy_sales}
        merged = [x for x in existing if (x.get("날짜"), x.get("_company", "")) not in new_keys] + legacy_sales
        merged.sort(key=lambda x: (x.get("날짜", ""), x.get("_company", "")))
        with open(existing_path, "w", encoding="utf-8") as fp:
            json.dump(merged, fp, ensure_ascii=False, separators=(",", ":"))
        print(f"    sales.json (총 {len(merged)}건, 신규 {len(legacy_sales)}건 양사 합산)")

    if parsed["SMT"]["inventory"] or parsed["GW"]["inventory"]:
        inv_map = {}
        for company in ("SMT", "GW"):
            for r in parsed[company]["inventory"]:
                code = r["품목코드"]
                if not code:
                    continue
                if code not in inv_map:
                    inv_map[code] = {
                        "품목계정": r["품목계정"], "품목코드": code,
                        "품목명": r["품목명"], "규격": r["규격"],
                        "기초수량": r["기초수량"], "입고합계": r["입고"],
                        "출고합계": r["출고"], "기말재고": r["기말재고"],
                    }
                else:
                    inv_map[code]["기초수량"] += r["기초수량"]
                    inv_map[code]["입고합계"] += r["입고"]
                    inv_map[code]["출고합계"] += r["출고"]
                    inv_map[code]["기말재고"] += r["기말재고"]
        inv_list = list(inv_map.values())
        with open(DIST_DIR / "inventory.json", "w", encoding="utf-8") as fp:
            json.dump(inv_list, fp, ensure_ascii=False, separators=(",", ":"))
        print(f"    inventory.json ({len(inv_list)}건, 양사 합산)")

    if parsed["SMT"]["receivable"] or parsed["GW"]["receivable"]:
        # 1) 거래처별 양사 합산 (옛 dashboard용)
        rec_map = {}
        for company in ("SMT", "GW"):
            for r in parsed[company]["receivable"]:
                name = r["거래처명"]
                if not name:
                    continue
                rec_map[name] = rec_map.get(name, 0) + r["미수금"]
        rec_list = [{"거래처명": k, "미수금": v} for k, v in rec_map.items() if abs(v) > 0]
        with open(DIST_DIR / "receivable.json", "w", encoding="utf-8") as fp:
            json.dump(rec_list, fp, ensure_ascii=False, separators=(",", ":"))
        print(f"    receivable.json ({len(rec_list)}거래처 양사 합산)")

        # 2) 일자별 history 누적 (변동 추적용)
        #    (per_company receivable_{}.json은 위 [저장] 양사 분리 블록에서 wrap 포맷으로 저장됨)
        history_dir = PROJECT_DIR / "data" / "receivable_history"
        history_dir.mkdir(parents=True, exist_ok=True)
        for company in ("SMT", "GW"):
            if parsed[company]["receivable"]:
                hpath = history_dir / f"{company}_{TODAY}.json"
                with open(hpath, "w", encoding="utf-8") as fp:
                    json.dump(parsed[company]["receivable"], fp, ensure_ascii=False, separators=(",", ":"))
                print(f"    history/{hpath.name} ({len(parsed[company]['receivable'])}거래처)")

    # 미수채권상세현황 저장 (거래처별 + 일자별 history)
    for company in ("SMT", "GW"):
        rd = parsed[company].get("receivable_detail", [])
        if rd:
            out_dir = PROJECT_DIR / "data" / "receivable_analysis"
            out_dir.mkdir(parents=True, exist_ok=True)
            outp = out_dir / f"detail_{company.lower()}_{TODAY}.json"
            with open(outp, "w", encoding="utf-8") as fp:
                json.dump(rd, fp, ensure_ascii=False, separators=(",", ":"))
            print(f"    receivable_analysis/{outp.name} ({len(rd)}거래처)")

    # 분개장 저장 (거래처별 적요·계정 패턴 학습)
    for company in ("SMT", "GW"):
        jr = parsed[company].get("journal", [])
        if jr:
            out_dir = PROJECT_DIR / "data" / "journal_patterns"
            out_dir.mkdir(parents=True, exist_ok=True)
            outp = out_dir / f"journal_{company.lower()}_{TODAY}.json"
            with open(outp, "w", encoding="utf-8") as fp:
                json.dump(jr, fp, ensure_ascii=False, separators=(",", ":"))
            print(f"    journal_patterns/{outp.name} ({len(jr)}분개)")

    print(f"\n[정리] 처리된 파일 → _처리완료/{TODAY}/")
    done_today = DONE_DIR / TODAY
    done_today.mkdir(exist_ok=True)
    for f in processed_files:
        try:
            shutil.move(str(f), str(done_today / f.name))
        except Exception as e:
            print(f"    이동 실패 {f.name}: {e}")

    print("\n[Git] 변경사항 commit + push 중...")
    try:
        subprocess.run(["git", "add", "."], cwd=str(PROJECT_DIR), check=True)
        msg = f"데이터 갱신 {TODAY}"
        r = subprocess.run(["git", "commit", "-m", msg], cwd=str(PROJECT_DIR), capture_output=True, text=True)
        if r.returncode != 0 and "nothing to commit" in (r.stdout + r.stderr):
            print("    (변경사항 없음)")
        else:
            print(f"    commit: {msg}")
            subprocess.run(["git", "push"], cwd=str(PROJECT_DIR), check=True)
            print("    push 완료 — Vercel이 1~2분 안에 자동 배포")
    except FileNotFoundError:
        print("    [경고] git 명령을 찾을 수 없습니다. Git for Windows 또는 GitHub Desktop을 설치하세요.")
    except subprocess.CalledProcessError as e:
        print(f"    [오류] git 실패. 인증 또는 네트워크 문제일 수 있습니다.")

    print("\n" + "=" * 60)
    print(f"  완료! 사이트에서 1~2분 뒤 새로고침하세요.")
    print("=" * 60)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\n[치명적 오류] {e}")
        import traceback
        traceback.print_exc()
    input("\n엔터를 누르면 종료...")

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
    flat = " ".join(s(c) for row in aoa[:5] for c in row)
    # 채권채무집계: 헤더에 "채권" + "채무" + "거래처" + "잔액"
    if ("채권" in flat and "채무" in flat and "거래처" in flat and "잔액" in flat):
        return "receivable"
    # 옛 미수채권상세현황 (보존)
    if "미수채권" in flat or re.search(r"거래처\s*:\s*\d+\s*\[", flat):
        return "receivable_old"
    if "품목계정" in flat and "기말재고" in flat and "기초수량" in flat:
        return "inventory"
    if "일자" in flat and "품목계정" in flat and "거래처명" in flat and "공급가액" in flat:
        return "sales"
    return None


def parse_sales(aoa, company):
    hi = -1
    for i, row in enumerate(aoa):
        if len(row) >= 2 and s(row[0]) == "일자" and s(row[1]) == "품목계정":
            hi = i
            break
    if hi < 0:
        raise ValueError("판매일보 헤더(일자/품목계정)를 찾지 못함")
    rows = []
    for r in aoa[hi + 1:]:
        if not r or s(r[0]) in ("", "일자", "합계", "소계"):
            continue
        date = s(r[0])
        if not re.match(r"\d{4}[-./]\d{2}[-./]\d{2}", date):
            continue
        date = date[:10].replace(".", "-").replace("/", "-")
        rows.append({
            "일자": date, "_month": date[:7],
            "품목계정": s(r[1]), "품목코드": s(r[2]),
            "품목명": s(r[3]), "규격": s(r[4]),
            "수량": num(r[7] if len(r) > 7 else 0),
            "단가": num(r[8] if len(r) > 8 else 0),
            "공급가액": num(r[9] if len(r) > 9 else 0),
            "부가세": num(r[10] if len(r) > 10 else 0),
            "합계금액": num(r[11] if len(r) > 11 else 0),
            "거래처명": s(r[13] if len(r) > 13 else ""),
            "유형": s(r[14] if len(r) > 14 else "판매출고"),
            "담당자": s(r[24] if len(r) > 24 else ""),
            "창고": s(r[27] if len(r) > 27 else ""),
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
           "receivable": parse_receivable, "receivable_old": parse_receivable_old}
TYPE_NAMES = {"sales": "판매일보", "inventory": "현재고",
              "receivable": "채권채무집계", "receivable_old": "미수채권(옛)"}


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

    parsed = {"SMT": {"sales": [], "inventory": [], "receivable": [], "receivable_old": []},
              "GW":  {"sales": [], "inventory": [], "receivable": [], "receivable_old": []}}
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

    print("\n[저장] 양사 분리 (dist/data/per_company/)")
    DATA_PC.mkdir(parents=True, exist_ok=True)
    year = datetime.now().year
    for company in ("SMT", "GW"):
        cl = company.lower()
        if parsed[company]["sales"]:
            outp = DATA_PC / f"sales_{cl}_{year}.json"
            with open(outp, "w", encoding="utf-8") as fp:
                json.dump(parsed[company]["sales"], fp, ensure_ascii=False, separators=(",", ":"))
            print(f"    {outp.name} ({len(parsed[company]['sales'])}건)")
        if parsed[company]["inventory"]:
            outp = DATA_PC / f"inventory_{cl}.json"
            with open(outp, "w", encoding="utf-8") as fp:
                json.dump(parsed[company]["inventory"], fp, ensure_ascii=False, separators=(",", ":"))
            print(f"    {outp.name} ({len(parsed[company]['inventory'])}건)")
        if parsed[company]["receivable"]:
            outp = DATA_PC / f"receivable_{cl}.json"
            with open(outp, "w", encoding="utf-8") as fp:
                json.dump(parsed[company]["receivable"], fp, ensure_ascii=False, separators=(",", ":"))
            print(f"    {outp.name} ({len(parsed[company]['receivable'])}건)")

    print("\n[저장] 양사 합산 (dist/) — 옛 dashboard.html용")
    src = parsed["SMT"]["sales"] if parsed["SMT"]["sales"] else parsed["GW"]["sales"]
    legacy_sales = []
    for r in src:
        legacy_sales.append({
            "날짜": r["일자"], "_month": r["_month"],
            "품목계정": r["품목계정"], "품목코드": r["품목코드"],
            "품목명": r["품목명"], "규격": r["규격"],
            "수량": r["수량"], "단가": r["단가"],
            "공급가액": r["공급가액"], "부가세": r["부가세"],
            "합계금액": r["합계금액"], "금액": r["공급가액"],
            "거래처명": r["거래처명"], "유형": r["유형"],
            "상태": r["유형"], "담당자": r["담당자"], "창고": r["창고"],
        })
    if legacy_sales:
        existing_path = DIST_DIR / "sales.json"
        existing = []
        if existing_path.exists():
            try:
                with open(existing_path, encoding="utf-8") as fp:
                    existing = json.load(fp)
            except Exception:
                existing = []
        new_dates = {x["날짜"] for x in legacy_sales}
        merged = [x for x in existing if x.get("날짜") not in new_dates] + legacy_sales
        merged.sort(key=lambda x: x.get("날짜", ""))
        with open(existing_path, "w", encoding="utf-8") as fp:
            json.dump(merged, fp, ensure_ascii=False, separators=(",", ":"))
        print(f"    sales.json (총 {len(merged)}건, 신규 {len(legacy_sales)}건 병합)")

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

        # 2) 양사 분리 (현재 시점)
        for company in ("SMT", "GW"):
            if parsed[company]["receivable"]:
                outp = DATA_PC / f"receivable_{company.lower()}.json"
                with open(outp, "w", encoding="utf-8") as fp:
                    json.dump(parsed[company]["receivable"], fp, ensure_ascii=False, separators=(",", ":"))
                print(f"    {outp.name} ({len(parsed[company]['receivable'])}거래처)")

        # 3) 일자별 history 누적 (변동 추적용)
        history_dir = PROJECT_DIR / "data" / "receivable_history"
        history_dir.mkdir(parents=True, exist_ok=True)
        for company in ("SMT", "GW"):
            if parsed[company]["receivable"]:
                hpath = history_dir / f"{company}_{TODAY}.json"
                with open(hpath, "w", encoding="utf-8") as fp:
                    json.dump(parsed[company]["receivable"], fp, ensure_ascii=False, separators=(",", ":"))
                print(f"    history/{hpath.name} ({len(parsed[company]['receivable'])}거래처)")

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

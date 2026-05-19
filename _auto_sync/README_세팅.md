# 자동 동기화 1회 세팅 가이드

부장님 PC에 처음 한 번만 세팅하면 그 뒤로는 **바탕화면 아이콘 더블클릭만** 하면 됩니다.

---

## 1. 사전 설치 (한 번만)

### A. Python 설치
1. https://www.python.org/downloads/ 접속
2. 노란 **"Download Python 3.x"** 버튼 클릭
3. 다운받은 설치파일 실행
4. **반드시 ✅ "Add Python to PATH" 체크** → Install Now
5. 완료 후 확인: `Win+R` → `cmd` → `python --version` 입력해서 버전 나오면 OK

### B. openpyxl 라이브러리 설치
1. `Win+R` → `cmd` 엔터
2. 다음 명령 입력:
   ```
   pip install openpyxl
   ```
3. "Successfully installed" 나오면 OK

### C. Git 설치 (GitHub Desktop이 가장 쉬움)
1. https://desktop.github.com 에서 GitHub Desktop 다운로드 + 설치
2. 실행해서 부장님 계정(Rickey65)으로 로그인
3. **File → Clone repository** → `Rickey65/smt-portal` 선택 → Clone
4. 어디에 clone 할지 묻는데 **이 작업폴더를 그대로 선택**하시면 됩니다
   (또는 새 경로 예: `C:\smt-portal`)

### D. 바탕화면 아이콘 만들기
1. 작업폴더 → `_auto_sync` 폴더 열기
2. `sync.bat` 파일 우클릭 → **보내기 → 바탕화면(바로 가기 만들기)**
3. 바탕화면에 "sync.bat - 바로 가기" 생성됨
4. 이름을 **"위하고 데이터 동기화"** 같은 걸로 바꿔두시면 좋습니다

---

## 2. 매일 사용법 (3단계)

### 1단계: 위하고에서 엑셀 다운로드
- **서울기연 로그인** → 판매일보(SHDV*) / 현재고총괄(SHSV*) / 미수채권상세(SHBM*) 다운로드
- **건웅 로그인** → 똑같이 3종 다운로드

### 2단계: 다운로드한 엑셀을 정해진 폴더에 옮기기
- 서울기연 엑셀들 → `_auto_sync/위하고다운로드/서울기연/`
- 건웅 엑셀들 → `_auto_sync/위하고다운로드/건웅/`

### 3단계: 바탕화면 "위하고 데이터 동기화" 아이콘 더블클릭
- 검은 창이 뜨고 자동으로:
  - 엑셀 → JSON 변환
  - GitHub에 자동 commit + push
  - Vercel이 1~2분 안에 사이트 자동 배포
- 완료 메시지 보이면 엔터 → 사이트 새로고침

---

## 3. 문제 해결

| 증상 | 해결 |
|---|---|
| "python: command not found" | Python 설치 시 "Add to PATH" 체크 안 한 것 → 재설치 |
| "ModuleNotFoundError: openpyxl" | cmd에서 `pip install openpyxl` |
| "git: command not found" | GitHub Desktop 설치 |
| "양식 감지 실패" | 다른 양식의 엑셀 → 변환에서 제외됨 (정상) |
| push 실패 | GitHub Desktop을 한 번 열어서 인증 통과 후 재시도 |

---

## 4. 주의사항

- `위하고다운로드/_처리완료/YYYY-MM-DD/` 폴더에 처리된 엑셀이 백업됨 (한 달 뒤 수동 삭제)
- 같은 날짜 거래는 새 데이터로 자동 덮어쓰기 → 중복 안 생김
- 위하고 엑셀 파일명은 아무거나 (위하고가 자동으로 붙이는 SHDV*.xlsx 등 그대로 둬도 됨)

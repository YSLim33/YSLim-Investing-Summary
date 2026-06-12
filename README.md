# YSLim Investing Summary

주간 투자 Summary를 **엑셀 파일 하나**로 관리하고, push만 하면 웹사이트가 자동으로 만들어지는 저장소.

## 운영 원칙

- 원본은 `YSLim_Investing Summary.xlsx` 단 하나. 모든 시트(Orientation, ~2024, 2025, 2026, Factset, Yield, Yield_Chart) 포함.
- 이 파일을 수정해서 GitHub에 올리면(push/업로드) 사이트가 자동으로 다시 빌드·배포됨.
- **Orientation의 흰색 글씨(비공개 메모)는 사이트에서 자동 제외됨.**
- 2026 시트 % 셀은 조건부 서식으로 상승 빨강/하락 파랑 자동 적용.

## 최초 설정 (1회)

1. github.com에서 새 저장소 생성 (예: `investing-summary`, Public)
2. 이 폴더 전체 업로드:
   ```bash
   git init && git add -A && git commit -m "initial"
   git remote add origin https://github.com/YSLim33/investing-summary.git
   git branch -M main && git push -u origin main
   ```
3. 저장소 Settings → Pages → Source: **GitHub Actions** 선택
4. 첫 배포 후 `https://yslim33.github.io/investing-summary/` 접속 확인 → 이 링크를 공유

## 매주 월요일 루틴 (회사 노트북)

**방법 A — 브라우저만으로 (가장 간단, 설치 불필요):**
1. (선택) 월요일 아침 자동으로 새 주차 스냅샷 행이 추가되어 있음 → github.com에서 파일 다운로드
2. Excel에서 해석 코멘트·그래프 추가 후 저장
3. github.com 저장소 → 파일 클릭 → 연필(Edit) 옆 ⋯ → 또는 "Add file > Upload files"로 같은 이름으로 업로드 → Commit
4. 1~2분 후 사이트 자동 갱신 완료

**방법 B — GitHub Desktop (설치 가능하면 더 편함):**
1. GitHub Desktop으로 저장소 clone (최초 1회)
2. clone 폴더의 엑셀 파일을 평소처럼 수정·저장
3. GitHub Desktop에서 Commit → Push → 끝

**방법 C — git 명령어:** `git pull` → 엑셀 수정 → `git add -A && git commit -m "W26xx" && git push`

## 자동화 구성

| 워크플로 | 트리거 | 동작 |
|---|---|---|
| `site.yml` | push 또는 수동 실행 | 엑셀 → 사이트 빌드 → GitHub Pages 배포 |
| `add-week.yml` | 매주 월요일 07:15 KST + 수동 | 직전 금요일 종가(stooq)로 새 W26xx 스냅샷 행을 엑셀에 추가 후 커밋. 이미 있으면 건너뜀 |

> add-week 첫 자동 실행 후에는 행이 제대로 들어갔는지 한 번 확인 권장. 러셀2000(^rut) 등 일부 심볼은 소스에서 누락될 수 있으며, 이 경우 해당 칸만 비워짐(직접 입력).

## 로컬에서 사이트 미리보기

```bash
pip install openpyxl
python tools/build_site.py "YSLim_Investing Summary.xlsx" .
# index.html을 브라우저로 열기
```

## 구조

```
YSLim_Investing Summary.xlsx   ← 관리하는 유일한 파일
tools/build_site.py            ← 엑셀 → 사이트 변환 (흰색 글씨 필터 포함)
tools/add_week.py              ← 주간 스냅샷 행 자동 추가
tools/chart.umd.js             ← 차트 라이브러리 (사이트에 내장됨)
.github/workflows/             ← 자동 빌드·배포 + 주간 자동 행 추가
```

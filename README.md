# 트리 기반 수강신청 시간표 제작 프로그램

전공/교직 과목 엑셀(`.xlsx`)을 넣으면, 시간 충돌 없이 우선순위를 반영한 상위 `n`개의 시간표를 자동으로 만들어주는 C 프로그램입니다.

## 1. 이 프로젝트가 하는 일
- 과목 데이터를 엑셀에서 읽습니다.
- 사용자가 원하는 조건(공강 선호, 늦은 시작 선호, 평점 선호)을 가중치로 받습니다.
- 트리(백트래킹) 탐색으로 가능한 시간표 조합을 만듭니다.
- 점수 순으로 정렬해서 상위 `n`개를 출력합니다.

## 2. 왜 필요한가
수강신청에서 직접 과목을 하나씩 넣어보며
- 시간 겹침 확인
- 우선순위 반영
- 대안 시간표 비교
를 하는 과정이 번거롭습니다.

이 프로그램은 그 과정을 자동화합니다.

## 3. 핵심 아이디어(아주 쉽게)
- 트리의 각 단계에서 "이 과목을 넣을지/안 넣을지"를 선택합니다.
- 시간이 겹치면 그 가지는 즉시 버립니다(가지치기).
- 목표 학점을 만족한 시간표만 결과 후보로 저장합니다.
- 후보들 중 점수가 높은 시간표를 Top-N으로 보여줍니다.

## 4. 프로젝트 구조
- `timetable.c` : 프로그램 전체(단일 C 파일)
- `docs/superpowers/specs/...` : 설계 문서
- `docs/superpowers/plans/...` : 구현 계획 문서

## 5. 개발 환경
- Windows (Excel COM 자동 변환 사용)
- C 컴파일러 1개

예시 컴파일러:
- `gcc`
- `zig cc`

## 6. 빌드 방법
### 방법 A) gcc가 있는 경우
```powershell
gcc -std=c17 -O2 -Wall timetable.c -o timetable.exe
```

### 방법 B) zig를 설치한 경우
```powershell
zig cc -std=c17 -O2 -Wall timetable.c -o timetable.exe
```

## 7. 실행 방법
```powershell
.\timetable.exe --xlsx "C:\Users\yangi\Documents\카카오톡 받은 파일\전공수업.xlsx" --target-credits 18 --top 5
```

### 주요 옵션
- `--xlsx <path>` : 입력 엑셀 파일 경로(필수)
- `--target-credits <N>` : 목표 학점 (기본 18)
- `--top <N>` : 출력할 시간표 개수 (기본 5)
- `--w-free <N>` : 공강 선호 가중치
- `--w-late <N>` : 늦은 시작 선호 가중치
- `--w-rating <N>` : 평점 선호 가중치
- `--must CODE1,CODE2,...` : 반드시 포함할 과목 코드
- `--self-test` : 내부 핵심 기능 테스트 실행

## 8. 입력 파일(`.xlsx`) 처리 방식
- 프로그램은 `.xlsx`를 직접 입력받습니다.
- 내부에서 1번 시트를 임시 CSV로 변환 후 파싱합니다.
- 현재 컬럼 기준으로 다음 정보를 사용합니다:
  - 과목코드
  - 과목명(한글)
  - 학점
  - 이수구분
  - 시간 문자열

시간이 `온라인/미정`인 과목은 고정 시간 없음으로 처리합니다.

## 9. 출력 예시(요약)
- `Loaded courses: 26`
- `Rank 1 ~ Rank N` 시간표 출력
- 각 시간표마다
  - 총 학점
  - 최종 점수
  - 과목 목록
  - 요일/시간

## 10. 구현된 필수 ADT 연산
다음 핵심 연산을 `timetable.c`에 구현했습니다.
- `createTimetable`
- `isConflict`
- `addCourse`
- `removeCourse`
- `getTotalCredits`
- `getSchedule`
- `getRating`
- `isLeaf`
- `scoreTimetable`
- `insertResult`
- `createSch`
- `appendSch`
- `addSch`
- `loadCoursesFromXlsx`

## 11. 빠른 점검
```powershell
.\timetable.exe --self-test
```
정상일 경우:
- `[SELF-TEST] all checks passed ...`

## 12. 현재 한계
- Excel COM 기반이므로 Windows + Excel 설치 환경에서 가장 안정적으로 동작합니다.
- 평점 데이터가 엑셀에 없으면 기본값(현재 3.5)으로 처리합니다.
- 과목 시간이 비어 있거나 `온라인/미정`이면 충돌 검사 대상에서 제외됩니다.

## 13. 다음 개선 후보
- 입력 컬럼 매핑을 설정 파일로 분리
- 온라인 과목 정책(완전 자유/시간대 제한) 선택 옵션 추가
- 점수 세부 항목(공강/시작시간/평점) 분해 출력 강화

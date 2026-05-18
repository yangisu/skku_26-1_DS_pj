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
- 명시적 트리 노드(`TreeNode`)를 만들고, 각 단계에서 "이 과목을 넣을지/안 넣을지"를 자식 노드로 분기합니다.
- 시간이 겹치면 그 가지는 즉시 버립니다(가지치기).
- 목표 학점을 만족한 시간표만 결과 후보로 저장합니다.
- 후보들 중 점수가 높은 시간표를 Top-N으로 보여줍니다.

실행 시 `[TREE] explicit nodes created: ...` 로그로 실제 생성된 트리 노드 수를 확인할 수 있습니다.

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
.\timetable.exe --xlsx "C:\Users\yangi\Documents\카카오톡 받은 파일\전공수업.xlsx" --max-courses 6 --top 5
```

### 주요 옵션
- `--xlsx <path>` : 입력 엑셀 파일 경로(필수)
- `--max-courses <N>` : 시간표에 담을 과목 수 (기본 6)
- `--top <N>` : 출력할 시간표 개수 (`0`이면 최대 `MAX_RESULTS`까지 출력)
- `--w-free <N>` : 공강 선호 가중치
- `--w-late <N>` : 늦은 시작 선호 가중치
- `--w-rating <N>` : 평점 선호 가중치
- `--must TOKEN1,TOKEN2,...` : 반드시 포함할 과목(여러 개 가능)
- `--pool TOKEN1,TOKEN2,...` : 탐색 후보 과목 집합(예: 10개 입력)
- `--self-test` : 내부 핵심 기능 테스트 실행

`N` 형식/범위:
- 형식: `N`은 공백 없는 10진 정수(예: `0`, `5`, `18`)
- `--max-courses N`: `1..16`
- `--top N`: `0..512` (`0`은 가능한 결과를 최대치까지 출력)
- `--w-free N`: `0..100`
- `--w-late N`: `0..100`
- `--w-rating N`: `0..100`

`--must` 입력 규칙:
- `COM3026` : `COM3026-01`, `COM3026-02` 같은 분반 중 **하나 이상** 포함
- `COM3026-02` : 해당 분반을 **정확히** 포함
- 여러 과목은 쉼표로 지정: `--must "COM3026,COM2002,COM3007-01"`

필수 정책(중요):
- 이 프로그램에서 "필수"는 **오직 `--must`로 입력한 과목**을 의미합니다.
- 전공/일반/전필 같은 이수구분 텍스트는 **자동 필수 강제에 사용하지 않습니다**.
- 트리 상단(루트 근처)에서 `--must` 관련 과목들을 먼저 분기한 뒤, 나머지 과목으로 확장합니다.
- 결과는 **정확히 `--max-courses`개 과목**인 시간표만 출력합니다.

분반 처리 규칙:
- 같은 과목 그룹(예: `COM3026`)의 분반은 시간표 1개에 **1개만** 들어갑니다.
- 따라서 분반이 3개인 과목을 `--must COM3026`으로 주면, 3개 분반 중 하나를 선택한 가지들로 시간표가 분화됩니다.
- `--must`를 사용하고 `--top`을 따로 주지 않으면, 프로그램은 가능한 결과를 넓게 보여주기 위해 자동으로 최대치(`MAX_RESULTS`)까지 출력합니다.

10개 후보 중 6개 생성 예시:
```powershell
.\timetable.exe --xlsx "C:\Users\yangi\Documents\카카오톡 받은 파일\전공수업.xlsx" --max-courses 6 --top 10 --w-free 0 --w-late 100 --w-rating 0 --pool "COM2002,COM2003,COM2015,COM2020,COM2023,COM3001,COM3006,COM3007,COM3029,COM3035" --must "COM2002,COM2003,COM2015,COM2020"
```
- `--pool`의 10개 안에서만 조합
- `--must` 4개는 고정 포함
- 남은 2자리를 가중치 점수 기준으로 분화/정렬

## 7-1. 코드 흐름(트리가 실제로 어떻게 진행되는지)
1. `main`이 인자를 파싱하고 `loadCoursesFromXlsx`로 과목 배열을 만듭니다.
2. `createSch`에서 과목 순서를 재배치합니다.
- 먼저 `--must`에 매칭되는 과목(분반 포함)을 앞쪽에 배치
- 나머지 과목을 뒤에 배치
3. 루트 노드(`TreeNode`)를 생성합니다.
- `level=0`
- `state=빈 시간표`
- `decision_course=NULL`, `took_course=0`
4. `build_and_score_tree`가 노드를 확장합니다.
- 현재 `level`의 과목에 대해 자식 노드 분기:
- `include_child`: 해당 과목을 시간표에 추가한 상태
- `exclude_child`: 해당 과목을 추가하지 않은 상태
5. 가지치기 조건에 걸리면 그 노드 아래는 더 이상 확장하지 않습니다.
6. 리프/완성 후보에서 점수를 계산해 `ResultSet`에 삽입합니다.
7. 탐색 종료 후 점수 상위 `top_n`개 시간표를 출력합니다.

실행 로그의 `[TREE] explicit nodes created: ...`는 실제 생성된 명시적 노드 수입니다.

## 7-2. 노드에 들어가는 데이터와 자식 분화
`TreeNode` 필드:
- `level`: 현재 몇 번째 과목을 결정 중인지
- `state`: 지금까지 선택 결과가 반영된 `Timetable` 상태
- `decision_course`: 이 노드에서 결정 대상으로 본 과목 포인터
- `took_course`: `1`이면 포함 분기, `0`이면 제외 분기
- `include_child`, `exclude_child`: 다음 분기 노드

자식 노드 개수:
- 기본적으로 최대 2개(포함/제외)
- 단, 포함 분기는 `addCourse` 실패(시간충돌/같은 과목 그룹 중복) 시 생성되지 않음
- 제외 분기는 현재 구현에서 항상 시도

## 7-3. 기본 자료형(ADT) 구조 요약
- `TimeSlot`: 요일/시작분/종료분
- `Course`: 과목 코드, 이름, 카테고리, 학점, 평점, 시간슬롯들
- `Preference`: 과목 수 제한(`max_courses`), top_n, 가중치, `--must`/`--pool` 토큰 목록
- `Timetable`: 현재 담긴 과목 목록, 총학점, 시간 점유표, 점수
- `TreeNode`: 명시적 탐색 노드
- `ResultSet`: 최종 후보 시간표 상위 목록
- `SearchContext`: 탐색 중 공통 참조(과목 배열, 우선순위, 결과 버퍼, suffix 학점 등)

## 7-4. 탐색 기준(확장/중단 판단)
포함 분기 생성 조건:
- `addCourse` 성공 필요
- 실패 사유: 시간 충돌(`isConflict`) 또는 같은 과목 그룹 분반 중복

노드 확장 중단(가지치기) 조건:
- 남은 후보로 `max_courses`를 채울 수 없음 (`can_reach_target_with_suffix`)
- 앞으로 가도 `--must`를 만족시킬 수 없음 (`can_still_satisfy_all_must`)
- 과목 수가 `max_courses`를 초과하면 즉시 중단 (`exceeds_credit_limit`)

결과 저장 조건:
- `isLeaf`로 과목 수 `max_courses` 도달
- `all_must_include_satisfied`로 `--must` 전부 충족
- `scoreTimetable` 점수 계산 후 `insertResult`로 상위 결과 유지

## 7-5. 함수 설명(핵심)
입력/전처리:
- `parse_args`: CLI 인자 파싱 및 범위 검증
- `loadCoursesFromXlsx`: 엑셀을 CSV로 변환 후 `Course` 배열 생성

시간표 ADT:
- `createTimetable`: 빈 시간표 초기화
- `isConflict`: 시간표와 과목 시간 충돌 검사
- `addCourse`: 충돌/분반중복이 없으면 과목 추가
- `removeCourse`: 과목 제거(탐색 복원용)
- `getTotalCredits`, `getSchedule`, `getRating`: 조회 함수

트리 탐색:
- `createSch`: `--must` 우선 정렬 + 루트 생성 + 전체 탐색 시작
- `create_tree_node`: 명시적 노드 생성
- `build_and_score_tree`: 포함/제외 분기 확장 및 재귀 탐색
- `destroy_tree`: 생성 노드 메모리 해제

필수/분반 정책:
- `course_matches_must_token`: `COM3026`/`COM3026-02` 매칭 규칙
- `can_still_satisfy_all_must`: 앞으로 `--must` 충족 가능성 검사
- `all_must_include_satisfied`: 결과 노드의 `--must` 최종 충족 검사

점수/결과:
- `scoreTimetable`: 가중치 기반 점수 계산
- `insertResult`: 점수순 상위 `N`개 유지

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

## 14. 실데이터 테스트 기록 (`전공수업.xlsx` + 6개 필수 과목)
테스트 입력 파일:
- `C:\Users\yangi\Documents\카카오톡 받은 파일\전공수업.xlsx`

사진 기준 필수 6과목:
- 기본프로그래밍
- 컴퓨터교육개론
- 피지컬컴퓨팅
- 머신러닝
- 자연어처리
- 교육용멀티미디어

코드 매핑:
- 기본프로그래밍 -> `COM2002` (`COM2002-01`, `COM2002-02` 분반 존재)
- 컴퓨터교육개론 -> `COM2003-01`
- 피지컬컴퓨팅 -> `COM2015-01`
- 머신러닝 -> `COM2020-01`
- 자연어처리 -> `COM2023-01`
- 교육용멀티미디어 -> `COM3001-01`

### 14-1. 실행 과정
1. 빌드:
```powershell
zig cc -std=c17 -O2 -Wall timetable.c -o timetable.exe
```
2. 분반 자유(기본프로그래밍은 아무 분반이나 허용):
```powershell
.\timetable.exe --xlsx "C:\Users\yangi\Documents\카카오톡 받은 파일\전공수업.xlsx" --max-courses 6 --top 10 --must "COM2002,COM2003,COM2015,COM2020,COM2023,COM3001"
```
3. 분반 고정(기본프로그래밍 01분반 강제):
```powershell
.\timetable.exe --xlsx "C:\Users\yangi\Documents\카카오톡 받은 파일\전공수업.xlsx" --max-courses 6 --top 3 --must "COM2002-01,COM2003,COM2015,COM2020,COM2023,COM3001"
```

### 14-2. 요청 조건 실행 (가중치: 늦은 시작만 높게)
실행 명령:
```powershell
.\timetable.exe --xlsx "C:\Users\yangi\Documents\카카오톡 받은 파일\전공수업.xlsx" --max-courses 6 --top 10 --w-free 0 --w-late 100 --w-rating 0 --must "COM2002,COM2003,COM2015,COM2020,COM2023,COM3001"
```

실행 결과 전체(원문):
```text
Loaded courses: 26
[TREE] explicit nodes created: 26

========== Rank 1 ==========
Total Credits: 18
Score: 1200.00
Courses (6):
  - COM2002-02 | 기본프로그래밍 | 일반수업 | 3cr
      time: Fri 10:30-11:45
  - COM2003-01 | 컴퓨터교육개론 | 국제어수업 | 3cr
      time: Tue 09:00-10:15
      time: Tue 10:30-11:45
  - COM2015-01 | 피지컬컴퓨팅 | 국제어수업 | 3cr
      time: Mon 09:00-10:15
      time: Mon 10:30-11:45
  - COM2020-01 | 머신러닝 | 일반수업 | 3cr
      time: Wed 18:00-19:15
      time: Wed 19:30-20:45
  - COM2023-01 | 자연어처리 | 국제어수업 | 3cr
      time: Thu 13:30-14:45
  - COM3001-01 | 교육용멀티미디어 | 일반수업 | 3cr
      time: Fri 12:00-13:15
      time: Fri 13:30-14:45

========== Rank 2 ==========
Total Credits: 18
Score: 1140.00
Courses (6):
  - COM2002-01 | 기본프로그래밍 | 국제어수업 | 3cr
      time: Wed 13:30-14:45
  - COM2003-01 | 컴퓨터교육개론 | 국제어수업 | 3cr
      time: Tue 09:00-10:15
      time: Tue 10:30-11:45
  - COM2015-01 | 피지컬컴퓨팅 | 국제어수업 | 3cr
      time: Mon 09:00-10:15
      time: Mon 10:30-11:45
  - COM2020-01 | 머신러닝 | 일반수업 | 3cr
      time: Wed 18:00-19:15
      time: Wed 19:30-20:45
  - COM2023-01 | 자연어처리 | 국제어수업 | 3cr
      time: Thu 13:30-14:45
  - COM3001-01 | 교육용멀티미디어 | 일반수업 | 3cr
      time: Fri 12:00-13:15
      time: Fri 13:30-14:45
```

### 14-3. 현재 실행 결과 출력 형태
프로그램 출력은 아래 순서로 고정됩니다.
1. `Loaded courses: <개수>`
2. `[TREE] explicit nodes created: <노드 수>`
3. `Rank 1..N` 블록 반복
4. 각 Rank 블록 내부:
- `Total Credits`
- `Score`
- `Courses (k)` (시간표에 담긴 과목 수)
- 과목별 라인: `과목코드 | 과목명 | 수업구분 | 학점`
- 시간 라인: `time: 요일 시작-종료` (복수 시간대면 여러 줄)

## 15. 요청 시나리오 테스트 (필수 3 + 고려 풀 4, 공강 최우선)
요청 입력:
- 필수: 기본프로그래밍, 컴퓨터교육개론, 피지컬 컴퓨팅
- 고려 풀: 기계학습, 교육용멀티미디어, 운영체제, 컴퓨터네트워크
- 가중치: 공강(`--w-free`) 최우선, 나머지 0
- 총 과목 수: 6개

실행 명령:
```powershell
.\timetable.exe --xlsx "C:\Users\yangi\Documents\카카오톡 받은 파일\전공수업.xlsx" --max-courses 6 --top 20 --w-free 100 --w-late 0 --w-rating 0 --must "기본프로그래밍,컴퓨터교육개론,피지컬 컴퓨팅" --pool "기계학습,교육용멀티미디어,운영체제,컴퓨터네트워크"
```

확인 포인트:
- 필수 3개는 항상 포함됨
- 나머지 3개는 고려 풀 4개 중에서만 선택됨
- 결과는 점수 내림차순으로 정렬됨
- 루트 근처 분기에서 필수 과목 분반(`기본프로그래밍`의 01/02)이 먼저 갈라짐

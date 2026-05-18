# 트리 기반 수강신청 시간표 생성기 설계서

**작성일:** 2026-05-18
**대상 저장소:** `yangisu/skku_26-1_DS_pj`

## 1) 목표와 범위
- 목표: 전공/교직 과목 데이터(`전공수업.xlsx`)와 사용자 우선순위를 입력받아, 조건을 만족하는 상위 `n`개 시간표를 생성한다.
- 범위: 콘솔 기반 단일 C 파일 프로그램(`timetable.c`)만 제공한다.
- 제외: GUI, 로그인, 실시간 수강신청 연동, 불필요한 통계 기능.

## 2) 필수 요구사항 재정의
- 입력 파일 형식은 `.xlsx` 경로를 받는다.
- 우선순위 조건(예: 공강 선호, 늦은 시작 선호, 과목 평점 선호)을 가중치로 입력받는다.
- 필수 과목/선택 과목 제약을 반영한다.
- 시간 충돌 없는 조합만 생성한다.
- 점수 기준으로 정렬된 상위 `n`개 시간표를 출력한다.
- 프로그램 전체는 C 소스 파일 1개에서 컴파일/실행 가능해야 한다.

## 3) 구현 방식(단일 C 파일 제약 대응)
- 실행 인자 예시:
  - `timetable.exe --xlsx "C:\\...\\전공수업.xlsx" --target-credits 18 --top 5`
- `.xlsx` 처리 전략(MVP):
  - C 프로그램이 `.xlsx` 경로를 입력받는다.
  - 내부에서 임시 CSV로 변환 후 파싱한다(Windows PowerShell + Excel COM 또는 사전 합의된 변환기 호출).
  - 즉, 사용자 입력은 끝까지 `.xlsx`이며, 내부 파이프라인만 단순화한다.
- 단일 파일 구조:
  - `typedef struct`와 함수 구현을 모두 `timetable.c`에 포함.
  - 필요 시 `#ifdef _WIN32` 분기만 사용.

## 4) ADT 재정의(최소 기능만)

### 4-1. `TimeSlot` ADT
```c
typedef struct {
    int day;        // 0=Mon ... 6=Sun
    int start_min;  // 분 단위
    int end_min;
} TimeSlot;
```

### 4-2. `Course` ADT
```c
typedef struct {
    char code[32];
    char name[64];
    char category[32];    // 전필/전선/교직 등
    int credits;
    float rating;
    int is_required;
    int slot_count;
    TimeSlot slots[8];
} Course;
```

### 4-3. `Preference` ADT
```c
typedef struct {
    int target_credits;
    int top_n;
    int prefer_free_days;         // 공강 선호 가중치
    int prefer_late_start;        // 늦은 시작 선호 가중치
    int prefer_high_rating;       // 평점 선호 가중치
    int must_include_count;
    char must_include_codes[32][32];
} Preference;
```

### 4-4. `Timetable` ADT
```c
typedef struct {
    Course *enrolled[16];
    int enrolled_count;
    int total_credits;
    int occupied[7][24 * 2];   // 30분 단위 점유
    float score;
} Timetable;
```

### 4-5. `SearchNode` ADT
```c
typedef struct {
    int level;                 // 현재 탐색 깊이
    Timetable current;
} SearchNode;
```

### 4-6. `ResultSet` ADT
```c
typedef struct {
    Timetable items[32];       // top_n 최대치 제한
    int count;
} ResultSet;
```

## 5) 필수 ADT 연산(10개 이상, 꼭 필요한 것만)
1. `void createTimetable(Timetable *tt);`
2. `int isConflict(const Timetable *tt, const Course *c);`
3. `int addCourse(Timetable *tt, Course *c);`
4. `void removeCourse(Timetable *tt, Course *c);`
5. `int getTotalCredits(const Timetable *tt);`
6. `const TimeSlot* getSchedule(const Course *c, int *count_out);`
7. `float getRating(const Course *c);`
8. `int isLeaf(const SearchNode *node, const Preference *pref);`
9. `float scoreTimetable(const Timetable *tt, const Preference *pref);`
10. `void insertResult(ResultSet *rs, const Timetable *tt, int top_n);`
11. `void createSch(const Course *courses, int course_count, const Preference *pref, ResultSet *out);`
12. `int appendSch(Timetable *tt, Course *course);`  // 수정 모드에서 과목 교체/추가
13. `int addSch(Timetable *tt, Course *course);`     // 사용자 직접 추가
14. `int loadCoursesFromXlsx(const char *xlsx_path, Course *out, int max_courses);`

## 6) 트리 탐색 규칙
- 레벨 정의:
  - `level 0`: 필수 과목 확정
  - `level 1..k`: 카테고리별 과목 선택(전필/전선/교직)
  - 마지막: 목표 학점 도달 여부 확인
- 가지치기(pruning):
  - 시간 충돌 발생 시 즉시 중단
  - 남은 과목으로 목표 학점 달성 불가 시 중단
  - 현재 점수 상한이 결과 하한보다 낮으면 중단

## 7) 출력 형식
- 상위 `n`개 시간표를 점수 내림차순으로 출력
- 각 시간표별:
  - 총학점
  - 선택 과목 목록(과목코드, 분반, 시간)
  - 점수 상세(공강 점수/시작 시간 점수/평점 점수)

## 8) 성공 기준
- 같은 입력에 대해 항상 동일한 정렬 결과를 재현한다.
- 시간 충돌 과목이 결과에 포함되지 않는다.
- 목표 학점 및 필수 과목 조건을 만족하는 결과만 출력한다.
- `top_n` 개수만큼(가능한 범위 내) 결과를 제시한다.

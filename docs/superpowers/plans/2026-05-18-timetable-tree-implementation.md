# Timetable Tree Generator Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a single-file C program that reads `전공수업.xlsx`, applies priority constraints, and outputs top `n` valid timetables.

**Architecture:** A DFS/backtracking tree search over course combinations with strict conflict pruning. `Timetable` ADT maintains occupancy grid and credits. `ResultSet` keeps top-scoring schedules sorted by score.

**Tech Stack:** C17, Windows PowerShell interop for `.xlsx` -> temp CSV conversion, standard C library.

---

### Task 1: Define Single-File Program Skeleton and Core ADTs

**Files:**
- Create: `D:/skku/2-1/DS/termpj/timetable.c`

- [ ] **Step 1: Create core structs and constants in one file**
```c
#define MAX_COURSES 256
#define MAX_SLOTS 8
#define MAX_ENROLLED 16
#define MAX_RESULTS 32

typedef struct { int day, start_min, end_min; } TimeSlot;
typedef struct { char code[32], name[64], category[32]; int credits; float rating; int is_required; int slot_count; TimeSlot slots[MAX_SLOTS]; } Course;
typedef struct { int target_credits, top_n; int prefer_free_days, prefer_late_start, prefer_high_rating; int must_include_count; char must_include_codes[32][32]; } Preference;
typedef struct { Course *enrolled[MAX_ENROLLED]; int enrolled_count, total_credits; int occupied[7][48]; float score; } Timetable;
typedef struct { int level; Timetable current; } SearchNode;
typedef struct { Timetable items[MAX_RESULTS]; int count; } ResultSet;
```

- [ ] **Step 2: Declare mandatory ADT function signatures (14개)**
```c
void createTimetable(Timetable *tt);
int isConflict(const Timetable *tt, const Course *c);
int addCourse(Timetable *tt, Course *c);
void removeCourse(Timetable *tt, Course *c);
int getTotalCredits(const Timetable *tt);
const TimeSlot* getSchedule(const Course *c, int *count_out);
float getRating(const Course *c);
int isLeaf(const SearchNode *node, const Preference *pref);
float scoreTimetable(const Timetable *tt, const Preference *pref);
void insertResult(ResultSet *rs, const Timetable *tt, int top_n);
void createSch(const Course *courses, int course_count, const Preference *pref, ResultSet *out);
int appendSch(Timetable *tt, Course *course);
int addSch(Timetable *tt, Course *course);
int loadCoursesFromXlsx(const char *xlsx_path, Course *out, int max_courses);
```

- [ ] **Step 3: Add CLI usage contract**
```c
// timetable.exe --xlsx <path> --target-credits <int> --top <int>
```

- [ ] **Step 4: Compile skeleton**
Run: `gcc -std=c17 -O2 -Wall timetable.c -o timetable.exe`
Expected: build success without undefined symbol errors.

### Task 2: Implement Timetable ADT and Conflict/Credit Logic

**Files:**
- Modify: `D:/skku/2-1/DS/termpj/timetable.c`

- [ ] **Step 1: Implement createTimetable / getTotalCredits**
```c
void createTimetable(Timetable *tt) {
    memset(tt, 0, sizeof(*tt));
}

int getTotalCredits(const Timetable *tt) {
    return tt->total_credits;
}
```

- [ ] **Step 2: Implement isConflict (30분 슬롯 점유 기반)**
```c
int isConflict(const Timetable *tt, const Course *c) {
    for (int i = 0; i < c->slot_count; ++i) {
        int day = c->slots[i].day;
        int s = c->slots[i].start_min / 30;
        int e = c->slots[i].end_min / 30;
        for (int t = s; t < e; ++t) {
            if (tt->occupied[day][t]) return 1;
        }
    }
    return 0;
}
```

- [ ] **Step 3: Implement addCourse / removeCourse / addSch / appendSch**
```c
int addCourse(Timetable *tt, Course *c) {
    if (tt->enrolled_count >= MAX_ENROLLED) return 0;
    if (isConflict(tt, c)) return 0;
    tt->enrolled[tt->enrolled_count++] = c;
    tt->total_credits += c->credits;
    for (int i = 0; i < c->slot_count; ++i) {
        int day = c->slots[i].day;
        for (int t = c->slots[i].start_min / 30; t < c->slots[i].end_min / 30; ++t) tt->occupied[day][t] = 1;
    }
    return 1;
}
```

- [ ] **Step 4: Compile and smoke-check conflict behavior**
Run: `gcc -std=c17 -O2 -Wall timetable.c -o timetable.exe`
Expected: build success and manual sample with overlapping courses returns reject.

### Task 3: Implement Scoring and Top-N Result Management

**Files:**
- Modify: `D:/skku/2-1/DS/termpj/timetable.c`

- [ ] **Step 1: Implement getSchedule / getRating helper**
```c
const TimeSlot* getSchedule(const Course *c, int *count_out) { *count_out = c->slot_count; return c->slots; }
float getRating(const Course *c) { return c->rating; }
```

- [ ] **Step 2: Implement scoreTimetable with only 3 priorities**
```c
float scoreTimetable(const Timetable *tt, const Preference *pref) {
    float rating_sum = 0.0f;
    for (int i = 0; i < tt->enrolled_count; ++i) rating_sum += tt->enrolled[i]->rating;
    float avg_rating = tt->enrolled_count ? rating_sum / tt->enrolled_count : 0.0f;
    int free_days = 0; /* weekday(월~금) 기준 공강 수 계산 */
    float late_start_score = 0.0f; /* 요일별 첫 수업 시작시각 기반 */
    return pref->prefer_high_rating * avg_rating
         + pref->prefer_free_days * free_days
         + pref->prefer_late_start * late_start_score;
}
```

- [ ] **Step 3: Implement insertResult (정렬 삽입 + 상한 컷)**
```c
void insertResult(ResultSet *rs, const Timetable *tt, int top_n) {
    // score 기준 내림차순 삽입, count > top_n 이면 마지막 제거
}
```

- [ ] **Step 4: Compile and validate deterministic ordering**
Run: `gcc -std=c17 -O2 -Wall timetable.c -o timetable.exe`
Expected: same input/weights on repeated runs produce same ordering.

### Task 4: Implement DFS/Backtracking Tree Search

**Files:**
- Modify: `D:/skku/2-1/DS/termpj/timetable.c`

- [ ] **Step 1: Implement isLeaf completion logic**
```c
int isLeaf(const SearchNode *node, const Preference *pref) {
    return node->current.total_credits >= pref->target_credits;
}
```

- [ ] **Step 2: Implement createSch DFS**
```c
void createSch(const Course *courses, int course_count, const Preference *pref, ResultSet *out) {
    // 재귀/스택 방식 DFS: 과목 포함/미포함 분기, 충돌 시 prune
}
```

- [ ] **Step 3: Add required-course and remaining-credit pruning**
```c
// 남은 depth에서 target_credits 달성 불가하면 return
// 필수 과목 누락 경로는 결과 저장 전 폐기
```

- [ ] **Step 4: Run end-to-end on sample input**
Run: `timetable.exe --xlsx "C:\Users\yangi\Documents\카카오톡 받은 파일\전공수업.xlsx" --target-credits 18 --top 5`
Expected: 조건을 만족하는 상위 5개(또는 가능한 최대 개수) 출력.

### Task 5: Implement XLSX Input Pipeline and CLI

**Files:**
- Modify: `D:/skku/2-1/DS/termpj/timetable.c`

- [ ] **Step 1: Implement loadCoursesFromXlsx interface**
```c
int loadCoursesFromXlsx(const char *xlsx_path, Course *out, int max_courses) {
    // 1) xlsx 존재 확인
    // 2) 임시 csv 경로 생성
    // 3) PowerShell 변환 호출
    // 4) csv 파싱 후 Course 배열 채움
}
```

- [ ] **Step 2: Fix expected column schema and parser**
```text
code,name,category,credits,rating,is_required,day,start,end
```

- [ ] **Step 3: Implement robust argument parsing + validation**
```c
// --xlsx, --target-credits, --top 미입력 시 usage 출력 후 종료
```

- [ ] **Step 4: Build and run with real file path**
Run: `gcc -std=c17 -O2 -Wall timetable.c -o timetable.exe`
Expected: `.xlsx` 경로를 받아 정상 파싱 또는 명확한 오류 메시지 출력.

### Task 6: Output Formatting and Final Verification

**Files:**
- Modify: `D:/skku/2-1/DS/termpj/timetable.c`

- [ ] **Step 1: Print top-n timetables with score breakdown**
```c
// Rank, Total Credits, Score, 과목 목록, 요일/시간 출력
```

- [ ] **Step 2: Add minimal self-check mode**
```c
// --self-test: 충돌 판정/학점 누적/정렬 삽입 핵심 단위 검증
```

- [ ] **Step 3: Run verification commands**
Run: `timetable.exe --self-test`
Expected: all checks passed.

Run: `timetable.exe --xlsx "C:\Users\yangi\Documents\카카오톡 받은 파일\전공수업.xlsx" --target-credits 18 --top 5`
Expected: 충돌 없는 결과 n개 출력.

- [ ] **Step 4: Commit baseline implementation**
```bash
git add timetable.c docs/superpowers/specs/2026-05-18-timetable-tree-design.md docs/superpowers/plans/2026-05-18-timetable-tree-implementation.md
git commit -m "docs: define ADT and implementation plan for timetable tree generator"
```

## Spec Coverage Check
- 단일 C 파일 제약: Task 1 전체와 Task 5/6에서 보장.
- `.xlsx` 입력: Task 5에서 명시적 처리.
- 최소 기능만 구현: 충돌/학점/우선순위/Top-N 생성만 포함.
- 결과 n개 출력: Task 4 + Task 6에서 보장.
- ADT 10개 이상: Task 1 함수 시그니처 14개로 충족.

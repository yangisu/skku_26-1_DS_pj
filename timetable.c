#include <ctype.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>

#ifdef _WIN32
#include <windows.h>
#endif

#define MAX_COURSES 512
#define MAX_SLOTS 8
#define MAX_ENROLLED 16
#define MAX_RESULTS 32
#define MAX_FIELD 256
#define MAX_LINE 4096
#define MAX_MUST_INCLUDE 32
#define DAY_COUNT 7
#define SLOT_COUNT_PER_DAY 48
#define MAX_CREDIT_OVERFLOW 3

typedef struct {
    int day;        /* 0=Mon ... 6=Sun */
    int start_min;  /* minutes from 00:00 */
    int end_min;
} TimeSlot;

typedef struct {
    char code[32];
    char name[64];
    char category[32];
    int credits;
    float rating;
    int is_required;
    int slot_count;
    TimeSlot slots[MAX_SLOTS];
} Course;

typedef struct {
    int target_credits;
    int top_n;
    int prefer_free_days;
    int prefer_late_start;
    int prefer_high_rating;
    int must_include_count;
    char must_include_codes[MAX_MUST_INCLUDE][32];
} Preference;

typedef struct {
    Course *enrolled[MAX_ENROLLED];
    int enrolled_count;
    int total_credits;
    int occupied[DAY_COUNT][SLOT_COUNT_PER_DAY];
    float score;
} Timetable;

typedef struct {
    int level;
    Timetable current;
} SearchNode;

typedef struct {
    Timetable items[MAX_RESULTS];
    int count;
} ResultSet;

typedef struct {
    const Course *courses;
    int course_count;
    const Preference *pref;
    ResultSet *out;
    int suffix_credits[MAX_COURSES + 1];
} SearchContext;

static int g_self_test = 0;

static void print_usage(void);
static int parse_args(int argc, char **argv, char *xlsx_path, size_t xlsx_path_cap, Preference *pref);
static void trim(char *s);
static int starts_with_ignore_case(const char *s, const char *prefix);
static int parse_csv_line(const char *line, char fields[][MAX_FIELD], int max_fields);
static int parse_day_token(const char *token);
static int parse_time_hhmm(const char *s, int *out_minutes);
static int parse_time_range(const char *token, TimeSlot *slot_out);
static void normalize_text_field(char *dst, size_t dst_cap, const char *src);
static int contains_ignore_case(const char *s, const char *needle);
static int is_required_category(const char *category);
static int course_is_must_include(const Course *course, const Preference *pref);
static int all_must_include_satisfied(const Timetable *tt, const Preference *pref);
static void print_course(const Course *c);
static void print_timetable(const Timetable *tt, int rank);
static void delete_file_if_exists(const char *path);
static int run_self_test(void);
static int timetable_equal_key(const Timetable *a, const Timetable *b);

void createTimetable(Timetable *tt);
int isConflict(const Timetable *tt, const Course *c);
int addCourse(Timetable *tt, Course *c);
void removeCourse(Timetable *tt, Course *c);
int getTotalCredits(const Timetable *tt);
const TimeSlot *getSchedule(const Course *c, int *count_out);
float getRating(const Course *c);
int isLeaf(const SearchNode *node, const Preference *pref);
float scoreTimetable(const Timetable *tt, const Preference *pref);
void insertResult(ResultSet *rs, const Timetable *tt, int top_n);
void createSch(const Course *courses, int course_count, const Preference *pref, ResultSet *out);
int appendSch(Timetable *tt, Course *course);
int addSch(Timetable *tt, Course *course);
int loadCoursesFromXlsx(const char *xlsx_path, Course *out, int max_courses);

static void create_temp_csv_path(char *out, size_t out_cap) {
#ifdef _WIN32
    char temp_path[MAX_PATH];
    char temp_file[MAX_PATH];
    DWORD n = GetTempPathA((DWORD)sizeof(temp_path), temp_path);
    if (n == 0 || n >= sizeof(temp_path)) {
        snprintf(out, out_cap, "timetable_temp_%lu.csv", (unsigned long)GetTickCount());
        return;
    }
    if (GetTempFileNameA(temp_path, "tt", 0, temp_file) == 0) {
        snprintf(out, out_cap, "%stimetable_temp_%lu.csv", temp_path, (unsigned long)GetTickCount());
        return;
    }
    delete_file_if_exists(temp_file);
    snprintf(out, out_cap, "%s.csv", temp_file);
#else
    snprintf(out, out_cap, "./timetable_temp.csv");
#endif
}

static void escape_ps_single_quotes(const char *src, char *dst, size_t cap) {
    size_t j = 0;
    size_t i;
    for (i = 0; src[i] != '\0' && j + 1 < cap; ++i) {
        if (src[i] == '\'') {
            if (j + 2 >= cap) break;
            dst[j++] = '\'';
            dst[j++] = '\'';
        } else {
            dst[j++] = src[i];
        }
    }
    dst[j] = '\0';
}

static int convert_xlsx_to_csv_utf8(const char *xlsx_path, const char *csv_path) {
#ifdef _WIN32
    char xesc[1024];
    char cesc[1024];
    char command[8192];
    escape_ps_single_quotes(xlsx_path, xesc, sizeof(xesc));
    escape_ps_single_quotes(csv_path, cesc, sizeof(cesc));

    snprintf(
        command,
        sizeof(command),
        "powershell -NoProfile -ExecutionPolicy Bypass -Command \"$ErrorActionPreference='Stop';"
        "$xlsx='%s';$csv='%s';"
        "$excel=New-Object -ComObject Excel.Application;"
        "$excel.Visible=$false;$excel.DisplayAlerts=$false;"
        "$wb=$excel.Workbooks.Open($xlsx);"
        "$ws=$wb.Worksheets.Item(1);"
        "$ws.SaveAs($csv,62);"
        "$wb.Close($false);$excel.Quit();"
        "[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ws)|Out-Null;"
        "[System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb)|Out-Null;"
        "[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)|Out-Null;\"",
        xesc,
        cesc
    );

    if (system(command) != 0) {
        /* Fallback to legacy CSV format if UTF-8 export is unavailable. */
        snprintf(
            command,
            sizeof(command),
            "powershell -NoProfile -ExecutionPolicy Bypass -Command \"$ErrorActionPreference='Stop';"
            "$xlsx='%s';$csv='%s';"
            "$excel=New-Object -ComObject Excel.Application;"
            "$excel.Visible=$false;$excel.DisplayAlerts=$false;"
            "$wb=$excel.Workbooks.Open($xlsx);"
            "$ws=$wb.Worksheets.Item(1);"
            "$ws.SaveAs($csv,6);"
            "$wb.Close($false);$excel.Quit();"
            "[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ws)|Out-Null;"
            "[System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb)|Out-Null;"
            "[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)|Out-Null;\"",
            xesc,
            cesc
        );
        if (system(command) != 0) {
            return 0;
        }
    }
    return 1;
#else
    (void)xlsx_path;
    (void)csv_path;
    return 0;
#endif
}

static Course *find_or_create_course(
    Course *courses,
    int *course_count,
    int max_courses,
    const char *code,
    const char *name,
    const char *category,
    int credits,
    float rating,
    int is_required
) {
    int i;
    for (i = 0; i < *course_count; ++i) {
        if (strcmp(courses[i].code, code) == 0) {
            return &courses[i];
        }
    }
    if (*course_count >= max_courses) {
        return NULL;
    }
    memset(&courses[*course_count], 0, sizeof(Course));
    snprintf(courses[*course_count].code, sizeof(courses[*course_count].code), "%s", code);
    snprintf(courses[*course_count].name, sizeof(courses[*course_count].name), "%s", name);
    snprintf(courses[*course_count].category, sizeof(courses[*course_count].category), "%s", category);
    courses[*course_count].credits = credits;
    courses[*course_count].rating = rating;
    courses[*course_count].is_required = is_required;
    (*course_count)++;
    return &courses[*course_count - 1];
}

int loadCoursesFromXlsx(const char *xlsx_path, Course *out, int max_courses) {
    char csv_path[1024];
    FILE *fp;
    char line[MAX_LINE];
    int line_no = 0;
    int course_count = 0;

    create_temp_csv_path(csv_path, sizeof(csv_path));

    if (!convert_xlsx_to_csv_utf8(xlsx_path, csv_path)) {
        fprintf(stderr, "[ERROR] Failed to convert xlsx to csv: %s\n", xlsx_path);
        return -1;
    }

    fp = fopen(csv_path, "rb");
    if (fp == NULL) {
        fprintf(stderr, "[ERROR] Failed to open converted csv: %s\n", csv_path);
        delete_file_if_exists(csv_path);
        return -1;
    }

    while (fgets(line, sizeof(line), fp) != NULL) {
        char fields[16][MAX_FIELD];
        int field_count;
        char code[32];
        char name[64];
        char category[32];
        int credits = 0;
        float rating = 3.5f;
        int is_required = 0;
        char *time_field;
        Course *course;

        line_no++;
        field_count = parse_csv_line(line, fields, 16);
        if (field_count < 8) {
            continue;
        }

        if (line_no == 1 && (contains_ignore_case(fields[0], "code") || strstr(fields[0], "과목") != NULL)) {
            continue;
        }

        normalize_text_field(code, sizeof(code), fields[0]);
        normalize_text_field(name, sizeof(name), fields[1]);
        if (name[0] == '\0') {
            normalize_text_field(name, sizeof(name), fields[2]);
        }
        normalize_text_field(category, sizeof(category), fields[6]);

        credits = atoi(fields[3]);
        if (credits <= 0) {
            credits = 3;
        }

        is_required = is_required_category(category);

        course = find_or_create_course(out, &course_count, max_courses, code, name, category, credits, rating, is_required);
        if (course == NULL) {
            break;
        }

        time_field = fields[7];
        trim(time_field);
        if (time_field[0] == '\0') {
            continue;
        }

        {
            char temp[MAX_FIELD];
            char *token;
            snprintf(temp, sizeof(temp), "%s", time_field);
            token = strtok(temp, ",");
            while (token != NULL && course->slot_count < MAX_SLOTS) {
                TimeSlot slot;
                trim(token);
                if (parse_time_range(token, &slot)) {
                    course->slots[course->slot_count++] = slot;
                }
                token = strtok(NULL, ",");
            }
        }
    }

    fclose(fp);
    delete_file_if_exists(csv_path);
    return course_count;
}

void createTimetable(Timetable *tt) {
    memset(tt, 0, sizeof(*tt));
}

int isConflict(const Timetable *tt, const Course *c) {
    int i;
    for (i = 0; i < c->slot_count; ++i) {
        int day = c->slots[i].day;
        int s = c->slots[i].start_min / 30;
        int e = c->slots[i].end_min / 30;
        int t;
        if (day < 0 || day >= DAY_COUNT) {
            continue;
        }
        if (s < 0) s = 0;
        if (e > SLOT_COUNT_PER_DAY) e = SLOT_COUNT_PER_DAY;
        for (t = s; t < e; ++t) {
            if (tt->occupied[day][t]) {
                return 1;
            }
        }
    }
    return 0;
}

int addCourse(Timetable *tt, Course *c) {
    int i;
    if (tt->enrolled_count >= MAX_ENROLLED) {
        return 0;
    }
    if (isConflict(tt, c)) {
        return 0;
    }

    tt->enrolled[tt->enrolled_count++] = c;
    tt->total_credits += c->credits;

    for (i = 0; i < c->slot_count; ++i) {
        int day = c->slots[i].day;
        int s = c->slots[i].start_min / 30;
        int e = c->slots[i].end_min / 30;
        int t;
        if (day < 0 || day >= DAY_COUNT) continue;
        if (s < 0) s = 0;
        if (e > SLOT_COUNT_PER_DAY) e = SLOT_COUNT_PER_DAY;
        for (t = s; t < e; ++t) {
            tt->occupied[day][t] = 1;
        }
    }
    return 1;
}

void removeCourse(Timetable *tt, Course *c) {
    int idx = -1;
    int i;
    for (i = 0; i < tt->enrolled_count; ++i) {
        if (tt->enrolled[i] == c) {
            idx = i;
            break;
        }
    }
    if (idx < 0) {
        return;
    }

    for (i = idx; i < tt->enrolled_count - 1; ++i) {
        tt->enrolled[i] = tt->enrolled[i + 1];
    }
    tt->enrolled_count--;
    tt->total_credits -= c->credits;

    for (i = 0; i < c->slot_count; ++i) {
        int day = c->slots[i].day;
        int s = c->slots[i].start_min / 30;
        int e = c->slots[i].end_min / 30;
        int t;
        if (day < 0 || day >= DAY_COUNT) continue;
        if (s < 0) s = 0;
        if (e > SLOT_COUNT_PER_DAY) e = SLOT_COUNT_PER_DAY;
        for (t = s; t < e; ++t) {
            int covered = 0;
            int j;
            for (j = 0; j < tt->enrolled_count; ++j) {
                Course *ec = tt->enrolled[j];
                int k;
                for (k = 0; k < ec->slot_count; ++k) {
                    int es = ec->slots[k].start_min / 30;
                    int ee = ec->slots[k].end_min / 30;
                    if (ec->slots[k].day == day && t >= es && t < ee) {
                        covered = 1;
                        break;
                    }
                }
                if (covered) break;
            }
            tt->occupied[day][t] = covered;
        }
    }
}

int getTotalCredits(const Timetable *tt) {
    return tt->total_credits;
}

const TimeSlot *getSchedule(const Course *c, int *count_out) {
    if (count_out != NULL) {
        *count_out = c->slot_count;
    }
    return c->slots;
}

float getRating(const Course *c) {
    return c->rating;
}

int isLeaf(const SearchNode *node, const Preference *pref) {
    return node->current.total_credits >= pref->target_credits;
}

float scoreTimetable(const Timetable *tt, const Preference *pref) {
    int i;
    float rating_sum = 0.0f;
    int free_days = 0;
    float avg_start_hour_sum = 0.0f;
    int active_days = 0;

    for (i = 0; i < tt->enrolled_count; ++i) {
        rating_sum += tt->enrolled[i]->rating;
    }

    for (i = 0; i < 5; ++i) {
        int t;
        int first = -1;
        int used = 0;
        for (t = 0; t < SLOT_COUNT_PER_DAY; ++t) {
            if (tt->occupied[i][t]) {
                used = 1;
                if (first < 0) first = t;
            }
        }
        if (!used) {
            free_days++;
        } else {
            avg_start_hour_sum += (float)first / 2.0f;
            active_days++;
        }
    }

    {
        float avg_rating = (tt->enrolled_count > 0) ? rating_sum / (float)tt->enrolled_count : 0.0f;
        float late_start = (active_days > 0) ? (avg_start_hour_sum / (float)active_days) : 12.0f;
        return pref->prefer_high_rating * avg_rating
            + pref->prefer_free_days * (float)free_days
            + pref->prefer_late_start * late_start;
    }
}

static int compare_timetable(const Timetable *a, const Timetable *b) {
    int i;
    if (a->score > b->score) return -1;
    if (a->score < b->score) return 1;
    if (a->total_credits > b->total_credits) return -1;
    if (a->total_credits < b->total_credits) return 1;

    for (i = 0; i < a->enrolled_count && i < b->enrolled_count; ++i) {
        int cmp = strcmp(a->enrolled[i]->code, b->enrolled[i]->code);
        if (cmp != 0) return cmp;
    }
    if (a->enrolled_count < b->enrolled_count) return -1;
    if (a->enrolled_count > b->enrolled_count) return 1;
    return 0;
}

void insertResult(ResultSet *rs, const Timetable *tt, int top_n) {
    Timetable candidate;
    int pos;
    int i;

    if (top_n <= 0) return;
    if (top_n > MAX_RESULTS) top_n = MAX_RESULTS;

    candidate = *tt;

    for (i = 0; i < rs->count; ++i) {
        if (timetable_equal_key(&rs->items[i], &candidate)) {
            return;
        }
    }

    pos = rs->count;
    while (pos > 0 && compare_timetable(&candidate, &rs->items[pos - 1]) < 0) {
        if (pos < MAX_RESULTS) {
            rs->items[pos] = rs->items[pos - 1];
        }
        pos--;
    }

    if (pos < top_n) {
        rs->items[pos] = candidate;
        if (rs->count < top_n) {
            rs->count++;
        }
    }

    if (rs->count > top_n) {
        rs->count = top_n;
    }
}

int appendSch(Timetable *tt, Course *course) {
    return addCourse(tt, course);
}

int addSch(Timetable *tt, Course *course) {
    return addCourse(tt, course);
}

static int can_reach_target_with_suffix(const SearchContext *ctx, int idx, int current_credits) {
    int possible_max = current_credits + ctx->suffix_credits[idx];
    return possible_max >= ctx->pref->target_credits;
}

static int exceeds_credit_limit(const Preference *pref, int total_credits) {
    return total_credits > pref->target_credits + MAX_CREDIT_OVERFLOW;
}

static void dfs_generate(SearchContext *ctx, int idx, Timetable *tt) {
    SearchNode node;

    node.level = idx;
    node.current = *tt;

    if (isLeaf(&node, ctx->pref)) {
        if (all_must_include_satisfied(tt, ctx->pref)) {
            tt->score = scoreTimetable(tt, ctx->pref);
            insertResult(ctx->out, tt, ctx->pref->top_n);
        }
        /* continue exploring only if still under small overflow limit */
        if (exceeds_credit_limit(ctx->pref, tt->total_credits)) {
            return;
        }
    }

    if (idx >= ctx->course_count) {
        return;
    }

    if (!can_reach_target_with_suffix(ctx, idx, tt->total_credits)) {
        return;
    }

    {
        Course *course = (Course *)&ctx->courses[idx];

        if (!exceeds_credit_limit(ctx->pref, tt->total_credits + course->credits) && addCourse(tt, course)) {
            dfs_generate(ctx, idx + 1, tt);
            removeCourse(tt, course);
        }

        if (!course->is_required && !course_is_must_include(course, ctx->pref)) {
            dfs_generate(ctx, idx + 1, tt);
        }
    }
}

void createSch(const Course *courses, int course_count, const Preference *pref, ResultSet *out) {
    SearchContext ctx;
    Timetable tt;
    int i;

    memset(&ctx, 0, sizeof(ctx));
    ctx.courses = courses;
    ctx.course_count = course_count;
    ctx.pref = pref;
    ctx.out = out;

    ctx.suffix_credits[course_count] = 0;
    for (i = course_count - 1; i >= 0; --i) {
        ctx.suffix_credits[i] = ctx.suffix_credits[i + 1] + courses[i].credits;
    }

    createTimetable(&tt);
    memset(out, 0, sizeof(*out));
    dfs_generate(&ctx, 0, &tt);
}

static int parse_args(int argc, char **argv, char *xlsx_path, size_t xlsx_path_cap, Preference *pref) {
    int i;
    memset(pref, 0, sizeof(*pref));

    pref->target_credits = 18;
    pref->top_n = 5;
    pref->prefer_free_days = 3;
    pref->prefer_late_start = 1;
    pref->prefer_high_rating = 3;

    xlsx_path[0] = '\0';

    for (i = 1; i < argc; ++i) {
        if (strcmp(argv[i], "--xlsx") == 0 && i + 1 < argc) {
            snprintf(xlsx_path, xlsx_path_cap, "%s", argv[++i]);
        } else if (strcmp(argv[i], "--target-credits") == 0 && i + 1 < argc) {
            pref->target_credits = atoi(argv[++i]);
        } else if (strcmp(argv[i], "--top") == 0 && i + 1 < argc) {
            pref->top_n = atoi(argv[++i]);
        } else if (strcmp(argv[i], "--w-free") == 0 && i + 1 < argc) {
            pref->prefer_free_days = atoi(argv[++i]);
        } else if (strcmp(argv[i], "--w-late") == 0 && i + 1 < argc) {
            pref->prefer_late_start = atoi(argv[++i]);
        } else if (strcmp(argv[i], "--w-rating") == 0 && i + 1 < argc) {
            pref->prefer_high_rating = atoi(argv[++i]);
        } else if (strcmp(argv[i], "--must") == 0 && i + 1 < argc) {
            char temp[1024];
            char *tok;
            snprintf(temp, sizeof(temp), "%s", argv[++i]);
            tok = strtok(temp, ",");
            while (tok != NULL && pref->must_include_count < MAX_MUST_INCLUDE) {
                trim(tok);
                snprintf(pref->must_include_codes[pref->must_include_count], sizeof(pref->must_include_codes[pref->must_include_count]), "%s", tok);
                pref->must_include_count++;
                tok = strtok(NULL, ",");
            }
        } else if (strcmp(argv[i], "--self-test") == 0) {
            g_self_test = 1;
        } else if (strcmp(argv[i], "--help") == 0) {
            print_usage();
            return 0;
        } else {
            fprintf(stderr, "[ERROR] Unknown argument: %s\n", argv[i]);
            return 0;
        }
    }

    if (pref->top_n <= 0) pref->top_n = 1;
    if (pref->top_n > MAX_RESULTS) pref->top_n = MAX_RESULTS;

    if (!g_self_test && xlsx_path[0] == '\0') {
        fprintf(stderr, "[ERROR] --xlsx is required unless --self-test is used.\n");
        return 0;
    }

    return 1;
}

static void print_usage(void) {
    printf("Usage:\n");
    printf("  timetable.exe --xlsx <path> [--target-credits N] [--top N]\\n");
    printf("               [--w-free N] [--w-late N] [--w-rating N]\\n");
    printf("               [--must CODE1,CODE2,...] [--self-test]\\n");
}

static void trim(char *s) {
    char *start;
    char *end;
    if (s == NULL) return;
    start = s;
    while (*start && isspace((unsigned char)*start)) start++;

    if (start != s) {
        memmove(s, start, strlen(start) + 1);
    }

    if (*s == '\0') {
        return;
    }

    end = s + strlen(s) - 1;
    while (end > s && (isspace((unsigned char)*end) || *end == '\r' || *end == '\n')) {
        *end = '\0';
        end--;
    }
}

static int starts_with_ignore_case(const char *s, const char *prefix) {
    while (*prefix && *s) {
        if (tolower((unsigned char)*s) != tolower((unsigned char)*prefix)) {
            return 0;
        }
        s++;
        prefix++;
    }
    return *prefix == '\0';
}

static int contains_ignore_case(const char *s, const char *needle) {
    size_t nlen = strlen(needle);
    size_t i;
    if (nlen == 0) return 1;
    for (i = 0; s[i] != '\0'; ++i) {
        if (tolower((unsigned char)s[i]) == tolower((unsigned char)needle[0])) {
            size_t j;
            for (j = 0; j < nlen; ++j) {
                if (s[i + j] == '\0') return 0;
                if (tolower((unsigned char)s[i + j]) != tolower((unsigned char)needle[j])) {
                    break;
                }
            }
            if (j == nlen) return 1;
        }
    }
    return 0;
}

static int parse_csv_line(const char *line, char fields[][MAX_FIELD], int max_fields) {
    int fi = 0;
    int in_quotes = 0;
    int ci = 0;
    int i;

    if (max_fields <= 0) return 0;

    memset(fields, 0, sizeof(char[MAX_FIELD]) * (size_t)max_fields);

    for (i = 0; line[i] != '\0'; ++i) {
        char ch = line[i];
        if (ch == '"') {
            if (in_quotes && line[i + 1] == '"') {
                if (ci < MAX_FIELD - 1) fields[fi][ci++] = '"';
                i++;
            } else {
                in_quotes = !in_quotes;
            }
        } else if (ch == ',' && !in_quotes) {
            if (fi + 1 >= max_fields) {
                break;
            }
            fields[fi][ci] = '\0';
            fi++;
            ci = 0;
        } else if (ch == '\r' || ch == '\n') {
            break;
        } else {
            if (ci < MAX_FIELD - 1) fields[fi][ci++] = ch;
        }
    }

    fields[fi][ci] = '\0';
    return fi + 1;
}

static void normalize_text_field(char *dst, size_t dst_cap, const char *src) {
    size_t j = 0;
    size_t i;
    if (src != NULL && (unsigned char)src[0] == 0xEF && (unsigned char)src[1] == 0xBB && (unsigned char)src[2] == 0xBF) {
        src += 3;
    }
    while (src != NULL && isspace((unsigned char)*src)) src++;
    for (i = 0; src != NULL && src[i] != '\0' && j + 1 < dst_cap; ++i) {
        if (src[i] == '\r' || src[i] == '\n') {
            break;
        }
        dst[j++] = src[i];
    }
    while (j > 0 && isspace((unsigned char)dst[j - 1])) {
        j--;
    }
    dst[j] = '\0';
}

static int parse_day_token(const char *token) {
    if (token == NULL || token[0] == '\0') return -1;

    if (starts_with_ignore_case(token, "Mon")) return 0;
    if (starts_with_ignore_case(token, "Tue")) return 1;
    if (starts_with_ignore_case(token, "Wed")) return 2;
    if (starts_with_ignore_case(token, "Thu")) return 3;
    if (starts_with_ignore_case(token, "Fri")) return 4;
    if (starts_with_ignore_case(token, "Sat")) return 5;
    if (starts_with_ignore_case(token, "Sun")) return 6;

    if (strstr(token, "월") != NULL) return 0;
    if (strstr(token, "화") != NULL) return 1;
    if (strstr(token, "수") != NULL) return 2;
    if (strstr(token, "목") != NULL) return 3;
    if (strstr(token, "금") != NULL) return 4;
    if (strstr(token, "토") != NULL) return 5;
    if (strstr(token, "일") != NULL) return 6;

    return -1;
}

static int parse_time_hhmm(const char *s, int *out_minutes) {
    int hh, mm;
    if (sscanf(s, "%d:%d", &hh, &mm) != 2) {
        return 0;
    }
    if (hh < 0 || hh >= 24 || mm < 0 || mm >= 60) {
        return 0;
    }
    *out_minutes = hh * 60 + mm;
    return 1;
}

static int parse_time_range(const char *token, TimeSlot *slot_out) {
    const char *dash;
    const char *colon;
    char left[64];
    char right[32];
    int day;
    int start;
    int end;
    size_t left_len;
    size_t day_len = 0;

    if (token == NULL || token[0] == '\0') return 0;
    if (contains_ignore_case(token, "온라인") || contains_ignore_case(token, "미정")) return 0;

    dash = strchr(token, '-');
    if (dash == NULL) return 0;

    colon = strchr(token, ':');
    if (colon == NULL) return 0;

    {
        const char *p = colon;
        while (p > token && isdigit((unsigned char)p[-1])) {
            p--;
        }
        day_len = (size_t)(p - token);
    }

    if (day_len == 0 || day_len >= sizeof(left)) return 0;
    memcpy(left, token, day_len);
    left[day_len] = '\0';

    day = parse_day_token(left);
    if (day < 0) return 0;

    left_len = (size_t)(dash - (token + day_len));
    if (left_len == 0 || left_len >= sizeof(left)) return 0;
    memcpy(left, token + day_len, left_len);
    left[left_len] = '\0';

    snprintf(right, sizeof(right), "%s", dash + 1);
    trim(left);
    trim(right);

    if (!parse_time_hhmm(left, &start)) return 0;
    if (!parse_time_hhmm(right, &end)) return 0;
    if (end <= start) return 0;

    slot_out->day = day;
    slot_out->start_min = start;
    slot_out->end_min = end;
    return 1;
}

static int is_required_category(const char *category) {
    if (category == NULL) return 0;
    if (strstr(category, "전필") != NULL) return 1;
    if (strstr(category, "전공필수") != NULL) return 1;
    if (strstr(category, "필수") != NULL) return 1;
    if (contains_ignore_case(category, "required")) return 1;
    return 0;
}

static int course_is_must_include(const Course *course, const Preference *pref) {
    int i;
    for (i = 0; i < pref->must_include_count; ++i) {
        if (strcmp(course->code, pref->must_include_codes[i]) == 0) {
            return 1;
        }
    }
    return 0;
}

static int all_must_include_satisfied(const Timetable *tt, const Preference *pref) {
    int i;
    for (i = 0; i < pref->must_include_count; ++i) {
        int found = 0;
        int j;
        for (j = 0; j < tt->enrolled_count; ++j) {
            if (strcmp(tt->enrolled[j]->code, pref->must_include_codes[i]) == 0) {
                found = 1;
                break;
            }
        }
        if (!found) return 0;
    }
    return 1;
}

static const char *day_name(int day) {
    static const char *names[] = {"Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"};
    if (day < 0 || day >= DAY_COUNT) return "?";
    return names[day];
}

static void format_hhmm(int minutes, char *buf, size_t cap) {
    int hh = minutes / 60;
    int mm = minutes % 60;
    snprintf(buf, cap, "%02d:%02d", hh, mm);
}

static void print_course(const Course *c) {
    int i;
    printf("  - %s | %s | %s | %dcr\n", c->code, c->name, c->category, c->credits);
    if (c->slot_count == 0) {
        printf("      time: (no fixed slot)\n");
        return;
    }
    for (i = 0; i < c->slot_count; ++i) {
        char s[8], e[8];
        format_hhmm(c->slots[i].start_min, s, sizeof(s));
        format_hhmm(c->slots[i].end_min, e, sizeof(e));
        printf("      time: %s %s-%s\n", day_name(c->slots[i].day), s, e);
    }
}

static void print_timetable(const Timetable *tt, int rank) {
    int i;
    printf("\n========== Rank %d ==========" "\n", rank);
    printf("Total Credits: %d\n", tt->total_credits);
    printf("Score: %.2f\n", tt->score);
    printf("Courses (%d):\n", tt->enrolled_count);
    for (i = 0; i < tt->enrolled_count; ++i) {
        print_course(tt->enrolled[i]);
    }
}

static void delete_file_if_exists(const char *path) {
    if (path == NULL || path[0] == '\0') return;
    remove(path);
}

static int timetable_equal_key(const Timetable *a, const Timetable *b) {
    int i;
    if (a->enrolled_count != b->enrolled_count) return 0;
    for (i = 0; i < a->enrolled_count; ++i) {
        if (strcmp(a->enrolled[i]->code, b->enrolled[i]->code) != 0) return 0;
    }
    return 1;
}

static int run_self_test(void) {
    Course a = {"A", "Alpha", "REQ", 3, 4.2f, 1, 1, {{0, 540, 615}}};
    Course b = {"B", "Beta", "SEL", 3, 3.9f, 0, 1, {{0, 570, 645}}};
    Course c = {"C", "Gamma", "SEL", 3, 4.5f, 0, 1, {{1, 540, 615}}};
    Course list[3];
    Preference p;
    ResultSet rs;
    Timetable tt;

    list[0] = a; list[1] = b; list[2] = c;

    createTimetable(&tt);
    if (!addCourse(&tt, &list[0])) return 0;
    if (addCourse(&tt, &list[1])) return 0; /* must conflict */
    if (!addCourse(&tt, &list[2])) return 0;
    if (getTotalCredits(&tt) != 6) return 0;

    memset(&p, 0, sizeof(p));
    p.target_credits = 6;
    p.top_n = 2;
    p.prefer_free_days = 3;
    p.prefer_late_start = 1;
    p.prefer_high_rating = 3;

    createSch(list, 3, &p, &rs);
    if (rs.count <= 0) return 0;

    printf("[SELF-TEST] all checks passed (%d results)\n", rs.count);
    return 1;
}

int main(int argc, char **argv) {
    char xlsx_path[1024];
    Preference pref;
    Course courses[MAX_COURSES];
    int course_count;
    ResultSet results;
    int i;

    if (!parse_args(argc, argv, xlsx_path, sizeof(xlsx_path), &pref)) {
        print_usage();
        return 1;
    }

    if (g_self_test) {
        return run_self_test() ? 0 : 1;
    }

    course_count = loadCoursesFromXlsx(xlsx_path, courses, MAX_COURSES);
    if (course_count <= 0) {
        fprintf(stderr, "[ERROR] No courses loaded from xlsx.\n");
        return 1;
    }

    printf("Loaded courses: %d\n", course_count);

    createSch(courses, course_count, &pref, &results);

    if (results.count == 0) {
        printf("No valid timetable matched the constraints.\n");
        return 0;
    }

    for (i = 0; i < results.count; ++i) {
        print_timetable(&results.items[i], i + 1);
    }

    return 0;
}

import java.net.URI;
import java.net.URLEncoder;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDateTime;
import java.time.LocalDate;
import java.time.DayOfWeek;
import java.time.YearMonth;
import java.time.Duration;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

/**
 * WBSO hours substantiation for Jira project VALCM — dependency-free (JDK only).
 *
 * Pulls every issue labelled "wbso", reconstructs each issue's status + assignee
 * history from the full changelog, converts story points to an estimate of hours
 * worked per person, per RVO project, per working day for a single calendar month,
 * and writes one .xlsx with tabs: Summary, Daily, DayTotals, Evidence, Detail,
 * Issues, StatusAudit, Notes.
 *
 * Methodology (see Notes tab in the output):
 *   - 1 story point = 4 hours.
 *   - Base = Story Points DEV (customfield_12867) + QA (customfield_12904), each x4.
 *     The total field (customfield_10025) is used only when both DEV and QA are empty.
 *   - DEV hours spread over whoever held the ticket during DEV-phase statuses, QA hours
 *     over QA-phase statuses; time-weighted. Each month receives the share of the
 *     budget equal to phase-time inside that month / phase-time over the issue's whole
 *     life, so monthly runs are consistent and sum to the issue total. Derived estimate.
 *   - Time held is measured on working days only (Mon-Fri) and split per calendar day,
 *     so the Daily tab is the per-person / per-project / per-day record RVO asks for.
 *   - --eligible <file>: one Jira display name per line (Validaide B.V. payroll only).
 *     Only listed people are counted; everyone else is shown separately so totals reconcile.
 *   - --rvo-project <id>: the S&O project number from the S&O-verklaring (default 2601).
 *     An issue label "wbso-<id>" overrides it per issue.
 *
 * Build:  javac WbsoHours.java
 * Run:    export JIRA_EMAIL="you@company.com"
 *         export JIRA_API_TOKEN="xxxxx"
 *         java WbsoHours --month 2026-01 --out wbso_2026-01.xlsx
 *         (optional: --project VALCM  --label wbso  --base https://skycell.atlassian.net
 *                    --eligible eligible.txt  --rvo-project 2601)
 *         java WbsoHours --selftest        # runs on synthetic data, no network
 *
 * SECURITY: the token is read from an environment variable and never written to disk.
 */
public class WbsoHours {

    // ---- configuration constants ----
    static final double HOURS_PER_POINT = 4.0;
    static final String CF_DEV = "customfield_12867", CF_QA = "customfield_12904", CF_TOTAL = "customfield_10025";
    static final Set<String> DEV_STATUSES = new HashSet<>(Arrays.asList(
            "under development", "in development", "under review", "verifying functionality"));
    static final Set<String> QA_STATUSES = new HashSet<>(Arrays.asList(
            "under testing", "under test review"));

    // ============================ data model ============================
    static final class Issue {
        String key, summary, issueType, currentStatus, currentAssignee, created, resolved, updated;
        Double devSp, qaSp, totalSp; List<String> labels = new ArrayList<>();
    }
    static final class Change {       // one status/assignee change from the changelog
        String key, field, fromStr, toStr, author; LocalDateTime ts;
    }
    static final class Detail {
        String issue, person, month, phase, method, rvo; double hours, timeHeld;
        Detail(String i, String p, String m, String ph, double h, double t, String me, String r) {
            issue = i; person = p; month = m; phase = ph; hours = h; timeHeld = t; method = me; rvo = r;
        }
    }
    static final class Daily {        // one row of the per-person / per-project / per-day record
        String issue, person, phase, method, rvo; LocalDate date; double hours;
        Daily(String i, String p, LocalDate d, String ph, double h, String me, String r) {
            issue = i; person = p; date = d; phase = ph; hours = h; method = me; rvo = r;
        }
    }
    static final class Meta {
        String issue, summary, issueType, status, rvo, method = "", note = "";
        double devSp, qaSp, totalSp, devBudget, qaBudget, devAlloc, qaAlloc;
    }

    // ============================ main ============================
    public static void main(String[] args) throws Exception {
        Map<String, String> opt = new HashMap<>();
        boolean selftest = false;
        for (int i = 0; i < args.length; i++) {
            if (args[i].equals("--selftest")) selftest = true;
            else if (args[i].startsWith("--") && i + 1 < args.length) opt.put(args[i].substring(2), args[++i]);
        }
        String project = opt.getOrDefault("project", "VALCM");
        String label   = opt.getOrDefault("label", "wbso");
        String monthArg = opt.getOrDefault("month", YearMonth.now().toString());
        if (opt.containsKey("start"))
            System.err.println("Note: --start is no longer used; this version reports a single month via --month YYYY-MM.");
        String out     = opt.getOrDefault("out", "wbso_hours.xlsx");
        String base    = opt.getOrDefault("base", env("JIRA_BASE_URL", "https://skycell.atlassian.net"));
        String rvoDefault = opt.getOrDefault("rvo-project", "2601");
        String eligibleFile = opt.get("eligible");
        Set<String> eligible = eligibleFile == null ? null : loadEligible(eligibleFile);
        LocalDateTime winStart = parseMonth(monthArg);
        LocalDateTime winEnd   = YearMonth.from(winStart).plusMonths(1).atDay(1).atStartOfDay();
        String targetMonth = YearMonth.from(winStart).toString();
        System.out.println("Reporting month: " + targetMonth);
        System.out.println("RVO project (default): " + rvoDefault + (eligible == null
                ? "  |  eligibility list: NONE (every named assignee is counted; pass --eligible <file>)"
                : "  |  eligibility list: " + eligibleFile + " (" + eligible.size() + " names)"));

        List<Issue> issues; Map<String, List<Change>> logs;
        if (selftest) {
            Object[] syn = synthetic();
            issues = (List<Issue>) syn[0];
            logs   = (Map<String, List<Change>>) syn[1];
            System.out.println("SELFTEST: " + issues.size() + " synthetic issues");
        } else {
            String email = env("JIRA_EMAIL", null), token = env("JIRA_API_TOKEN", null);
            if (email == null || token == null) {
                System.err.println("Set JIRA_EMAIL and JIRA_API_TOKEN environment variables.");
                System.exit(2); return;
            }
            Jira jira = new Jira(base, email, token);
            String jql = "project = " + project + " AND labels = " + label + " ORDER BY key ASC";
            System.out.println("Searching: " + jql);
            issues = jira.searchIssues(jql);
            System.out.println("  found " + issues.size() + " issues");
            logs = new LinkedHashMap<>();
            int n = 0;
            for (Issue is : issues) {
                System.out.print("  [" + (++n) + "/" + issues.size() + "] " + is.key + " changelog…");
                List<Change> cl = jira.changelog(is.key);
                logs.put(is.key, cl);
                System.out.println(" " + cl.size() + " changes");
            }
        }

        List<Detail> detail = new ArrayList<>();
        List<Daily> daily = new ArrayList<>();
        List<Meta> metas = new ArrayList<>();
        TreeMap<String, String> statusAudit = new TreeMap<>();
        for (Issue is : issues) allocate(is, logs.getOrDefault(is.key, Collections.emptyList()), winStart, winEnd, targetMonth, rvoDefault, detail, daily, metas, statusAudit);

        double total = detail.stream().mapToDouble(d -> d.hours).sum();
        double counted = detail.stream().filter(d -> isEligible(eligible, d.person)).mapToDouble(d -> d.hours).sum();
        Set<String> people = new TreeSet<>();
        for (Detail d : detail) people.add(d.person);
        System.out.printf("Grand total: %.1f h across %d people for %s%n", total, people.size(), targetMonth);
        System.out.printf("Counted: %.1f h  |  not counted (not on the list, or unassigned): %.1f h%n", counted, total - counted);

        writeWorkbook(out, detail, daily, metas, logs, statusAudit, project, label, targetMonth, base, rvoDefault, eligible, eligibleFile, winStart, winEnd);
        System.out.println("Wrote " + out);
    }

    static String env(String k, String def) { String v = System.getenv(k); return v == null ? def : v; }

    // ============================ eligibility / project / working days ============================
    static Set<String> loadEligible(String file) throws Exception {
        Set<String> s = new TreeSet<>();
        for (String line : Files.readAllLines(Path.of(file), StandardCharsets.UTF_8)) {
            String t = line.trim();
            if (t.isEmpty() || t.startsWith("#")) continue;
            s.add(t.toLowerCase());
        }
        return s;
    }
    // null list = no filter configured: every named assignee counts (flagged in Notes). "(unassigned)" never counts.
    static boolean isEligible(Set<String> eligible, String person) {
        if (person == null || person.equals("(unassigned)")) return false;
        return eligible == null || eligible.contains(person.trim().toLowerCase());
    }
    static String eligLabel(Set<String> eligible, String person) {
        if (eligible == null) return person.equals("(unassigned)") ? "no" : "n/a (no list)";
        return isEligible(eligible, person) ? "yes" : "no";
    }
    // RVO project number: an issue label "wbso-<id>" wins, otherwise --rvo-project.
    static String rvoProject(Issue is, String def) {
        for (String l : is.labels) if (l.toLowerCase().startsWith("wbso-") && l.length() > 5) return l.substring(5);
        return def;
    }
    // Seconds of [a, b) per calendar day, Saturdays and Sundays excluded.
    static Map<LocalDate, Double> workingSecs(LocalDateTime a, LocalDateTime b) {
        Map<LocalDate, Double> m = new LinkedHashMap<>();
        if (a == null || b == null || !b.isAfter(a)) return m;
        for (LocalDate d = a.toLocalDate(); !d.isAfter(b.toLocalDate()); d = d.plusDays(1)) {
            LocalDateTime s0 = d.atStartOfDay(), s1 = d.plusDays(1).atStartOfDay();
            LocalDateTime lo = a.isAfter(s0) ? a : s0, hi = b.isBefore(s1) ? b : s1;
            DayOfWeek dow = d.getDayOfWeek();
            if (hi.isAfter(lo) && dow != DayOfWeek.SATURDAY && dow != DayOfWeek.SUNDAY) m.put(d, secs(lo, hi));
        }
        return m;
    }
    static String nz(String s) { return s == null ? "" : s; }

    // ============================ allocation ============================
    // Hours for an issue's point-budget attributable to the target month = budget x
    // (phase-time inside the window) / (phase-time over the issue's whole life).
    // Run per month, the months sum to the issue total — no double counting.
    static void allocate(Issue is, List<Change> changes, LocalDateTime winStart, LocalDateTime winEnd, String targetMonth,
                         String rvoDefault, List<Detail> detail, List<Daily> daily, List<Meta> metas, Map<String, String> audit) {
        Meta m = new Meta();
        m.issue = is.key; m.summary = is.summary; m.issueType = is.issueType; m.status = is.currentStatus; m.rvo = rvoProject(is, rvoDefault);
        m.devSp = nz(is.devSp); m.qaSp = nz(is.qaSp); m.totalSp = nz(is.totalSp);

        String baseMode;
        double devBudget, qaBudget;
        if (m.devSp > 0 || m.qaSp > 0) { baseMode = "DEV+QA"; devBudget = m.devSp * HOURS_PER_POINT; qaBudget = m.qaSp * HOURS_PER_POINT; }
        else if (m.totalSp > 0) { baseMode = "TOTAL"; devBudget = m.totalSp * HOURS_PER_POINT; qaBudget = 0; }
        else { baseMode = "NONE"; devBudget = qaBudget = 0; }
        m.devBudget = round(devBudget, 2); m.qaBudget = round(qaBudget, 2);

        if (baseMode.equals("NONE")) {
            m.method = "skip"; m.note = "No story points on any field.";
            metas.add(m); return;
        }

        LocalDateTime created = parseTs(is.created);
        LocalDateTime endTime = is.resolved != null ? parseTs(is.resolved)
                              : is.updated != null ? parseTs(is.updated) : LocalDateTime.now();

        List<Change> statusCh = new ArrayList<>(), assignCh = new ArrayList<>(), all = new ArrayList<>();
        for (Change c : changes) {
            if (c.field.equals("status")) statusCh.add(c);
            else if (c.field.equals("assignee")) assignCh.add(c);
        }
        all.addAll(statusCh); all.addAll(assignCh);
        all.sort(Comparator.comparing(c -> c.ts));
        statusCh.sort(Comparator.comparing(c -> c.ts));
        assignCh.sort(Comparator.comparing(c -> c.ts));

        String curS = statusCh.isEmpty() ? is.currentStatus : statusCh.get(0).fromStr;
        String curA = assignCh.isEmpty() ? is.currentAssignee : assignCh.get(0).fromStr;
        LocalDateTime lastTs = created;

        List<Object[]> intervals = new ArrayList<>(); // {start, end, status, assignee}
        for (Change c : all) {
            if (c.ts.isAfter(lastTs)) intervals.add(new Object[]{lastTs, c.ts, curS, curA});
            if (c.field.equals("status")) curS = c.toStr; else curA = c.toStr;
            lastTs = c.ts;
        }
        if (endTime.isAfter(lastTs)) intervals.add(new Object[]{lastTs, endTime, curS, curA});

        // whole-life phase totals (denominators) + per-person, per-day time inside the window (numerators).
        // Time is counted on working days only (Mon-Fri), in the denominator and the numerator alike.
        double lifeDev = 0, lifeQa = 0;
        Map<String, Double> winDev = new LinkedHashMap<>(), winQa = new LinkedHashMap<>();
        Map<String, Map<LocalDate, Double>> dayDev = new LinkedHashMap<>(), dayQa = new LinkedHashMap<>();
        for (Object[] iv : intervals) {
            String st = (String) iv[2];
            if (st != null) audit.put(norm(st), classify(st));
            String phase = classify(st);
            if (phase == null) continue;
            LocalDateTime s0 = (LocalDateTime) iv[0], s1 = (LocalDateTime) iv[1];
            double full = 0; for (double v : workingSecs(s0, s1).values()) full += v;
            if (phase.equals("DEV")) lifeDev += full; else lifeQa += full;
            LocalDateTime lo = s0.isAfter(winStart) ? s0 : winStart;
            LocalDateTime hi = s1.isBefore(winEnd) ? s1 : winEnd;
            if (hi.isAfter(lo)) {
                String person = iv[3] == null ? "(unassigned)" : (String) iv[3];
                Map<String, Double> win = phase.equals("DEV") ? winDev : winQa;
                Map<LocalDate, Double> days = (phase.equals("DEV") ? dayDev : dayQa).computeIfAbsent(person, k -> new TreeMap<>());
                for (Map.Entry<LocalDate, Double> e : workingSecs(lo, hi).entrySet()) {
                    win.merge(person, e.getValue(), Double::sum);
                    days.merge(e.getKey(), e.getValue(), Double::sum);
                }
            }
        }

        if (baseMode.equals("DEV+QA")) {
            allocPhase(is.key, winDev, dayDev, lifeDev, devBudget, "DEV", targetMonth, detail, daily, m);
            allocPhase(is.key, winQa, dayQa, lifeQa, qaBudget, "QA", targetMonth, detail, daily, m);
            m.method = "time-weighted (DEV+QA)";
            fallbackPhase(is, devBudget, lifeDev, "DEV", winStart, winEnd, targetMonth, detail, daily, m);
            fallbackPhase(is, qaBudget, lifeQa, "QA", winStart, winEnd, targetMonth, detail, daily, m);
        } else { // TOTAL fallback: pool DEV+QA, one budget
            Map<String, Double> pooled = new LinkedHashMap<>();
            Map<String, Map<LocalDate, Double>> pooledDays = new LinkedHashMap<>();
            winDev.forEach((k, v) -> pooled.merge(k, v, Double::sum));
            winQa.forEach((k, v) -> pooled.merge(k, v, Double::sum));
            for (Map<String, Map<LocalDate, Double>> src : Arrays.asList(dayDev, dayQa))
                src.forEach((p, ds) -> ds.forEach((d, v) -> pooledDays.computeIfAbsent(p, k -> new TreeMap<>()).merge(d, v, Double::sum)));
            double lifeAll = lifeDev + lifeQa;
            allocPhase(is.key, pooled, pooledDays, lifeAll, devBudget, "DEV", targetMonth, detail, daily, m);
            m.method = "time-weighted (total fallback)";
            fallbackPhase(is, devBudget, lifeAll, "DEV", winStart, winEnd, targetMonth, detail, daily, m);
        }
        m.devAlloc = round(m.devAlloc, 2); m.qaAlloc = round(m.qaAlloc, 2);
        metas.add(m);
    }

    // distribute a phase budget across in-window holders, weighted by in-window time
    // over the issue's whole-life phase time. No-op when the phase has no timeline.
    static void allocPhase(String key, Map<String, Double> win, Map<String, Map<LocalDate, Double>> days, double life, double budget,
                          String phase, String month, List<Detail> detail, List<Daily> daily, Meta m) {
        if (budget <= 0 || life <= 0) return;
        for (Map.Entry<String, Double> e : win.entrySet()) {
            double w = e.getValue();
            if (w <= 0) continue;
            double hrs = budget * w / life;
            detail.add(new Detail(key, e.getKey(), month, phase, round(hrs, 3), round(w / 3600.0, 2), "time-weighted", m.rvo));
            for (Map.Entry<LocalDate, Double> d : days.getOrDefault(e.getKey(), Collections.emptyMap()).entrySet())
                daily.add(new Daily(key, e.getKey(), d.getKey(), phase, round(budget * d.getValue() / life, 3), "time-weighted", m.rvo));
            if (phase.equals("DEV")) m.devAlloc += hrs; else m.qaAlloc += hrs;
        }
    }

    // fires only when a phase has budget but NO timeline at all. Attributes the whole
    // phase budget to the current assignee, but only in the month it was last updated/
    // resolved — so across monthly runs it lands exactly once.
    static void fallbackPhase(Issue is, double budget, double life, String phase,
                             LocalDateTime winStart, LocalDateTime winEnd, String month,
                             List<Detail> detail, List<Daily> daily, Meta m) {
        if (budget <= 0 || life > 0) return;
        LocalDateTime attrib = is.resolved != null ? parseTs(is.resolved)
                             : is.updated != null ? parseTs(is.updated) : null;
        if (attrib == null) { m.note = note(m.note, phase + ": no timeline and no date; skipped."); return; }
        boolean inWindow = !attrib.isBefore(winStart) && attrib.isBefore(winEnd);
        if (!inWindow) { m.note = note(m.note, phase + ": no timeline; belongs to " + monthKey(attrib) + ", not this month."); return; }
        boolean reached = classify(is.currentStatus) != null || is.resolved != null;
        if (!reached) { m.note = note(m.note, phase + ": not yet worked."); return; }
        String person = is.currentAssignee == null ? "(unassigned)" : is.currentAssignee;
        detail.add(new Detail(is.key, person, month, phase, round(budget, 3), 0, "FALLBACK current-assignee", m.rvo));
        daily.add(new Daily(is.key, person, attrib.toLocalDate(), phase, round(budget, 3), "FALLBACK current-assignee", m.rvo));
        if (phase.equals("DEV")) m.devAlloc += budget; else m.qaAlloc += budget;
        m.method += " + FALLBACK";
        m.note = note(m.note, phase + ": no changelog timeline; attributed to current assignee this month.");
    }
    static String note(String existing, String add) { return existing == null || existing.isEmpty() ? add : existing + " " + add; }
    static double secs(LocalDateTime a, LocalDateTime b) { return Duration.between(a, b).getSeconds(); }

    // ============================ time helpers ============================
    static final DateTimeFormatter TSF = DateTimeFormatter.ofPattern("yyyy-MM-dd'T'HH:mm:ss");
    static LocalDateTime parseTs(String s) {
        if (s == null) return null;
        try { return LocalDateTime.parse(s.substring(0, 19), TSF); } catch (Exception e) { return null; }
    }
    static LocalDateTime parseMonth(String s) {
        int y = Integer.parseInt(s.substring(0, 4)), mo = Integer.parseInt(s.substring(5, 7));
        return LocalDateTime.of(y, mo, 1, 0, 0);
    }
    static String monthKey(LocalDateTime dt) { return YearMonth.from(dt).toString(); }
    static String norm(String st) { return st.replace(" (migrated)", "").trim().toLowerCase(); }
    static String classify(String st) {
        if (st == null) return null;
        String s = norm(st);
        if (DEV_STATUSES.contains(s)) return "DEV";
        if (QA_STATUSES.contains(s)) return "QA";
        return null;
    }
    static double nz(Double d) { return d == null ? 0 : d; }
    static double round(double v, int dp) { double f = Math.pow(10, dp); return Math.round(v * f) / f; }

    // ============================ Jira REST client ============================
    static final class Jira {
        final String base, auth;
        final HttpClient http = HttpClient.newBuilder().connectTimeout(Duration.ofSeconds(30)).build();
        static final int MAX_ATTEMPTS = 8;
        Jira(String base, String email, String token) {
            this.base = base.replaceAll("/+$", "");
            this.auth = "Basic " + Base64.getEncoder().encodeToString((email + ":" + token).getBytes(StandardCharsets.UTF_8));
        }
        // Retries transient failures with exponential backoff (2, 4, 8 ... 60 s): rate limits (429), server
        // errors (5xx), network hiccups, and the occasional HTTP 400 whose body wraps a backend 500
        // (Jira Cloud does this for database faults). Genuine client errors fail immediately.
        String getJson(String url) throws Exception {
            Exception last = null;
            for (int attempt = 1; attempt <= MAX_ATTEMPTS; attempt++) {
                long wait = Math.min(60000L, 2000L << (attempt - 1));
                try {
                    HttpRequest req = HttpRequest.newBuilder(URI.create(url)).timeout(Duration.ofSeconds(90))
                            .header("Authorization", auth).header("Accept", "application/json").GET().build();
                    HttpResponse<String> r = http.send(req, HttpResponse.BodyHandlers.ofString());
                    int sc = r.statusCode();
                    if (sc >= 200 && sc < 300) return r.body();
                    boolean transientFault = sc == 429 || sc >= 500 || (sc == 400 && looksLikeServerFault(r.body()));
                    if (!transientFault) throw new RuntimeException("HTTP " + sc + ": " + r.body());
                    if (sc == 429) wait = Math.max(wait, retryAfterMillis(r, 5000L));
                    last = new RuntimeException("HTTP " + sc + ": " + abbreviate(r.body()));
                    System.err.println("  transient Jira error (HTTP " + sc + "), retry " + attempt + "/" + MAX_ATTEMPTS + " in " + wait / 1000 + "s");
                } catch (java.io.IOException e) {   // connection reset, timeout, DNS blip
                    last = e;
                    System.err.println("  network error (" + e.getClass().getSimpleName() + "), retry " + attempt + "/" + MAX_ATTEMPTS + " in " + wait / 1000 + "s");
                }
                Thread.sleep(wait);
            }
            throw new RuntimeException("Gave up after " + MAX_ATTEMPTS + " attempts for " + url + " — last error: " + last, last);
        }
        static boolean looksLikeServerFault(String body) {
            if (body == null) return false;
            return body.contains("\"status-code\":5") || body.contains("Exception") || body.contains("Internal Server Error");
        }
        static long retryAfterMillis(HttpResponse<String> r, long def) {
            try { return r.headers().firstValue("Retry-After").map(v -> Long.parseLong(v.trim()) * 1000L).orElse(def); }
            catch (Exception e) { return def; }
        }
        static String abbreviate(String s) { return s == null ? "" : (s.length() > 300 ? s.substring(0, 300) + "..." : s); }
        List<Issue> searchIssues(String jql) throws Exception {
            String fields = String.join(",", CF_DEV, CF_QA, CF_TOTAL, "assignee", "status", "created", "resolutiondate", "updated", "issuetype", "labels", "summary");
            List<Issue> result = new ArrayList<>();
            String token = null;
            do {
                String url = base + "/rest/api/3/search/jql?jql=" + enc(jql) + "&fields=" + enc(fields) + "&maxResults=100"
                        + (token != null ? "&nextPageToken=" + enc(token) : "");
                Object root = Json.parse(getJson(url));
                List<Object> arr = (List<Object>) Json.get(root, "issues");
                if (arr != null) for (Object o : arr) result.add(toIssue(o));
                token = (String) Json.get(root, "nextPageToken");
            } while (token != null);
            return result;
        }
        List<Change> changelog(String key) throws Exception {
            List<Change> out = new ArrayList<>();
            int startAt = 0, total = Integer.MAX_VALUE;
            while (startAt < total) {
                String url = base + "/rest/api/3/issue/" + key + "/changelog?startAt=" + startAt + "&maxResults=100";
                Object root = Json.parse(getJson(url));
                total = ((Number) Json.getOr(root, "total", 0.0)).intValue();
                List<Object> vals = (List<Object>) Json.get(root, "values");
                if (vals == null || vals.isEmpty()) break;
                for (Object h : vals) {
                    String ts = Json.getStr(h, "created");
                    Object au = Json.get(h, "author");
                    String author = au == null ? null : Json.getStr(au, "displayName");
                    List<Object> items = (List<Object>) Json.get(h, "items");
                    if (items == null) continue;
                    for (Object it : items) {
                        String field = Json.getStr(it, "field");
                        if (!"status".equals(field) && !"assignee".equals(field)) continue;
                        Change c = new Change();
                        c.key = key; c.field = field; c.ts = parseTs(ts); c.author = author;
                        c.fromStr = Json.getStr(it, "fromString"); c.toStr = Json.getStr(it, "toString");
                        if (c.ts != null) out.add(c);
                    }
                }
                startAt += 100;
            }
            return out;
        }
        Issue toIssue(Object o) {
            Issue is = new Issue();
            is.key = Json.getStr(o, "key");
            Object f = Json.get(o, "fields");
            is.summary = Json.getStr(f, "summary");
            Object lb = Json.get(f, "labels");
            if (lb instanceof List) for (Object l : (List<Object>) lb) if (l != null) is.labels.add(l.toString());
            is.devSp = Json.getNum(f, CF_DEV); is.qaSp = Json.getNum(f, CF_QA); is.totalSp = Json.getNum(f, CF_TOTAL);
            is.created = Json.getStr(f, "created"); is.resolved = Json.getStr(f, "resolutiondate"); is.updated = Json.getStr(f, "updated");
            Object as = Json.get(f, "assignee"); is.currentAssignee = as == null ? null : Json.getStr(as, "displayName");
            Object st = Json.get(f, "status"); is.currentStatus = st == null ? null : Json.getStr(st, "name");
            Object ty = Json.get(f, "issuetype"); is.issueType = ty == null ? null : Json.getStr(ty, "name");
            return is;
        }
    }
    static String enc(String s) { return URLEncoder.encode(s, StandardCharsets.UTF_8); }

    // ============================ minimal JSON parser ============================
    static final class Json {
        private final String s; private int i;
        private Json(String s) { this.s = s; }
        static Object parse(String s) { Json j = new Json(s); j.ws(); return j.value(); }
        private void ws() { while (i < s.length() && Character.isWhitespace(s.charAt(i))) i++; }
        private Object value() {
            ws(); char c = s.charAt(i);
            switch (c) {
                case '{': return obj();
                case '[': return arr();
                case '"': return str();
                case 't': i += 4; return Boolean.TRUE;
                case 'f': i += 5; return Boolean.FALSE;
                case 'n': i += 4; return null;
                default: return num();
            }
        }
        private Map<String, Object> obj() {
            Map<String, Object> m = new LinkedHashMap<>(); i++; ws();
            if (s.charAt(i) == '}') { i++; return m; }
            while (true) {
                ws(); String k = str(); ws(); i++; /* : */ Object v = value(); m.put(k, v);
                ws(); char c = s.charAt(i++); if (c == '}') break;
            }
            return m;
        }
        private List<Object> arr() {
            List<Object> a = new ArrayList<>(); i++; ws();
            if (s.charAt(i) == ']') { i++; return a; }
            while (true) { a.add(value()); ws(); char c = s.charAt(i++); if (c == ']') break; }
            return a;
        }
        private String str() {
            StringBuilder b = new StringBuilder(); i++; /* opening quote */
            while (true) {
                char c = s.charAt(i++);
                if (c == '"') break;
                if (c == '\\') {
                    char e = s.charAt(i++);
                    switch (e) {
                        case '"': b.append('"'); break;
                        case '\\': b.append('\\'); break;
                        case '/': b.append('/'); break;
                        case 'b': b.append('\b'); break;
                        case 'f': b.append('\f'); break;
                        case 'n': b.append('\n'); break;
                        case 'r': b.append('\r'); break;
                        case 't': b.append('\t'); break;
                        case 'u': b.append((char) Integer.parseInt(s.substring(i, i + 4), 16)); i += 4; break;
                    }
                } else b.append(c);
            }
            return b.toString();
        }
        private Object num() {
            int st = i;
            while (i < s.length() && "+-0123456789.eE".indexOf(s.charAt(i)) >= 0) i++;
            return Double.parseDouble(s.substring(st, i));
        }
        // accessors
        static Object get(Object o, String k) { return o == null ? null : ((Map<String, Object>) o).get(k); }
        static Object getOr(Object o, String k, Object def) { Object v = get(o, k); return v == null ? def : v; }
        static String getStr(Object o, String k) { Object v = get(o, k); return v == null ? null : v.toString(); }
        static Double getNum(Object o, String k) { Object v = get(o, k); return v == null ? null : ((Number) v).doubleValue(); }
    }

    // ============================ minimal XLSX writer ============================
    static final class Sheet {
        String name; List<List<Object[]>> rows = new ArrayList<>(); // cell = {value, styleIndex}
        Sheet(String name) { this.name = name; }
        void row(Object... cells) { // each cell is value; default style 0; headers use rowStyled
            List<Object[]> r = new ArrayList<>();
            for (Object c : cells) r.add(new Object[]{c, 0});
            rows.add(r);
        }
        void header(Object... cells) {
            List<Object[]> r = new ArrayList<>();
            for (Object c : cells) r.add(new Object[]{c, 1});
            rows.add(r);
        }
    }

    static void writeWorkbook(String out, List<Detail> detail, List<Daily> daily, List<Meta> metas, Map<String, List<Change>> logs,
                             TreeMap<String, String> audit, String project, String label, String targetMonth, String base,
                             String rvoDefault, Set<String> eligible, String eligibleFile, LocalDateTime winStart, LocalDateTime winEnd) throws Exception {
        List<Sheet> sheets = new ArrayList<>();
        Map<String, Meta> metaByKey = new HashMap<>();
        for (Meta m : metas) metaByKey.put(m.issue, m);

        // Summary: person x month, split into COUNTED (eligible) and NOT COUNTED so the totals reconcile
        TreeSet<String> months = new TreeSet<>(), people = new TreeSet<>();
        Map<String, Double> grid = new HashMap<>();
        for (Detail d : detail) { months.add(d.month); people.add(d.person); grid.merge(d.person + "\u0001" + d.month, d.hours, Double::sum); }
        Sheet sum = new Sheet("Summary");
        List<Object> head = new ArrayList<>(); head.add("Person"); head.add("Eligible"); head.addAll(months); head.add("Total");
        sum.header(head.toArray());
        double[] totals = new double[2];
        for (int block = 0; block < 2; block++) {
            boolean counted = block == 0;
            sum.row(counted ? (eligible == null ? "COUNTED - no eligibility list configured, every named assignee (pass --eligible <file>)" : "COUNTED - on the eligibility list")
                            : "NOT COUNTED - not on the list, or unassigned");
            for (String p : people) {
                if (isEligible(eligible, p) != counted) continue;
                List<Object> r = new ArrayList<>(); r.add(p); r.add(eligLabel(eligible, p)); double tot = 0;
                for (String mo : months) { double v = grid.getOrDefault(p + "\u0001" + mo, 0.0); r.add(v == 0 ? null : round(v, 1)); tot += v; }
                r.add(round(tot, 1)); sum.row(r.toArray()); totals[block] += tot;
            }
            List<Object> tr = new ArrayList<>(); tr.add(counted ? "TOTAL counted" : "TOTAL not counted"); tr.add("");
            for (String mo : months) { double c = 0; for (String p : people) if (isEligible(eligible, p) == counted) c += grid.getOrDefault(p + "\u0001" + mo, 0.0); tr.add(round(c, 1)); }
            tr.add(round(totals[block], 1)); sum.header(tr.toArray());
            sum.row();
        }
        List<Object> all = new ArrayList<>(); all.add("TOTAL all assignees"); all.add("");
        for (String mo : months) { double c = 0; for (String p : people) c += grid.getOrDefault(p + "\u0001" + mo, 0.0); all.add(round(c, 1)); }
        all.add(round(totals[0] + totals[1], 1)); sum.header(all.toArray());
        sheets.add(sum);

        // Daily: the per-person / per-project / per-day record (working days only)
        Sheet dy = new Sheet("Daily");
        dy.header("Date", "Weekday", "Person", "Eligible", "RVO project", "Issue", "Summary", "Phase", "Hours", "Method", "Flag");
        daily.sort(Comparator.comparing((Daily d) -> d.date).thenComparing(d -> d.person).thenComparing(d -> d.issue).thenComparing(d -> d.phase));
        Map<String, Double> perPersonDay = new HashMap<>();
        for (Daily d : daily) perPersonDay.merge(d.person + "\u0001" + d.date, d.hours, Double::sum);
        for (Daily d : daily) {
            if (d.hours <= 0) continue;
            double dayTot = perPersonDay.get(d.person + "\u0001" + d.date);
            Meta mm = metaByKey.get(d.issue);
            dy.row(d.date.toString(), d.date.getDayOfWeek().toString().substring(0, 3), d.person, eligLabel(eligible, d.person), d.rvo, d.issue,
                   mm == null ? "" : nz(mm.summary), d.phase, round(d.hours, 3), d.method,
                   dayTot > 8.0 ? "> 8 h on this day (" + round(dayTot, 1) + " h in total)" : "");
        }
        sheets.add(dy);

        // DayTotals: person x day of the month (the classic hours sheet)
        Sheet dt = new Sheet("DayTotals");
        LocalDate first = winStart.toLocalDate(), last = winEnd.toLocalDate().minusDays(1);
        List<Object> dh = new ArrayList<>(); dh.add("Person"); dh.add("Eligible"); dh.add("RVO project");
        for (LocalDate d = first; !d.isAfter(last); d = d.plusDays(1)) dh.add(d.getDayOfMonth());
        dh.add("Total"); dt.header(dh.toArray());
        TreeMap<String, Map<LocalDate, Double>> pd = new TreeMap<>();
        Map<String, Set<String>> rvoByPerson = new HashMap<>();
        for (Daily d : daily) {
            pd.computeIfAbsent(d.person, k -> new HashMap<>()).merge(d.date, d.hours, Double::sum);
            rvoByPerson.computeIfAbsent(d.person, k -> new TreeSet<>()).add(d.rvo);
        }
        for (Map.Entry<String, Map<LocalDate, Double>> e : pd.entrySet()) {
            List<Object> r = new ArrayList<>(); r.add(e.getKey()); r.add(eligLabel(eligible, e.getKey())); r.add(String.join(", ", rvoByPerson.get(e.getKey()))); double tot = 0;
            for (LocalDate d = first; !d.isAfter(last); d = d.plusDays(1)) { double v = e.getValue().getOrDefault(d, 0.0); r.add(v == 0 ? null : round(v, 2)); tot += v; }
            r.add(round(tot, 1)); dt.row(r.toArray());
        }
        sheets.add(dt);

        // Evidence: every status/assignee change inside the month, with author and link (the progress trail)
        Sheet ev = new Sheet("Evidence");
        ev.header("Issue", "Summary", "RVO project", "Timestamp", "Author", "Field", "From", "To", "Link");
        List<Change> inMonth = new ArrayList<>();
        for (List<Change> cl : logs.values()) for (Change c : cl) if (c.ts != null && !c.ts.isBefore(winStart) && c.ts.isBefore(winEnd)) inMonth.add(c);
        inMonth.sort(Comparator.comparing((Change c) -> c.key).thenComparing(c -> c.ts));
        for (Change c : inMonth) {
            Meta mm = metaByKey.get(c.key);
            ev.row(c.key, mm == null ? "" : nz(mm.summary), mm == null ? rvoDefault : mm.rvo, c.ts.toString().replace('T', ' '), nz(c.author),
                   c.field, nz(c.fromStr), nz(c.toStr), base + "/browse/" + c.key);
        }
        sheets.add(ev);

        // Detail
        Sheet det = new Sheet("Detail");
        det.header("Issue", "Person", "Eligible", "RVO project", "Month", "Phase", "Hours", "Time held (h, working days)", "Method");
        detail.sort(Comparator.comparing((Detail d) -> d.issue).thenComparing(d -> d.month).thenComparing(d -> d.person).thenComparing(d -> d.phase));
        for (Detail d : detail) det.row(d.issue, d.person, eligLabel(eligible, d.person), d.rvo, d.month, d.phase, round(d.hours, 3), round(d.timeHeld, 2), d.method);
        sheets.add(det);

        // Issues
        Sheet iss = new Sheet("Issues");
        iss.header("Issue", "Summary", "RVO project", "Type", "Current status", "DEV SP", "QA SP", "Total SP", "DEV budget (h)", "QA budget (h)", "DEV allocated", "QA allocated", "Method", "Note", "Link");
        metas.sort(Comparator.comparing(x -> x.issue));
        for (Meta m : metas) iss.row(m.issue, nz(m.summary), m.rvo, m.issueType, m.status, m.devSp, m.qaSp, m.totalSp, m.devBudget, m.qaBudget, m.devAlloc, m.qaAlloc, m.method, m.note, base + "/browse/" + m.issue);
        sheets.add(iss);

        // StatusAudit
        Sheet sa = new Sheet("StatusAudit");
        sa.header("Status seen (migrated suffix stripped)", "Classified as");
        for (Map.Entry<String, String> e : audit.entrySet()) sa.row(e.getKey(), e.getValue() == null ? "— (no phase, 0h)" : e.getValue());
        sheets.add(sa);

        // Notes
        Sheet no = new Sheet("Notes");
        no.header("WBSO hours substantiation — methodology", "");
        no.row("Generated", LocalDateTime.now().toString());
        no.row("Project", project); no.row("Label", label); no.row("Reporting month", targetMonth);
        no.row("RVO project", "Default " + rvoDefault + " (--rvo-project). An issue label 'wbso-<id>' overrides it per issue. Daily, DayTotals, Detail, Issues and Evidence carry the number.");
        no.row("Eligibility", eligible == null
                ? "NO list configured: every named assignee is counted. Before anything goes to SubVice, pass --eligible <file> with one Jira display name per line (Validaide B.V. payroll only)."
                : "List " + eligibleFile + " (" + eligible.size() + " names). Only listed people are counted; others and unassigned time sit in the NOT COUNTED block of Summary so totals still reconcile.");
        no.row("Hours per story point", HOURS_PER_POINT);
        no.row("Base of truth", "Story Points DEV + QA (separate). Total used only when both are empty.");
        no.row("Per-person split", "Role-based, time-weighted: DEV hours to DEV-phase holders, QA to QA-phase holders, proportional to time held on working days (Mon-Fri). Weekend time is ignored in numerator and denominator.");
        no.row("Per-day split", "Each person's hours on an issue are spread over the working days the ticket was held in that phase, proportional to the time held per day. Daily rows sum to the Detail rows; DayTotals is Daily pivoted per person and day.");
        no.row("Month attribution", "Each month gets the share of the issue's point-budget equal to the phase-time inside that month divided by the issue's whole-life phase-time. Months are mutually consistent and sum to the issue total.");
        no.row("DEV-phase statuses", String.join(", ", new TreeSet<>(DEV_STATUSES)));
        no.row("QA-phase statuses", String.join(", ", new TreeSet<>(QA_STATUSES)));
        no.row("Nature of figures", "Derived ESTIMATE: story points x 4, allocated by recorded status history. Not a log of actual hours. Every Summary figure traces through Detail and Daily to an issue/person/project/day/phase.");
        no.row("Evidence tab", "Every status and assignee change inside the month, with timestamp, author and a link to the issue: the progress trail behind the hours.");
        no.row("Daily flag", "A person with more than 8 h on one day is flagged in Daily. It usually means a large ticket was held briefly; review before filing.");
        no.row("Fallback rows", "Issues with story points but NO DEV/QA-phase changelog timeline are attributed to the current assignee, on the day they were last updated/resolved (flagged in Issues), so they land once across monthly runs. Pre-refinement tickets get zero.");
        no.row("Unassigned", "Work-phase time while unassigned is bucketed as '(unassigned)' and never counted, so totals reconcile and the gap is visible.");
        sheets.add(no);

        zipXlsx(out, sheets);
    }

    static void zipXlsx(String out, List<Sheet> sheets) throws Exception {
        try (ZipOutputStream z = new ZipOutputStream(Files.newOutputStream(Path.of(out)))) {
            put(z, "[Content_Types].xml", contentTypes(sheets.size()));
            put(z, "_rels/.rels",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>");
            put(z, "xl/workbook.xml", workbookXml(sheets));
            put(z, "xl/_rels/workbook.xml.rels", workbookRels(sheets.size()));
            put(z, "xl/styles.xml", STYLES);
            for (int i = 0; i < sheets.size(); i++) put(z, "xl/worksheets/sheet" + (i + 1) + ".xml", sheetXml(sheets.get(i)));
        }
    }
    static void put(ZipOutputStream z, String name, String content) throws Exception {
        z.putNextEntry(new ZipEntry(name)); z.write(content.getBytes(StandardCharsets.UTF_8)); z.closeEntry();
    }
    static String contentTypes(int n) {
        StringBuilder b = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
            "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
            "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
            "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>" +
            "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>");
        for (int i = 1; i <= n; i++) b.append("<Override PartName=\"/xl/worksheets/sheet").append(i)
            .append(".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>");
        return b.append("</Types>").toString();
    }
    static String workbookXml(List<Sheet> sheets) {
        StringBuilder b = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" " +
            "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets>");
        for (int i = 0; i < sheets.size(); i++) b.append("<sheet name=\"").append(xml(sheets.get(i).name))
            .append("\" sheetId=\"").append(i + 1).append("\" r:id=\"rId").append(i + 1).append("\"/>");
        return b.append("</sheets></workbook>").toString();
    }
    static String workbookRels(int n) {
        StringBuilder b = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
        for (int i = 1; i <= n; i++) b.append("<Relationship Id=\"rId").append(i)
            .append("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet").append(i).append(".xml\"/>");
        b.append("<Relationship Id=\"rId").append(n + 1)
            .append("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>");
        return b.append("</Relationships>").toString();
    }
    static String sheetXml(Sheet sh) {
        StringBuilder b = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>");
        for (int r = 0; r < sh.rows.size(); r++) {
            List<Object[]> row = sh.rows.get(r);
            b.append("<row r=\"").append(r + 1).append("\">");
            for (int c = 0; c < row.size(); c++) {
                Object val = row.get(c)[0]; int style = (Integer) row.get(c)[1];
                String ref = colLetter(c) + (r + 1);
                if (val == null) { if (style != 0) b.append("<c r=\"").append(ref).append("\" s=\"").append(style).append("\"/>"); continue; }
                if (val instanceof Number) {
                    int st = style == 0 ? 2 : style; // numbers use 2-dp style unless header
                    b.append("<c r=\"").append(ref).append("\" s=\"").append(st).append("\"><v>").append(((Number) val).doubleValue()).append("</v></c>");
                } else {
                    b.append("<c r=\"").append(ref).append("\" s=\"").append(style).append("\" t=\"inlineStr\"><is><t xml:space=\"preserve\">")
                     .append(xml(val.toString())).append("</t></is></c>");
                }
            }
            b.append("</row>");
        }
        return b.append("</sheetData></worksheet>").toString();
    }
    static String colLetter(int n) {
        StringBuilder b = new StringBuilder(); n++;
        while (n > 0) { int rem = (n - 1) % 26; b.insert(0, (char) ('A' + rem)); n = (n - 1) / 26; }
        return b.toString();
    }
    static String xml(String s) {
        return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;");
    }
    static final String STYLES =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
        "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
        "<numFmts count=\"1\"><numFmt numFmtId=\"164\" formatCode=\"0.00\"/></numFmts>" +
        "<fonts count=\"2\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font>" +
        "<font><b/><sz val=\"11\"/><color rgb=\"FFFFFFFF\"/><name val=\"Calibri\"/></font></fonts>" +
        "<fills count=\"3\"><fill><patternFill patternType=\"none\"/></fill>" +
        "<fill><patternFill patternType=\"gray125\"/></fill>" +
        "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FF1F3864\"/><bgColor indexed=\"64\"/></patternFill></fill></fills>" +
        "<borders count=\"1\"><border/></borders>" +
        "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>" +
        "<cellXfs count=\"3\">" +
        "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>" +
        "<xf numFmtId=\"0\" fontId=\"1\" fillId=\"2\" borderId=\"0\" xfId=\"0\" applyFont=\"1\" applyFill=\"1\"/>" +
        "<xf numFmtId=\"164\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/>" +
        "</cellXfs>" +
        "<cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles></styleSheet>";

    // ============================ synthetic self-test data ============================
    static Object[] synthetic() {
        List<Issue> issues = new ArrayList<>();
        Map<String, List<Change>> logs = new LinkedHashMap<>();

        // VALCM-1: 10 DEV SP, 5 QA SP. Dev by Alice (Jan), reassigned to Bob for review (Feb), QA by Carol (Feb).
        Issue a = new Issue(); a.key = "VALCM-1"; a.summary = "Bulk lane upload"; a.issueType = "Story"; a.devSp = 10.0; a.qaSp = 5.0; a.labels.add("wbso");
        a.created = "2026-01-05T09:00:00.000+0000"; a.currentStatus = "Done"; a.currentAssignee = "Carol";
        a.resolved = "2026-02-20T17:00:00.000+0000"; a.updated = "2026-02-20T17:00:00.000+0000";
        issues.add(a);
        logs.put("VALCM-1", Arrays.asList(
            ch("VALCM-1", "2026-01-10T09:00:00.000+0000", "status", "To Do", "Under Development"),
            ch("VALCM-1", "2026-01-10T09:00:00.000+0000", "assignee", null, "Alice"),
            ch("VALCM-1", "2026-02-02T09:00:00.000+0000", "status", "Under Development", "Under review"),
            ch("VALCM-1", "2026-02-02T09:00:00.000+0000", "assignee", "Alice", "Bob"),
            ch("VALCM-1", "2026-02-10T09:00:00.000+0000", "status", "Under review", "Under Testing"),
            ch("VALCM-1", "2026-02-10T09:00:00.000+0000", "assignee", "Bob", "Carol"),
            ch("VALCM-1", "2026-02-20T17:00:00.000+0000", "status", "Under Testing", "Done")
        ));

        // VALCM-2: only total SP = 4 (fallback). Dev by Dan entirely in Jan.
        Issue b = new Issue(); b.key = "VALCM-2"; b.summary = "Weather cache expiry"; b.issueType = "Task"; b.totalSp = 4.0; b.labels.add("wbso"); b.labels.add("wbso-2602");
        b.created = "2026-01-15T09:00:00.000+0000"; b.currentStatus = "Done"; b.currentAssignee = "Dan";
        b.resolved = "2026-01-25T12:00:00.000+0000"; b.updated = "2026-01-25T12:00:00.000+0000";
        issues.add(b);
        logs.put("VALCM-2", Arrays.asList(
            ch("VALCM-2", "2026-01-16T09:00:00.000+0000", "status", "To Do", "In Development"),
            ch("VALCM-2", "2026-01-16T09:00:00.000+0000", "assignee", null, "Dan"),
            ch("VALCM-2", "2026-01-25T12:00:00.000+0000", "status", "In Development", "Done")
        ));

        // VALCM-3: work happened entirely in 2025 (before clip) -> should contribute 0.
        Issue c = new Issue(); c.key = "VALCM-3"; c.summary = "Old work from 2025"; c.issueType = "Story"; c.devSp = 8.0; c.labels.add("wbso");
        c.created = "2025-11-01T09:00:00.000+0000"; c.currentStatus = "Done"; c.currentAssignee = "Eve";
        c.resolved = "2025-12-15T09:00:00.000+0000"; c.updated = "2025-12-15T09:00:00.000+0000";
        issues.add(c);
        logs.put("VALCM-3", Arrays.asList(
            ch("VALCM-3", "2025-11-05T09:00:00.000+0000", "status", "To Do", "Under Development"),
            ch("VALCM-3", "2025-11-05T09:00:00.000+0000", "assignee", null, "Eve"),
            ch("VALCM-3", "2025-12-15T09:00:00.000+0000", "status", "Under Development", "Done")
        ));

        return new Object[]{issues, logs};
    }
    static Change ch(String key, String ts, String field, String from, String to) {
        Change c = new Change(); c.key = key; c.ts = parseTs(ts); c.field = field; c.fromStr = from; c.toStr = to; c.author = "synthetic"; return c;
    }
}

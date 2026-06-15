import java.net.URI;
import java.net.URLEncoder;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDateTime;
import java.time.YearMonth;
import java.time.Duration;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;
import java.io.ByteArrayOutputStream;
import java.io.OutputStream;

/**
 * WBSO hours substantiation for Jira project VALCM — dependency-free (JDK only).
 *
 * Pulls every issue labelled "wbso", reconstructs each issue's status + assignee
 * history from the full changelog, converts story points to an estimate of hours
 * worked per person per month (from a start month), and writes one .xlsx with
 * tabs: Summary, Detail, Issues, StatusAudit, Notes.
 *
 * Methodology (see Notes tab in the output):
 *   - 1 story point = 4 hours.
 *   - Base = Story Points DEV (customfield_12867) + QA (customfield_12904), each x4.
 *     The total field (customfield_10025) is used only when both DEV and QA are empty.
 *   - DEV hours spread over whoever held the ticket during DEV-phase statuses, QA hours
 *     over QA-phase statuses; time-weighted; months from transition timestamps; clipped
 *     to the start month. This is a derived estimate, not logged time.
 *
 * Build:  javac WbsoHours.java
 * Run:    export JIRA_EMAIL="you@company.com"
 *         export JIRA_API_TOKEN="xxxxx"
 *         java WbsoHours --start 2026-01 --out wbso_hours.xlsx
 *         (optional: --project VALCM  --label wbso  --base https://skycell.atlassian.net)
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
        String key, issueType, currentStatus, currentAssignee, created, resolved, updated;
        Double devSp, qaSp, totalSp;
    }
    static final class Change {       // one status/assignee change from the changelog
        String key, field, fromStr, toStr; LocalDateTime ts;
    }
    static final class Detail {
        String issue, person, month, phase, method; double hours, timeHeld;
        Detail(String i, String p, String m, String ph, double h, double t, String me) {
            issue = i; person = p; month = m; phase = ph; hours = h; timeHeld = t; method = me;
        }
    }
    static final class Meta {
        String issue, issueType, status, method = "", note = "";
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
        String startM  = opt.getOrDefault("start", "2026-01");
        String out     = opt.getOrDefault("out", "wbso_hours.xlsx");
        String base    = opt.getOrDefault("base", env("JIRA_BASE_URL", "https://skycell.atlassian.net"));
        LocalDateTime clip = parseStartMonth(startM);

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
        List<Meta> metas = new ArrayList<>();
        TreeMap<String, String> statusAudit = new TreeMap<>();
        for (Issue is : issues) allocate(is, logs.getOrDefault(is.key, Collections.emptyList()), clip, detail, metas, statusAudit);

        double total = detail.stream().mapToDouble(d -> d.hours).sum();
        Set<String> people = new TreeSet<>();
        for (Detail d : detail) people.add(d.person);
        System.out.printf("Grand total: %.1f h across %d people%n", total, people.size());

        writeWorkbook(out, detail, metas, statusAudit, project, label, startM);
        System.out.println("Wrote " + out);
    }

    static String env(String k, String def) { String v = System.getenv(k); return v == null ? def : v; }

    // ============================ allocation ============================
    static void allocate(Issue is, List<Change> changes, LocalDateTime clip,
                         List<Detail> detail, List<Meta> metas, Map<String, String> audit) {
        Meta m = new Meta();
        m.issue = is.key; m.issueType = is.issueType; m.status = is.currentStatus;
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

        // split changes by field, sorted
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

        // weights per (person|month|phase) for this issue
        Map<String, Double> wDev = new LinkedHashMap<>(), wQa = new LinkedHashMap<>();
        double sumDev = 0, sumQa = 0;

        List<LocalDateTime[]> dummy = new ArrayList<>(); // not used; intervals processed inline
        List<Object[]> intervals = new ArrayList<>();     // {start, end, status, assignee}
        for (Change c : all) {
            if (c.ts.isAfter(lastTs)) intervals.add(new Object[]{lastTs, c.ts, curS, curA});
            if (c.field.equals("status")) curS = c.toStr; else curA = c.toStr;
            lastTs = c.ts;
        }
        if (endTime.isAfter(lastTs)) intervals.add(new Object[]{lastTs, endTime, curS, curA});

        boolean hadWorkPhase = false;  // any DEV/QA interval at all, even before the clip date
        for (Object[] iv : intervals) {
            String st = (String) iv[2];
            if (st != null) { String s = norm(st); audit.put(s, classify(st)); }
            String phase = classify(st);
            if (phase == null) continue;
            hadWorkPhase = true;
            LocalDateTime segStart = (LocalDateTime) iv[0], segEnd = (LocalDateTime) iv[1];
            if (segStart.isBefore(clip)) segStart = clip;
            if (!segEnd.isAfter(segStart)) continue;
            String person = iv[3] == null ? "(unassigned)" : (String) iv[3];
            for (Object[] ms : monthSlices(segStart, segEnd)) {
                String month = (String) ms[0]; double secs = (Double) ms[1];
                if (secs <= 0) continue;
                String k = person + "\u0001" + month;
                if (phase.equals("DEV")) { wDev.merge(k, secs, Double::sum); sumDev += secs; }
                else { wQa.merge(k, secs, Double::sum); sumQa += secs; }
            }
        }

        if (baseMode.equals("DEV+QA")) {
            distribute(is.key, wDev, sumDev, devBudget, "DEV", "time-weighted", detail, m);
            distribute(is.key, wQa, sumQa, qaBudget, "QA", "time-weighted", detail, m);
            m.method = "time-weighted (DEV+QA)";
        } else { // TOTAL fallback: pool DEV+QA weights, one budget
            Map<String, Double> pooled = new LinkedHashMap<>();
            wDev.forEach((k, v) -> pooled.merge(k, v, Double::sum));
            wQa.forEach((k, v) -> pooled.merge(k, v, Double::sum));
            double sumAll = sumDev + sumQa;
            distribute(is.key, pooled, sumAll, devBudget, "DEV", "time-weighted (total pool)", detail, m);
            m.method = "time-weighted (total fallback)";
        }

        // fallback: budget but nothing allocated in range
        double allocated = m.devAlloc + m.qaAlloc, budgetTotal = devBudget + qaBudget;
        if (allocated < 0.01 && budgetTotal > 0) {
            LocalDateTime attrib = is.updated != null ? parseTs(is.updated) : endTime;
            if (hadWorkPhase) {
                // there was DEV/QA-phase activity, but all of it before the start month -> counts as 0
                m.method += " (pre-window)";
                m.note = "Work-phase time was entirely before the start month; counts as 0.";
            } else if (attrib != null && attrib.isBefore(clip)) {
                m.method = "zero (activity before start month)";
                m.note = "Issue last updated before the start month; counts as 0.";
            } else {
                boolean reached = classify(is.currentStatus) != null || is.resolved != null;
                if (!reached) {
                    m.method = "zero (no work phase reached)"; m.note = "No DEV/QA-phase status yet; not yet worked.";
                } else {
                    String person = is.currentAssignee == null ? "(unassigned)" : is.currentAssignee;
                    String month = monthKey(attrib != null ? attrib : endTime);
                    if (devBudget > 0) { detail.add(new Detail(is.key, person, month, "DEV", round(devBudget, 3), 0, "FALLBACK current-assignee")); m.devAlloc += devBudget; }
                    if (qaBudget > 0) { detail.add(new Detail(is.key, person, month, "QA", round(qaBudget, 3), 0, "FALLBACK current-assignee")); m.qaAlloc += qaBudget; }
                    m.method += " + FALLBACK";
                    m.note = "No changelog work-phase intervals; attributed to current assignee in last-update month.";
                }
            }
        }
        m.devAlloc = round(m.devAlloc, 2); m.qaAlloc = round(m.qaAlloc, 2);
        metas.add(m);
    }

    static void distribute(String key, Map<String, Double> weights, double sum, double budget,
                          String phase, String tag, List<Detail> detail, Meta m) {
        if (budget <= 0 || sum <= 0) return;
        for (Map.Entry<String, Double> e : weights.entrySet()) {
            String[] pk = e.getKey().split("\u0001", 2);
            double secs = e.getValue();
            double hrs = budget * secs / sum;
            if (hrs <= 0) continue;
            detail.add(new Detail(key, pk[0], pk[1], phase, round(hrs, 3), round(secs / 3600.0, 2), tag));
            if (phase.equals("DEV")) m.devAlloc += hrs; else m.qaAlloc += hrs;
        }
    }

    // ============================ time helpers ============================
    static final DateTimeFormatter TSF = DateTimeFormatter.ofPattern("yyyy-MM-dd'T'HH:mm:ss");
    static LocalDateTime parseTs(String s) {
        if (s == null) return null;
        try { return LocalDateTime.parse(s.substring(0, 19), TSF); } catch (Exception e) { return null; }
    }
    static LocalDateTime parseStartMonth(String s) {
        int y = Integer.parseInt(s.substring(0, 4)), mo = Integer.parseInt(s.substring(5, 7));
        return LocalDateTime.of(y, mo, 1, 0, 0);
    }
    static String monthKey(LocalDateTime dt) { return YearMonth.from(dt).toString(); }
    static List<Object[]> monthSlices(LocalDateTime a, LocalDateTime b) {
        List<Object[]> out = new ArrayList<>();
        LocalDateTime cur = a;
        while (cur.isBefore(b)) {
            LocalDateTime nextMonth = YearMonth.from(cur).plusMonths(1).atDay(1).atStartOfDay();
            LocalDateTime segEnd = nextMonth.isBefore(b) ? nextMonth : b;
            out.add(new Object[]{YearMonth.from(cur).toString(), (double) Duration.between(cur, segEnd).getSeconds()});
            cur = segEnd;
        }
        return out;
    }
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
        final String base, auth; final HttpClient http = HttpClient.newHttpClient();
        Jira(String base, String email, String token) {
            this.base = base.replaceAll("/+$", "");
            this.auth = "Basic " + Base64.getEncoder().encodeToString((email + ":" + token).getBytes(StandardCharsets.UTF_8));
        }
        String getJson(String url) throws Exception {
            for (int attempt = 0; attempt < 5; attempt++) {
                HttpRequest req = HttpRequest.newBuilder(URI.create(url))
                        .header("Authorization", auth).header("Accept", "application/json").GET().build();
                HttpResponse<String> r = http.send(req, HttpResponse.BodyHandlers.ofString());
                if (r.statusCode() == 429) { Thread.sleep(5000); continue; }
                if (r.statusCode() >= 500) { Thread.sleep(2000L * (attempt + 1)); continue; }
                if (r.statusCode() >= 400) throw new RuntimeException("HTTP " + r.statusCode() + ": " + r.body());
                return r.body();
            }
            throw new RuntimeException("Too many retries for " + url);
        }
        List<Issue> searchIssues(String jql) throws Exception {
            String fields = String.join(",", CF_DEV, CF_QA, CF_TOTAL, "assignee", "status", "created", "resolutiondate", "updated", "issuetype", "labels");
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
                    List<Object> items = (List<Object>) Json.get(h, "items");
                    if (items == null) continue;
                    for (Object it : items) {
                        String field = Json.getStr(it, "field");
                        if (!"status".equals(field) && !"assignee".equals(field)) continue;
                        Change c = new Change();
                        c.key = key; c.field = field; c.ts = parseTs(ts);
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

    static void writeWorkbook(String out, List<Detail> detail, List<Meta> metas,
                             TreeMap<String, String> audit, String project, String label, String startM) throws Exception {
        List<Sheet> sheets = new ArrayList<>();

        // Summary: person x month
        TreeSet<String> months = new TreeSet<>(), people = new TreeSet<>();
        Map<String, Double> grid = new HashMap<>();
        for (Detail d : detail) { months.add(d.month); people.add(d.person); grid.merge(d.person + "\u0001" + d.month, d.hours, Double::sum); }
        Sheet sum = new Sheet("Summary");
        List<Object> head = new ArrayList<>(); head.add("Person"); head.addAll(months); head.add("Total");
        sum.header(head.toArray());
        for (String p : people) {
            List<Object> r = new ArrayList<>(); r.add(p); double tot = 0;
            for (String mo : months) { double v = grid.getOrDefault(p + "\u0001" + mo, 0.0); r.add(v == 0 ? null : round(v, 1)); tot += v; }
            r.add(round(tot, 1)); sum.row(r.toArray());
        }
        List<Object> tr = new ArrayList<>(); tr.add("TOTAL"); double gt = 0;
        for (String mo : months) { double c = 0; for (String p : people) c += grid.getOrDefault(p + "\u0001" + mo, 0.0); tr.add(round(c, 1)); gt += c; }
        tr.add(round(gt, 1)); sum.header(tr.toArray());
        sheets.add(sum);

        // Detail
        Sheet det = new Sheet("Detail");
        det.header("Issue", "Person", "Month", "Phase", "Hours", "Time held (h)", "Method");
        detail.sort(Comparator.comparing((Detail d) -> d.issue).thenComparing(d -> d.month).thenComparing(d -> d.person).thenComparing(d -> d.phase));
        for (Detail d : detail) det.row(d.issue, d.person, d.month, d.phase, round(d.hours, 3), round(d.timeHeld, 2), d.method);
        sheets.add(det);

        // Issues
        Sheet iss = new Sheet("Issues");
        iss.header("Issue", "Type", "Current status", "DEV SP", "QA SP", "Total SP", "DEV budget (h)", "QA budget (h)", "DEV allocated", "QA allocated", "Method", "Note");
        metas.sort(Comparator.comparing(x -> x.issue));
        for (Meta m : metas) iss.row(m.issue, m.issueType, m.status, m.devSp, m.qaSp, m.totalSp, m.devBudget, m.qaBudget, m.devAlloc, m.qaAlloc, m.method, m.note);
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
        no.row("Project", project); no.row("Label", label); no.row("Counted from", startM);
        no.row("Hours per story point", HOURS_PER_POINT);
        no.row("Base of truth", "Story Points DEV + QA (separate). Total used only when both are empty.");
        no.row("Per-person split", "Role-based, time-weighted: DEV hours to DEV-phase holders, QA to QA-phase holders, proportional to time held.");
        no.row("Month attribution", "By status-transition timestamps, split across month boundaries, clipped to the start month.");
        no.row("DEV-phase statuses", String.join(", ", new TreeSet<>(DEV_STATUSES)));
        no.row("QA-phase statuses", String.join(", ", new TreeSet<>(QA_STATUSES)));
        no.row("Nature of figures", "Derived ESTIMATE: story points x 4, allocated by recorded status history. Not a log of actual hours. Every Summary figure traces through Detail to an issue/person/month/phase.");
        no.row("Fallback rows", "Issues with no DEV/QA-phase time after the start date are attributed to the current assignee in the last-update month (flagged in Issues); pre-refinement tickets get zero.");
        no.row("Unassigned", "Work-phase time while unassigned is bucketed as '(unassigned)' so totals reconcile.");
        no.row("Dutch-employee filter", "NOT applied; all assignees included. Filter downstream.");
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
        Issue a = new Issue(); a.key = "VALCM-1"; a.issueType = "Story"; a.devSp = 10.0; a.qaSp = 5.0;
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
        Issue b = new Issue(); b.key = "VALCM-2"; b.issueType = "Task"; b.totalSp = 4.0;
        b.created = "2026-01-15T09:00:00.000+0000"; b.currentStatus = "Done"; b.currentAssignee = "Dan";
        b.resolved = "2026-01-25T12:00:00.000+0000"; b.updated = "2026-01-25T12:00:00.000+0000";
        issues.add(b);
        logs.put("VALCM-2", Arrays.asList(
            ch("VALCM-2", "2026-01-16T09:00:00.000+0000", "status", "To Do", "In Development"),
            ch("VALCM-2", "2026-01-16T09:00:00.000+0000", "assignee", null, "Dan"),
            ch("VALCM-2", "2026-01-25T12:00:00.000+0000", "status", "In Development", "Done")
        ));

        // VALCM-3: work happened entirely in 2025 (before clip) -> should contribute 0.
        Issue c = new Issue(); c.key = "VALCM-3"; c.issueType = "Story"; c.devSp = 8.0;
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
        Change c = new Change(); c.key = key; c.ts = parseTs(ts); c.field = field; c.fromStr = from; c.toStr = to; return c;
    }
}

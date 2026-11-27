"""VJ-TimesheetTracker core: scans git repos, stores commits/pushes, exports Excel.

Your data lives in %APPDATA%\\VJ-TimesheetTracker (override with VJTT_HOME),
never in the program folder:
  config.json  - projects, start date, daily schedule time
  data.db      - SQLite store of every commit and push seen (kept forever,
                 so push history survives git's ~90-day reflog expiry)
  logs/        - daily update logs
"""
import json
import os
import sqlite3
import subprocess
import sys
import datetime as dt
from collections import OrderedDict, defaultdict

APP_NAME = "VJ-TimesheetTracker"
CODE_DIR = os.path.dirname(os.path.abspath(__file__))
HOME = os.environ.get("VJTT_HOME") or os.path.join(
    os.environ.get("APPDATA") or os.path.expanduser("~"), APP_NAME)
os.makedirs(HOME, exist_ok=True)
CONFIG = os.path.join(HOME, "config.json")
DB = os.path.join(HOME, "data.db")
LOGS = os.path.join(HOME, "logs")
TASK_NAME = f"{APP_NAME} Daily Update"
SKIP_DIRS = {"node_modules", "vendor", ".git", "dist", "build", "storage", ".venv", "venv"}
NO_WINDOW = 0x08000000 if os.name == "nt" else 0

DEFAULT_CONFIG = {"since": f"{dt.date.today().year - 1}-01-01", "daily_time": "19:00", "projects": []}


# ---------------------------------------------------------------- config
def load_config():
    if not os.path.exists(CONFIG):
        save_config(DEFAULT_CONFIG)
    with open(CONFIG, encoding="utf-8") as f:
        cfg = json.load(f)
    for k, v in DEFAULT_CONFIG.items():
        cfg.setdefault(k, v)
    return cfg


def save_config(cfg):
    with open(CONFIG, "w", encoding="utf-8", newline="\n") as f:
        json.dump(cfg, f, indent=2)


def add_project(name, path, emails=None):
    cfg = load_config()
    path = os.path.normpath(path)
    if any(os.path.normcase(p["path"]) == os.path.normcase(path) for p in cfg["projects"]):
        raise ValueError("This folder is already added.")
    if any(p["name"].lower() == name.lower() for p in cfg["projects"]):
        raise ValueError("A project with this name already exists.")
    repos = find_repos(path)
    if not repos:
        raise ValueError("No git repository found in this folder (or 2 levels below it).")
    if not emails:
        emails = detect_emails(repos)
    cfg["projects"].append({"name": name, "path": path, "emails": emails})
    save_config(cfg)
    return repos, emails


def remove_project(name, delete_data=False):
    cfg = load_config()
    cfg["projects"] = [p for p in cfg["projects"] if p["name"] != name]
    save_config(cfg)
    if delete_data:
        with db() as c:
            c.execute("DELETE FROM commits WHERE project=?", (name,))
            c.execute("DELETE FROM pushes WHERE project=?", (name,))


def set_since(date_iso):
    """Change how far back commits are collected. Earlier dates need a rescan, later ones prune."""
    dt.date.fromisoformat(date_iso)
    cfg = load_config()
    cfg["since"] = date_iso
    save_config(cfg)
    with db() as c:
        c.execute("DELETE FROM commits WHERE date < ?", (date_iso,))


def update_project_emails(name, emails):
    cfg = load_config()
    for p in cfg["projects"]:
        if p["name"] == name:
            p["emails"] = emails
    save_config(cfg)


# ---------------------------------------------------------------- git
def git(repo, *args):
    r = subprocess.run(["git", "-C", repo, *args], capture_output=True, text=True,
                       encoding="utf-8", errors="replace", creationflags=NO_WINDOW,
                       env={**os.environ, "GIT_TERMINAL_PROMPT": "0"})
    return r.stdout if r.returncode == 0 else ""


def find_repos(path, depth=2):
    """The folder itself if it's a repo, plus nested repos up to `depth` levels."""
    found = []

    def walk(p, d):
        if os.path.exists(os.path.join(p, ".git")):
            found.append(p)
        if d == 0:
            return
        try:
            entries = os.scandir(p)
        except OSError:
            return
        for e in entries:
            if e.is_dir() and e.name not in SKIP_DIRS and not e.name.startswith("."):
                walk(e.path, d - 1)

    walk(os.path.normpath(path), depth)
    return found


def detect_emails(repos):
    emails = []
    for r in repos:
        e = git(r, "config", "user.email").strip().lower()
        if e and e not in emails:
            emails.append(e)
    return emails


def scan_repo(repo, since, emails):
    sep = "\x1f"
    out = git(repo, "log", "--all", "--no-merges", f"--since={since}",
              "--date=iso-strict", f"--format=%H{sep}%h{sep}%ae{sep}%ad{sep}%s%x1e")
    commits = []
    for rec in out.split("\x1e"):
        rec = rec.strip("\r\n")
        if not rec:
            continue
        parts = rec.split(sep, 4)
        if len(parts) < 5:
            continue
        full, short, email, date, subject = parts
        if emails and email.lower() not in emails:
            continue
        t = dt.datetime.fromisoformat(date)
        commits.append((full, short, email.lower(), t.strftime("%Y-%m-%d"), t.strftime("%H:%M"),
                        " ".join(subject.split())))
    pushes = []
    for ref in git(repo, "for-each-ref", "--format=%(refname)", "refs/remotes").split():
        log = git(repo, "reflog", "show", "--date=iso-strict", f"--format=%H{sep}%gd{sep}%gs", ref)
        for line in log.splitlines():
            p = line.split(sep)
            if len(p) == 3 and "update by push" in p[2] and "@{" in p[1]:
                when = dt.datetime.fromisoformat(p[1].split("@{", 1)[1].rstrip("}"))
                pushes.append((p[0][:10], ref.replace("refs/remotes/", ""),
                               when.strftime("%Y-%m-%d"), when.strftime("%H:%M")))
    return commits, pushes


# ---------------------------------------------------------------- store
def db():
    c = sqlite3.connect(DB)
    c.execute("""CREATE TABLE IF NOT EXISTS commits(project TEXT, repo TEXT, hash TEXT, short TEXT,
                 email TEXT, date TEXT, time TEXT, subject TEXT, PRIMARY KEY(repo, hash))""")
    c.execute("""CREATE TABLE IF NOT EXISTS pushes(project TEXT, repo TEXT, branch TEXT, hash TEXT,
                 date TEXT, time TEXT, PRIMARY KEY(repo, branch, date, time, hash))""")
    c.execute("CREATE TABLE IF NOT EXISTS meta(k TEXT PRIMARY KEY, v TEXT)")
    return c


def update_all(log=print):
    cfg = load_config()
    totals = {}
    with db() as c:
        for p in cfg["projects"]:
            emails = [e.lower() for e in p.get("emails", [])]
            n_c = n_p = 0
            if not os.path.isdir(p["path"]):
                log(f"[{p['name']}] folder missing: {p['path']}")
                continue
            for repo in find_repos(p["path"]):
                rel = os.path.relpath(repo, p["path"])
                rel = os.path.basename(p["path"]) if rel == "." else rel.replace("\\", "/")
                commits, pushes = scan_repo(repo, cfg["since"], emails)
                c.executemany("INSERT OR IGNORE INTO commits VALUES(?,?,?,?,?,?,?,?)",
                              [(p["name"], rel, *x) for x in commits])
                c.executemany("INSERT OR IGNORE INTO pushes VALUES(?,?,?,?,?,?)",
                              [(p["name"], rel, x[1], x[0], x[2], x[3]) for x in pushes])
                n_c += len(commits)
                n_p += len(pushes)
            totals[p["name"]] = (n_c, n_p)
            log(f"[{p['name']}] {n_c} commits, {n_p} pushes found")
        c.execute("INSERT OR REPLACE INTO meta VALUES('last_update', ?)",
                  (dt.datetime.now().strftime("%Y-%m-%d %H:%M"),))
    return totals


def last_update():
    with db() as c:
        r = c.execute("SELECT v FROM meta WHERE k='last_update'").fetchone()
    return r[0] if r else "never"


def daily_rows(date_from, date_to, project=None):
    """One row per (date, project): times, counts, repos, unique subjects."""
    q = "SELECT date, time, project, repo, subject FROM commits WHERE date BETWEEN ? AND ?"
    args = [date_from, date_to]
    if project:
        q += " AND project=?"
        args.append(project)
    groups = OrderedDict()
    with db() as c:
        for d, t, proj, repo, subj in c.execute(q + " ORDER BY date, project, time", args):
            g = groups.setdefault((d, proj), {"times": [], "repos": set(), "subjects": []})
            g["times"].append(t)
            g["repos"].add(repo)
            if subj and subj not in g["subjects"]:
                g["subjects"].append(subj)
    rows = []
    for (d, proj), g in groups.items():
        first, last = min(g["times"]), max(g["times"])
        span = (dt.datetime.strptime(last, "%H:%M") - dt.datetime.strptime(first, "%H:%M")).seconds / 3600
        rows.append({"date": d, "project": proj, "first": first, "last": last,
                     "span": round(span * 2) / 2, "commits": len(g["times"]),
                     "repos": ", ".join(sorted(g["repos"])), "subjects": g["subjects"]})
    return rows


# ---------------------------------------------------------------- export
def export_excel(path, date_from, date_to, project=None):
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill
    from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE as IL

    clean = lambda s: IL.sub("", s or "")
    wb = Workbook()
    hdr, fill = Font(bold=True, color="FFFFFF"), PatternFill("solid", fgColor="162F6A")

    def sheet(ws, cols, widths):
        ws.append(cols)
        for i, cell in enumerate(ws[1]):
            cell.font, cell.fill = hdr, fill
            ws.column_dimensions[cell.column_letter].width = widths[i]
        ws.freeze_panes = "A2"

    ws = wb.active
    ws.title = "Timesheet"
    sheet(ws, ["Date", "Day", "Project", "First commit", "Last commit", "Commit span (h)",
               "Hours worked", "Commits", "Repos", "Work done"], [13, 6, 16, 12, 12, 14, 13, 9, 30, 110])
    rows = daily_rows(date_from, date_to, project)
    for r in rows:
        d = dt.date.fromisoformat(r["date"])
        ws.append([d, d.strftime("%a"), r["project"], r["first"], r["last"], r["span"], None,
                   r["commits"], r["repos"], clean("\n".join("• " + s[:200] for s in r["subjects"]))])
        ws.cell(ws.max_row, 1).number_format = "DD-MMM-YYYY"
        for col in range(1, 11):
            ws.cell(ws.max_row, col).alignment = Alignment(vertical="top", wrap_text=(col == 10))
        ws.cell(ws.max_row, 7).fill = PatternFill("solid", fgColor="FFF4CC")

    m = wb.create_sheet("Monthly Summary")
    sheet(m, ["Month", "Project", "Active days", "Commits"], [12, 16, 12, 10])
    agg = defaultdict(lambda: [0, 0])
    for r in rows:
        a = agg[(r["date"][:7], r["project"])]
        a[0] += 1
        a[1] += r["commits"]
    for (mon, proj), (days, n) in sorted(agg.items()):
        m.append([dt.datetime.strptime(mon, "%Y-%m").strftime("%b %Y"), proj, days, n])
    m.append(["Total", "", sum(v[0] for v in agg.values()), sum(v[1] for v in agg.values())])
    m.cell(m.max_row, 1).font = Font(bold=True)

    q_extra, args = "", [date_from, date_to]
    if project:
        q_extra, args = " AND project=?", args + [project]
    with db() as c:
        a = wb.create_sheet("All Commits")
        sheet(a, ["Date", "Time", "Project", "Repo", "Commit", "Message"], [13, 8, 16, 24, 11, 130])
        for d, t, proj, repo, short, subj in c.execute(
                "SELECT date, time, project, repo, short, subject FROM commits WHERE date BETWEEN ? AND ?"
                + q_extra + " ORDER BY date, time", args):
            a.append([dt.date.fromisoformat(d), t, proj, repo, short, clean(subj)])
            a.cell(a.max_row, 1).number_format = "DD-MMM-YYYY"
        pz = wb.create_sheet("Pushes")
        sheet(pz, ["Date", "Time", "Project", "Repo", "Branch", "Commit"], [13, 8, 16, 24, 30, 11])
        for d, t, proj, repo, br, h in c.execute(
                "SELECT date, time, project, repo, branch, hash FROM pushes WHERE date BETWEEN ? AND ?"
                + q_extra + " ORDER BY date, time", args):
            pz.append([dt.date.fromisoformat(d), t, proj, repo, br, h])
            pz.cell(pz.max_row, 1).number_format = "DD-MMM-YYYY"
    wb.save(path)
    return len(rows)


# ---------------------------------------------------------------- schedule
def set_schedule(hhmm):
    """(Re)create the daily Windows scheduled task. Runs missed jobs at next logon/boot."""
    dt.datetime.strptime(hhmm, "%H:%M")
    pyw = os.path.join(os.path.dirname(os.path.realpath(sys.executable)), "pythonw.exe")
    if not os.path.exists(pyw):
        pyw = sys.executable
    script = os.path.join(CODE_DIR, "daily_update.py")
    q = lambda s: s.replace("'", "''")
    ps = (
        f"$a = New-ScheduledTaskAction -Execute '{q(pyw)}' -Argument '\"{q(script)}\"' "
        f"-WorkingDirectory '{q(CODE_DIR)}';"
        f"$t = New-ScheduledTaskTrigger -Daily -At '{hhmm}';"
        "$s = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries "
        "-DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Minutes 30);"
        f"Register-ScheduledTask -TaskName '{TASK_NAME}' -Action $a -Trigger $t -Settings $s "
        "-Description 'Collects git commits/pushes for timesheet' -Force | Out-Null"
    )
    r = subprocess.run(["powershell", "-NoProfile", "-Command", ps], capture_output=True, text=True,
                       creationflags=NO_WINDOW)
    if r.returncode != 0:
        raise RuntimeError(r.stderr.strip() or "Could not create scheduled task")
    cfg = load_config()
    cfg["daily_time"] = hhmm
    save_config(cfg)


def remove_schedule():
    subprocess.run(["powershell", "-NoProfile", "-Command",
                    f"Unregister-ScheduledTask -TaskName '{TASK_NAME}' -Confirm:$false -ErrorAction SilentlyContinue"],
                   capture_output=True, creationflags=NO_WINDOW)


# ---------------------------------------------------------------- CLI
if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(prog="tracker_core.py", description=f"{APP_NAME} command line")
    sub = ap.add_subparsers(dest="cmd", required=True)
    a = sub.add_parser("add", help="track a project folder")
    a.add_argument("path")
    a.add_argument("--name")
    a.add_argument("--email", action="append", help="your git email (repeatable); default: auto-detect")
    sub.add_parser("list", help="list tracked projects")
    sub.add_parser("update", help="rescan all projects now")
    e = sub.add_parser("export", help="write an Excel timesheet")
    e.add_argument("out")
    e.add_argument("--from", dest="date_from", required=True)
    e.add_argument("--to", dest="date_to", default=dt.date.today().isoformat())
    e.add_argument("--project")
    s = sub.add_parser("schedule", help="set the daily update time (HH:MM)")
    s.add_argument("time")
    sub.add_parser("unschedule", help="remove the daily update task")
    args = ap.parse_args()
    if args.cmd == "add":
        name = args.name or os.path.basename(os.path.normpath(args.path))
        repos, emails = add_project(name, args.path, [x.lower() for x in args.email or []])
        print(f"Added '{name}': {len(repos)} repo(s), emails: {', '.join(emails) or '(everyone)'}")
    elif args.cmd == "list":
        for p in load_config()["projects"]:
            print(f"{p['name']}\t{p['path']}\t{', '.join(p.get('emails', []))}")
    elif args.cmd == "update":
        update_all()
    elif args.cmd == "export":
        print(f"{export_excel(args.out, args.date_from, args.date_to, args.project)} day-rows -> {args.out}")
    elif args.cmd == "schedule":
        set_schedule(args.time)
        print(f"Daily update at {args.time}")
    elif args.cmd == "unschedule":
        remove_schedule()

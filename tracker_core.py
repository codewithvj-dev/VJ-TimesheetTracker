"""VJ-TimesheetTracker core: collects your commits/pushes, stores them, exports Excel.

Sources
  local   - a folder on disk; every git repo in it (and 2 levels below) is scanned
  github  - repos on GitHub read through a connected account (no clone needed)

Git accounts
  One or more identities (work, personal, ...). Each has the emails you commit
  with and can optionally be connected to GitHub with a personal access token.
  The token is encrypted with Windows DPAPI, so only your Windows user can read it.

Your data lives in %APPDATA%\\VJ-TimesheetTracker (override with VJTT_HOME),
never in the program folder:
  config.json  - accounts, projects, start date, daily schedule time
  data.db      - SQLite store of every commit and push seen (kept forever,
                 so push history survives git's ~90-day reflog expiry)
  logs/        - daily update logs
"""
import base64
import ctypes
import json
import os
import re
import sqlite3
import subprocess
import sys
import uuid
import urllib.error
import urllib.parse
import urllib.request
import datetime as dt
from collections import OrderedDict, defaultdict

APP_NAME = "VJ-TimesheetTracker"
__version__ = "1.1.0"
REPO_URL = "https://github.com/codewithvj-dev/VJ-TimesheetTracker"
FROZEN = getattr(sys, "frozen", False)  # True when running as the packaged .exe
CODE_DIR = os.path.dirname(sys.executable) if FROZEN else os.path.dirname(os.path.abspath(__file__))
INSTALL_DIR = os.path.join(os.environ.get("LOCALAPPDATA") or os.path.expanduser("~"), "Programs", APP_NAME)
INSTALLED_EXE = os.path.join(INSTALL_DIR, f"{APP_NAME}.exe")
HOME = os.environ.get("VJTT_HOME") or os.path.join(
    os.environ.get("APPDATA") or os.path.expanduser("~"), APP_NAME)
os.makedirs(HOME, exist_ok=True)
CONFIG = os.path.join(HOME, "config.json")
DB = os.path.join(HOME, "data.db")
LOGS = os.path.join(HOME, "logs")
TASK_NAME = f"{APP_NAME} Daily Update"
GITHUB_API = "https://api.github.com"
SKIP_DIRS = {"node_modules", "vendor", ".git", "dist", "build", "storage", ".venv", "venv"}
NO_WINDOW = 0x08000000 if os.name == "nt" else 0

DEFAULT_CONFIG = {"since": f"{dt.date.today().year - 1}-01-01", "daily_time": "19:00",
                  "accounts": [], "projects": []}


# ---------------------------------------------------------------- config
def load_config():
    if not os.path.exists(CONFIG):
        save_config(json.loads(json.dumps(DEFAULT_CONFIG)))
    with open(CONFIG, encoding="utf-8") as f:
        cfg = json.load(f)
    changed = False
    for k, v in DEFAULT_CONFIG.items():
        if k not in cfg:
            cfg[k] = json.loads(json.dumps(v))
            changed = True
    for p in cfg["projects"]:
        if "type" not in p:
            p["type"] = "local"
            changed = True
    if not cfg.get("accounts_initialized"):
        # first run: create a default account from git's global identity + existing project emails
        emails = []
        g = git(None, "config", "--global", "user.email").strip().lower()
        for e in ([g] if g else []) + [e for p in cfg["projects"] for e in p.get("emails", [])]:
            if e and e not in emails:
                emails.append(e)
        if emails and not cfg["accounts"]:
            name = git(None, "config", "--global", "user.name").strip() or "Me"
            cfg["accounts"].append({"id": _new_id(), "label": name, "emails": emails, "github": None})
        cfg["accounts_initialized"] = True
        changed = True
    if changed:
        save_config(cfg)
    return cfg


def save_config(cfg):
    tmp = CONFIG + ".tmp"
    with open(tmp, "w", encoding="utf-8", newline="\n") as f:
        json.dump(cfg, f, indent=2)
    os.replace(tmp, CONFIG)


def _new_id():
    return uuid.uuid4().hex[:8]


def set_since(date_iso):
    """Change how far back commits are collected. Earlier dates need a rescan, later ones prune."""
    dt.date.fromisoformat(date_iso)
    cfg = load_config()
    cfg["since"] = date_iso
    save_config(cfg)
    with db() as c:
        c.execute("DELETE FROM commits WHERE date < ?", (date_iso,))


# ---------------------------------------------------------------- accounts
def get_account(account_id, cfg=None):
    cfg = cfg or load_config()
    for a in cfg["accounts"]:
        if a["id"] == account_id:
            return a
    raise ValueError("Account not found.")


def add_account(label, emails):
    cfg = load_config()
    label = label.strip()
    if not label:
        raise ValueError("Give the account a name, e.g. Work or Personal.")
    if any(a["label"].lower() == label.lower() for a in cfg["accounts"]):
        raise ValueError("An account with this name already exists.")
    acc = {"id": _new_id(), "label": label, "emails": _clean_emails(emails), "github": None}
    cfg["accounts"].append(acc)
    save_config(cfg)
    return acc


def update_account(account_id, label=None, emails=None):
    cfg = load_config()
    a = get_account(account_id, cfg)
    if label is not None and label.strip():
        a["label"] = label.strip()
    if emails is not None:
        a["emails"] = _clean_emails(emails)
    save_config(cfg)


def remove_account(account_id):
    cfg = load_config()
    users = [p["name"] for p in cfg["projects"] if p.get("account") == account_id]
    if users:
        raise ValueError("Remove these GitHub projects first: " + ", ".join(users))
    cfg["accounts"] = [a for a in cfg["accounts"] if a["id"] != account_id]
    save_config(cfg)


def all_account_emails(cfg=None):
    cfg = cfg or load_config()
    out = []
    for a in cfg["accounts"]:
        for e in a.get("emails", []):
            if e not in out:
                out.append(e)
    return out


def _clean_emails(emails):
    if isinstance(emails, str):
        emails = emails.replace(";", ",").split(",")
    out = []
    for e in emails or []:
        e = e.strip().lower()
        if e and e not in out:
            out.append(e)
    return out


# ---------------------------------------------------------------- token encryption (Windows DPAPI)
class _Blob(ctypes.Structure):
    _fields_ = [("cbData", ctypes.c_ulong), ("pbData", ctypes.POINTER(ctypes.c_char))]


def _dpapi(data, protect):
    if os.name != "nt":
        raise RuntimeError("Saving a GitHub token needs Windows (DPAPI).")
    crypt32, kernel32 = ctypes.windll.crypt32, ctypes.windll.kernel32
    kernel32.LocalFree.argtypes = [ctypes.c_void_p]
    buf = ctypes.create_string_buffer(data, len(data))
    ent_raw = APP_NAME.encode()
    ent_buf = ctypes.create_string_buffer(ent_raw, len(ent_raw))
    blob_in = _Blob(len(data), ctypes.cast(buf, ctypes.POINTER(ctypes.c_char)))
    blob_ent = _Blob(len(ent_raw), ctypes.cast(ent_buf, ctypes.POINTER(ctypes.c_char)))
    blob_out = _Blob()
    fn = crypt32.CryptProtectData if protect else crypt32.CryptUnprotectData
    if not fn(ctypes.byref(blob_in), None, ctypes.byref(blob_ent), None, None, 0x1, ctypes.byref(blob_out)):
        raise RuntimeError("Windows could not " + ("encrypt" if protect else "decrypt") + " the token.")
    try:
        return ctypes.string_at(blob_out.pbData, blob_out.cbData)
    finally:
        kernel32.LocalFree(ctypes.cast(blob_out.pbData, ctypes.c_void_p))


def encrypt_token(token):
    return base64.b64encode(_dpapi(token.encode(), True)).decode()


def decrypt_token(blob):
    return _dpapi(base64.b64decode(blob), False).decode()


# ---------------------------------------------------------------- GitHub API
class GitHubError(Exception):
    pass


def _gh(api, token, path, params=None, paginate=False, max_pages=50):
    url = api.rstrip("/") + path + ("?" + urllib.parse.urlencode(params) if params else "")
    headers = {"Accept": "application/vnd.github+json", "User-Agent": APP_NAME,
               "X-GitHub-Api-Version": "2022-11-28"}
    if token:
        headers["Authorization"] = f"Bearer {token}"
    items, pages = [], 0
    while url and pages < max_pages:
        req = urllib.request.Request(url, headers=headers)
        try:
            with urllib.request.urlopen(req, timeout=30) as r:
                data = json.loads(r.read().decode("utf-8"))
                link = r.headers.get("Link", "")
        except urllib.error.HTTPError as e:
            if e.code == 401:
                raise GitHubError("GitHub rejected the token (401). Reconnect the account with a new token.")
            if e.code == 403:
                raise GitHubError("GitHub refused access (403): rate limit, SSO authorisation "
                                  "or missing token permission.")
            if e.code == 404:
                raise GitHubError(f"Not found or no access (404): {path}")
            if e.code == 409:  # empty repository
                return []
            raise GitHubError(f"GitHub error {e.code} for {path}")
        except urllib.error.URLError as e:
            raise GitHubError(f"Cannot reach GitHub ({e.reason}). Check your internet connection.")
        if not paginate:
            return data
        items.extend(data)
        m = re.search(r'<([^>]+)>;\s*rel="next"', link)
        url = m.group(1) if m else None
        pages += 1
    return items


def connect_github(account_id, token, api=GITHUB_API):
    """Verify the token, remember the login (encrypted token) and add GitHub's noreply emails."""
    token = token.strip()
    if not token:
        raise ValueError("Paste a GitHub token.")
    me = _gh(api, token, "/user")
    cfg = load_config()
    a = get_account(account_id, cfg)
    for other in cfg["accounts"]:
        g = other.get("github")
        if other["id"] != account_id and g and g["login"].lower() == me["login"].lower() and g["api"] == api:
            raise ValueError(f"GitHub user {me['login']} is already connected to '{other['label']}'.")
    a["github"] = {"login": me["login"], "api": api.rstrip("/"), "token_enc": encrypt_token(token)}
    extra = [f"{me['id']}+{me['login']}@users.noreply.github.com".lower(),
             f"{me['login']}@users.noreply.github.com".lower()]
    if me.get("email"):
        extra.insert(0, me["email"].lower())
    a["emails"] = _clean_emails(a.get("emails", []) + extra)
    save_config(cfg)
    return me["login"]


def disconnect_github(account_id):
    cfg = load_config()
    get_account(account_id, cfg)["github"] = None
    save_config(cfg)


def _account_token(a):
    g = a.get("github")
    if not g:
        raise GitHubError(f"Account '{a['label']}' is not connected to GitHub.")
    return g["api"], decrypt_token(g["token_enc"])


def list_github_repos(account_id):
    a = get_account(account_id)
    api, token = _account_token(a)
    repos = _gh(api, token, "/user/repos", {"per_page": 100, "sort": "pushed",
                                             "affiliation": "owner,collaborator,organization_member"},
                paginate=True)
    return [{"full_name": r["full_name"], "private": r["private"], "pushed_at": (r.get("pushed_at") or "")[:10]}
            for r in repos]


def scan_github(project, since, account):
    api, token = _account_token(account)
    login = (account.get("github") or {}).get("login")
    if not login:
        raise GitHubError(f"Account '{account['label']}' is not connected to GitHub.")
    idents = [login] + [e for e in account.get("emails", []) if "@" in e]
    commits, pushes = {}, []
    for repo in project["repos"]:
        for who in idents:
            for it in _gh(api, token, f"/repos/{repo}/commits",
                          {"since": f"{since}T00:00:00Z", "author": who, "per_page": 100}, paginate=True):
                if len(it.get("parents", [])) > 1 or it["sha"] in commits:
                    continue
                c = it["commit"]
                when = dt.datetime.fromisoformat(c["author"]["date"].replace("Z", "+00:00")).astimezone()
                subject = " ".join((c.get("message") or "").splitlines()[0].split()) if c.get("message") else ""
                commits[it["sha"]] = (repo, it["sha"], it["sha"][:7], (c["author"].get("email") or "").lower(),
                                      when.strftime("%Y-%m-%d"), when.strftime("%H:%M"), subject)
    wanted = {r.lower() for r in project["repos"]}
    for ev in _gh(api, token, f"/users/{login}/events", {"per_page": 100}, paginate=True, max_pages=3):
        if ev.get("type") != "PushEvent" or ev["repo"]["name"].lower() not in wanted:
            continue
        when = dt.datetime.fromisoformat(ev["created_at"].replace("Z", "+00:00")).astimezone()
        branch = (ev["payload"].get("ref") or "").replace("refs/heads/", "")
        pushes.append((ev["repo"]["name"], branch, (ev["payload"].get("head") or "")[:10],
                       when.strftime("%Y-%m-%d"), when.strftime("%H:%M")))
    return list(commits.values()), pushes


# ---------------------------------------------------------------- projects
def add_project(name, path, emails=None):
    """Track a local folder. emails=None/[] means 'count commits from all my accounts'."""
    cfg = load_config()
    path = os.path.normpath(path)
    name = name.strip()
    if any(p.get("type") == "local" and os.path.normcase(p["path"]) == os.path.normcase(path)
           for p in cfg["projects"]):
        raise ValueError("This folder is already added.")
    _check_name(cfg, name)
    repos = find_repos(path)
    if not repos:
        raise ValueError("No git repository found in this folder (or 2 levels below it).")
    cfg["projects"].append({"name": name, "type": "local", "path": path, "emails": _clean_emails(emails)})
    save_config(cfg)
    return repos, project_emails(cfg["projects"][-1], cfg)


def add_github_project(name, account_id, repos):
    cfg = load_config()
    name = name.strip()
    _check_name(cfg, name)
    get_account(account_id, cfg)
    if not repos:
        raise ValueError("Pick at least one repository.")
    cfg["projects"].append({"name": name, "type": "github", "account": account_id, "repos": list(repos)})
    save_config(cfg)


def _check_name(cfg, name):
    if not name:
        raise ValueError("Give the project a name.")
    if any(p["name"].lower() == name.lower() for p in cfg["projects"]):
        raise ValueError("A project with this name already exists.")


def remove_project(name, delete_data=False):
    cfg = load_config()
    cfg["projects"] = [p for p in cfg["projects"] if p["name"] != name]
    save_config(cfg)
    if delete_data:
        with db() as c:
            c.execute("DELETE FROM commits WHERE project=?", (name,))
            c.execute("DELETE FROM pushes WHERE project=?", (name,))


def update_project_emails(name, emails):
    cfg = load_config()
    for p in cfg["projects"]:
        if p["name"] == name:
            p["emails"] = _clean_emails(emails)
    save_config(cfg)
    with db() as c:  # rescanned on next update with the new filter
        c.execute("DELETE FROM commits WHERE project=?", (name,))


def project_emails(p, cfg=None):
    """Emails whose commits count for a local project: its own list, else all accounts."""
    return p.get("emails") or all_account_emails(cfg)


def project_source(p, cfg=None):
    if p.get("type") == "github":
        try:
            label = get_account(p["account"], cfg)["label"]
        except ValueError:
            label = "?"
        return f"GitHub ({label}): " + ", ".join(p["repos"])
    return p["path"]


def project_counting(p, cfg=None):
    if p.get("type") == "github":
        try:
            a = get_account(p["account"], cfg)
            return f"{a['github']['login'] if a.get('github') else '(disconnected)'} + {len(a['emails'])} email(s)"
        except ValueError:
            return "(account removed)"
    return ", ".join(p["emails"]) if p.get("emails") else "All my accounts"


# ---------------------------------------------------------------- git (local)
def git(repo, *args):
    cmd = ["git"] + (["-C", repo] if repo else []) + list(args)
    try:
        r = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace",
                           creationflags=NO_WINDOW, env={**os.environ, "GIT_TERMINAL_PROMPT": "0"})
    except FileNotFoundError:
        return ""
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
            entries = list(os.scandir(p))
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
                 email TEXT, date TEXT, time TEXT, subject TEXT, PRIMARY KEY(project, repo, hash))""")
    c.execute("""CREATE TABLE IF NOT EXISTS pushes(project TEXT, repo TEXT, branch TEXT, hash TEXT,
                 date TEXT, time TEXT, PRIMARY KEY(project, repo, branch, date, time, hash))""")
    c.execute("CREATE TABLE IF NOT EXISTS meta(k TEXT PRIMARY KEY, v TEXT)")
    _migrate_keys(c)
    return c


def _migrate_keys(c):
    """v1 keyed rows by repo only; two projects with a same-named repo collided. Re-key by project."""
    if c.execute("SELECT v FROM meta WHERE k='schema'").fetchone():
        return
    for table, cols in (("commits", "project TEXT, repo TEXT, hash TEXT, short TEXT, email TEXT, date TEXT, "
                                    "time TEXT, subject TEXT, PRIMARY KEY(project, repo, hash)"),
                        ("pushes", "project TEXT, repo TEXT, branch TEXT, hash TEXT, date TEXT, time TEXT, "
                                   "PRIMARY KEY(project, repo, branch, date, time, hash)")):
        pk = [r[1] for r in c.execute(f"PRAGMA table_info({table})") if r[5]]
        if "project" not in pk:
            c.execute(f"ALTER TABLE {table} RENAME TO {table}_v1")
            c.execute(f"CREATE TABLE {table}({cols})")
            c.execute(f"INSERT OR IGNORE INTO {table} SELECT * FROM {table}_v1")
            c.execute(f"DROP TABLE {table}_v1")
    c.execute("INSERT OR REPLACE INTO meta VALUES('schema', '2')")
    c.commit()


def update_all(log=print):
    cfg = load_config()
    totals = {}
    with db() as c:
        for p in cfg["projects"]:
            try:
                if p.get("type") == "github":
                    commits, pushes = scan_github(p, cfg["since"], get_account(p["account"], cfg))
                else:
                    commits, pushes = _scan_local_project(p, cfg)
            except (GitHubError, ValueError, RuntimeError) as e:
                log(f"[{p['name']}] skipped: {e}")
                totals[p["name"]] = (0, 0, str(e))
                continue
            c.executemany("INSERT OR IGNORE INTO commits VALUES(?,?,?,?,?,?,?,?)",
                          [(p["name"], *x) for x in commits])
            c.executemany("INSERT OR IGNORE INTO pushes VALUES(?,?,?,?,?,?)",
                          [(p["name"], *x) for x in pushes])
            totals[p["name"]] = (len(commits), len(pushes), None)
            log(f"[{p['name']}] {len(commits)} commits, {len(pushes)} pushes found")
        c.execute("INSERT OR REPLACE INTO meta VALUES('last_update', ?)",
                  (dt.datetime.now().strftime("%Y-%m-%d %H:%M"),))
    return totals


def _scan_local_project(p, cfg):
    if not os.path.isdir(p["path"]):
        raise ValueError(f"folder missing: {p['path']}")
    repos = find_repos(p["path"])
    emails = project_emails(p, cfg) or detect_emails(repos)
    commits, pushes = [], []
    for repo in repos:
        rel = os.path.relpath(repo, p["path"])
        rel = os.path.basename(p["path"]) if rel == "." else rel.replace("\\", "/")
        cs, ps = scan_repo(repo, cfg["since"], emails)
        commits += [(rel, *x) for x in cs]
        pushes += [(rel, x[1], x[0], x[2], x[3]) for x in ps]
    return commits, pushes


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
def _ps(script):
    return subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", script],
                          capture_output=True, text=True, creationflags=NO_WINDOW)


def _q(s):
    return s.replace("'", "''")


def set_schedule(hhmm, exe=None):
    """(Re)create the daily Windows scheduled task. Runs missed jobs at next logon/boot."""
    dt.datetime.strptime(hhmm, "%H:%M")
    if FROZEN:
        execute = exe or sys.executable
        argument = "--daily-update"
        workdir = os.path.dirname(execute)
    else:
        execute = os.path.join(os.path.dirname(os.path.realpath(sys.executable)), "pythonw.exe")
        if not os.path.exists(execute):
            execute = sys.executable
        argument = f'"{os.path.join(CODE_DIR, "daily_update.py")}"'
        workdir = CODE_DIR
    ps = (
        f"$a = New-ScheduledTaskAction -Execute '{_q(execute)}' -Argument '{_q(argument)}' "
        f"-WorkingDirectory '{_q(workdir)}';"
        f"$t = New-ScheduledTaskTrigger -Daily -At '{hhmm}';"
        "$s = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries "
        "-DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Minutes 30);"
        f"Register-ScheduledTask -TaskName '{TASK_NAME}' -Action $a -Trigger $t -Settings $s "
        "-Description 'Collects git commits/pushes for timesheet' -Force | Out-Null"
    )
    r = _ps(ps)
    if r.returncode != 0:
        raise RuntimeError(r.stderr.strip() or "Could not create scheduled task")
    cfg = load_config()
    cfg["daily_time"] = hhmm
    save_config(cfg)


def remove_schedule():
    _ps(f"Unregister-ScheduledTask -TaskName '{TASK_NAME}' -Confirm:$false -ErrorAction SilentlyContinue")


def run_daily_update():
    """What the scheduled task runs: rescan everything, append to logs/YYYY-MM.log."""
    import traceback
    os.makedirs(LOGS, exist_ok=True)
    with open(os.path.join(LOGS, dt.date.today().strftime("%Y-%m") + ".log"), "a", encoding="utf-8") as f:
        stamp = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        try:
            update_all(log=lambda m: f.write(f"{stamp} {m}\n"))
            f.write(f"{stamp} update OK\n")
        except Exception:  # noqa: BLE001 - logged for the user
            f.write(f"{stamp} update FAILED\n{traceback.format_exc()}\n")


def git_available():
    try:
        return subprocess.run(["git", "--version"], capture_output=True, creationflags=NO_WINDOW).returncode == 0
    except OSError:
        return False


# ---------------------------------------------------------------- exe install / uninstall
_UNINSTALL_KEY = rf"Software\Microsoft\Windows\CurrentVersion\Uninstall\{APP_NAME}"


def _shortcut_paths():
    r = _ps("[Environment]::GetFolderPath('Desktop'); [Environment]::GetFolderPath('Programs')")
    return [os.path.join(p.strip(), f"{APP_NAME}.lnk") for p in r.stdout.splitlines() if p.strip()]


def is_installed_copy():
    return FROZEN and os.path.normcase(os.path.abspath(sys.executable)) == os.path.normcase(INSTALLED_EXE)


def install_self():
    """Packaged exe only: copy to %LOCALAPPDATA%\\Programs, add shortcuts, daily task, Apps & features entry."""
    import shutil
    import winreg
    if not FROZEN:
        raise RuntimeError("install_self() is for the packaged .exe; use install.ps1 for the source version.")
    os.makedirs(INSTALL_DIR, exist_ok=True)
    if not is_installed_copy():
        try:
            shutil.copy2(sys.executable, INSTALLED_EXE)
        except PermissionError:
            raise RuntimeError(f"{APP_NAME} is already running from {INSTALL_DIR}. Close it and try again.")
    for lnk in _shortcut_paths():
        r = _ps(f"$s=(New-Object -ComObject WScript.Shell).CreateShortcut('{_q(lnk)}');"
                f"$s.TargetPath='{_q(INSTALLED_EXE)}';$s.WorkingDirectory='{_q(INSTALL_DIR)}';"
                f"$s.IconLocation='{_q(INSTALLED_EXE)},0';$s.Description='Timesheet from your git commits';$s.Save()")
        if r.returncode != 0:
            raise RuntimeError(f"Could not create shortcut {lnk}: {r.stderr.strip()}")
    set_schedule(load_config()["daily_time"], exe=INSTALLED_EXE)
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, _UNINSTALL_KEY) as k:
        for name, val in [("DisplayName", APP_NAME), ("DisplayVersion", __version__), ("Publisher", "codewithvj"),
                          ("DisplayIcon", f"{INSTALLED_EXE},0"), ("InstallLocation", INSTALL_DIR),
                          ("URLInfoAbout", REPO_URL), ("UninstallString", f'"{INSTALLED_EXE}" --uninstall')]:
            winreg.SetValueEx(k, name, 0, winreg.REG_SZ, val)
        winreg.SetValueEx(k, "NoModify", 0, winreg.REG_DWORD, 1)
        winreg.SetValueEx(k, "NoRepair", 0, winreg.REG_DWORD, 1)
    return INSTALLED_EXE


def uninstall_self(remove_data=False):
    """Remove task, shortcuts, Apps & features entry; delete the program folder once this process exits."""
    import winreg
    remove_schedule()
    for lnk in _shortcut_paths():
        try:
            os.remove(lnk)
        except OSError:
            pass
    try:
        winreg.DeleteKey(winreg.HKEY_CURRENT_USER, _UNINSTALL_KEY)
    except OSError:
        pass
    targets = [INSTALL_DIR] + ([HOME] if remove_data else [])
    rm = " & ".join(f'rmdir /s /q "{t}"' for t in targets)
    subprocess.Popen(f'cmd /c timeout /t 3 /nobreak >nul & {rm}', shell=True, creationflags=NO_WINDOW,
                     cwd=os.environ.get("TEMP") or None)


# ---------------------------------------------------------------- CLI
def _cli():
    import argparse
    import getpass
    ap = argparse.ArgumentParser(prog="tracker_core.py", description=f"{APP_NAME} command line")
    sub = ap.add_subparsers(dest="cmd", required=True)

    sub.add_parser("accounts", help="list git accounts")
    x = sub.add_parser("account-add", help="add a git account (identity)")
    x.add_argument("label")
    x.add_argument("--email", action="append", required=True, help="email you commit with (repeatable)")
    x = sub.add_parser("account-remove", help="remove a git account")
    x.add_argument("account_id")
    x = sub.add_parser("github-connect", help="connect an account to GitHub (asks for a token)")
    x.add_argument("account_id")
    x.add_argument("--api", default=GITHUB_API, help="API URL, for GitHub Enterprise")
    x = sub.add_parser("github-disconnect", help="forget an account's GitHub token")
    x.add_argument("account_id")
    x = sub.add_parser("github-repos", help="list repos an account can read")
    x.add_argument("account_id")

    x = sub.add_parser("add", help="track a local project folder")
    x.add_argument("path")
    x.add_argument("--name")
    x.add_argument("--email", action="append", help="only count these emails (default: all my accounts)")
    x = sub.add_parser("add-github", help="track GitHub repos through a connected account")
    x.add_argument("name")
    x.add_argument("--account", required=True)
    x.add_argument("--repo", action="append", required=True, help="owner/name (repeatable)")
    sub.add_parser("list", help="list tracked projects")
    x = sub.add_parser("remove", help="stop tracking a project")
    x.add_argument("name")
    x.add_argument("--delete-data", action="store_true")
    sub.add_parser("update", help="rescan all projects now")
    x = sub.add_parser("export", help="write an Excel timesheet")
    x.add_argument("out")
    x.add_argument("--from", dest="date_from", required=True)
    x.add_argument("--to", dest="date_to", default=dt.date.today().isoformat())
    x.add_argument("--project")
    x = sub.add_parser("since", help="set the start date (YYYY-MM-DD)")
    x.add_argument("date")
    x = sub.add_parser("schedule", help="set the daily update time (HH:MM)")
    x.add_argument("time")
    sub.add_parser("unschedule", help="remove the daily update task")

    a = ap.parse_args()
    if a.cmd == "accounts":
        for acc in load_config()["accounts"]:
            gh = acc["github"]["login"] if acc.get("github") else "-"
            print(f"{acc['id']}\t{acc['label']}\tGitHub: {gh}\t{', '.join(acc['emails'])}")
    elif a.cmd == "account-add":
        acc = add_account(a.label, a.email)
        print(f"Added account {acc['id']} ({acc['label']})")
    elif a.cmd == "account-remove":
        remove_account(a.account_id)
    elif a.cmd == "github-connect":
        print(f"Connected as {connect_github(a.account_id, getpass.getpass('GitHub token: '), a.api)}")
    elif a.cmd == "github-disconnect":
        disconnect_github(a.account_id)
    elif a.cmd == "github-repos":
        for r in list_github_repos(a.account_id):
            print(f"{r['full_name']}\t{'private' if r['private'] else 'public'}\tlast push {r['pushed_at']}")
    elif a.cmd == "add":
        name = a.name or os.path.basename(os.path.normpath(a.path))
        repos, emails = add_project(name, a.path, a.email)
        print(f"Added '{name}': {len(repos)} repo(s), counting: {', '.join(emails) or '(everyone)'}")
    elif a.cmd == "add-github":
        add_github_project(a.name, a.account, a.repo)
        print(f"Added '{a.name}': {len(a.repo)} GitHub repo(s)")
    elif a.cmd == "list":
        cfg = load_config()
        for p in cfg["projects"]:
            print(f"{p['name']}\t{project_source(p, cfg)}\t{project_counting(p, cfg)}")
    elif a.cmd == "remove":
        remove_project(a.name, a.delete_data)
    elif a.cmd == "update":
        update_all()
    elif a.cmd == "export":
        print(f"{export_excel(a.out, a.date_from, a.date_to, a.project)} day-rows -> {a.out}")
    elif a.cmd == "since":
        set_since(a.date)
    elif a.cmd == "schedule":
        set_schedule(a.time)
        print(f"Daily update at {a.time}")
    elif a.cmd == "unschedule":
        remove_schedule()


if __name__ == "__main__":
    try:
        _cli()
    except (ValueError, GitHubError, RuntimeError) as e:
        sys.exit(f"Error: {e}")

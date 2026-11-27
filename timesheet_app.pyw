"""VJ-TimesheetTracker - desktop window.

Git accounts: one or many identities (work, personal...), optionally connected to GitHub.
Projects: local folders, or GitHub repos read through a connected account.
Update now: rescan everything (also runs daily via Task Scheduler).
Download: Excel timesheet for a date range and project.
"""
import os
import threading
import webbrowser
import datetime as dt
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

import tracker_core as core

NAVY = "#162F6A"
TOKEN_HELP_URL = "https://github.com/settings/personal-access-tokens/new"


def run_bg(widget, work, done, fail=None):
    """Run work() off the UI thread, then done(result) / fail(exc) on it."""
    def go():
        try:
            res = work()
            widget.after(0, lambda: done(res))
        except Exception as e:  # noqa: BLE001 - shown to the user
            widget.after(0, lambda: (fail or (lambda x: messagebox.showerror(core.APP_NAME, str(x),
                                                                              parent=widget)))(e))
    threading.Thread(target=go, daemon=True).start()


# ====================================================================== dialogs
class TokenDialog(simpledialog.Dialog):
    def __init__(self, parent, account_label):
        self.account_label = account_label
        self.result = None
        super().__init__(parent, f"Connect '{account_label}' to GitHub")

    def body(self, f):
        ttk.Label(f, justify="left", text=(
            "Paste a GitHub personal access token for this account.\n\n"
            "Recommended: a fine-grained token with Repository access = All repositories\n"
            "(or the ones you want) and read-only 'Contents' + 'Metadata' permissions.\n"
            "A classic token needs the 'repo' scope for private repos.\n\n"
            "The token is encrypted with Windows (DPAPI) - only your Windows user can read it.")
        ).grid(row=0, column=0, columnspan=2, sticky="w")
        ttk.Button(f, text="Create a token on GitHub…",
                   command=lambda: webbrowser.open(TOKEN_HELP_URL)).grid(row=1, column=0, sticky="w", pady=6)
        ttk.Label(f, text="Token").grid(row=2, column=0, sticky="w")
        self.e_token = ttk.Entry(f, show="•", width=60)
        self.e_token.grid(row=3, column=0, columnspan=2, sticky="we")
        ttk.Label(f, text="API URL (change only for GitHub Enterprise)").grid(row=4, column=0, sticky="w", pady=(8, 0))
        self.e_api = ttk.Entry(f, width=60)
        self.e_api.insert(0, core.GITHUB_API)
        self.e_api.grid(row=5, column=0, columnspan=2, sticky="we")
        return self.e_token

    def apply(self):
        self.result = (self.e_token.get().strip(), self.e_api.get().strip() or core.GITHUB_API)


class AccountsDialog(tk.Toplevel):
    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self.title("Git accounts")
        self.geometry("860x360")
        self.transient(app)
        ttk.Label(self, padding=(10, 10, 10, 0), justify="left", text=(
            "Add one account per identity you commit with (e.g. Work and Personal). Local projects count "
            "commits from all\nthese emails unless the project has its own list. Connect an account to "
            "GitHub to track repos you haven't cloned.")).pack(anchor="w")
        body = ttk.Frame(self, padding=10)
        body.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(body, columns=("github", "emails"))
        self.tree.heading("#0", text="Account")
        self.tree.heading("github", text="GitHub")
        self.tree.heading("emails", text="Emails")
        self.tree.column("#0", width=150)
        self.tree.column("github", width=140)
        self.tree.column("emails", width=380)
        self.tree.pack(side="left", fill="both", expand=True)
        b = ttk.Frame(body)
        b.pack(side="left", fill="y", padx=(8, 0))
        for text, cmd in [("Add account…", self.add), ("Edit…", self.edit), ("Remove", self.remove),
                          ("Connect GitHub…", self.connect), ("Disconnect GitHub", self.disconnect),
                          ("Close", self.destroy)]:
            ttk.Button(b, text=text, command=cmd, width=18).pack(pady=2)
        self.load()

    def load(self):
        self.tree.delete(*self.tree.get_children())
        for a in core.load_config()["accounts"]:
            gh = a["github"]["login"] if a.get("github") else "not connected"
            self.tree.insert("", "end", iid=a["id"], text=a["label"], values=(gh, ", ".join(a["emails"])))

    def selected(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo(core.APP_NAME, "Select an account first.", parent=self)
            return None
        return sel[0]

    def changed(self):
        self.load()
        self.app.load_projects()

    def add(self):
        label = simpledialog.askstring("Add account", "Account name (e.g. Work, Personal):", parent=self)
        if not label:
            return
        emails = simpledialog.askstring("Add account",
                                        "Emails you commit with on this account (comma-separated).\n"
                                        "Tip: run  git config user.email  in a repo to see it:", parent=self)
        if emails is None:
            return
        try:
            core.add_account(label, emails)
        except ValueError as e:
            messagebox.showerror("Add account", str(e), parent=self)
            return
        self.changed()
        self.app.update_now()

    def edit(self):
        aid = self.selected()
        if not aid:
            return
        a = core.get_account(aid)
        label = simpledialog.askstring("Edit account", "Account name:", initialvalue=a["label"], parent=self)
        if label is None:
            return
        emails = simpledialog.askstring("Edit account", "Emails (comma-separated):",
                                        initialvalue=", ".join(a["emails"]), parent=self)
        if emails is None:
            return
        core.update_account(aid, label, emails)
        self.changed()
        self.app.update_now()

    def remove(self):
        aid = self.selected()
        if not aid:
            return
        if not messagebox.askyesno("Remove account", "Remove this account? Its commits stop being counted "
                                   "for projects that use 'All my accounts'.", parent=self):
            return
        try:
            core.remove_account(aid)
        except ValueError as e:
            messagebox.showerror("Remove account", str(e), parent=self)
            return
        self.changed()

    def connect(self):
        aid = self.selected()
        if not aid:
            return
        d = TokenDialog(self, core.get_account(aid)["label"])
        if not d.result:
            return
        token, api = d.result
        self.config(cursor="watch")

        def done(login):
            self.config(cursor="")
            messagebox.showinfo("GitHub", f"Connected as {login}.\nGitHub's no-reply emails were added to this "
                                "account.\nNow use 'Add GitHub repos…' in the main window.", parent=self)
            self.changed()

        def fail(e):
            self.config(cursor="")
            messagebox.showerror("GitHub", str(e), parent=self)

        run_bg(self, lambda: core.connect_github(aid, token, api), done, fail)

    def disconnect(self):
        aid = self.selected()
        if aid and messagebox.askyesno("GitHub", "Forget this account's GitHub token? GitHub projects using "
                                       "it will be skipped until you reconnect.", parent=self):
            core.disconnect_github(aid)
            self.changed()


class GitHubReposDialog(tk.Toplevel):
    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self.title("Add GitHub repos")
        self.geometry("640x520")
        self.transient(app)
        self.repos = []
        cfg = core.load_config()
        self.accounts = [a for a in cfg["accounts"] if a.get("github")]
        f = ttk.Frame(self, padding=10)
        f.pack(fill="both", expand=True)
        if not self.accounts:
            ttk.Label(f, text="No account is connected to GitHub yet.\n"
                              "Open 'Git accounts…', select an account and click 'Connect GitHub…'.").pack()
            ttk.Button(f, text="Open Git accounts…", command=lambda: (self.destroy(), app.accounts())).pack(pady=8)
            return
        row = ttk.Frame(f)
        row.pack(fill="x")
        ttk.Label(row, text="Account").pack(side="left")
        self.v_acc = tk.StringVar(value=self._label(self.accounts[0]))
        box = ttk.Combobox(row, textvariable=self.v_acc, state="readonly", width=34,
                           values=[self._label(a) for a in self.accounts])
        box.pack(side="left", padx=6)
        box.bind("<<ComboboxSelected>>", lambda e: self.load())
        ttk.Label(row, text="Search").pack(side="left", padx=(10, 4))
        self.v_q = tk.StringVar()
        self.v_q.trace_add("write", lambda *_: self.fill())
        ttk.Entry(row, textvariable=self.v_q, width=20).pack(side="left")
        ttk.Label(f, text="Select one or more repos (Ctrl/Shift-click). They become one project.").pack(
            anchor="w", pady=(8, 2))
        lf = ttk.Frame(f)
        lf.pack(fill="both", expand=True)
        self.lb = tk.Listbox(lf, selectmode="extended", activestyle="none")
        sb = ttk.Scrollbar(lf, command=self.lb.yview)
        self.lb.configure(yscrollcommand=sb.set)
        self.lb.pack(side="left", fill="both", expand=True)
        sb.pack(side="left", fill="y")
        bottom = ttk.Frame(f)
        bottom.pack(fill="x", pady=(8, 0))
        ttk.Label(bottom, text="Project name").pack(side="left")
        self.v_name = tk.StringVar()
        ttk.Entry(bottom, textvariable=self.v_name, width=24).pack(side="left", padx=6)
        ttk.Button(bottom, text="Add project", command=self.add).pack(side="right")
        self.status = tk.StringVar()
        ttk.Label(f, textvariable=self.status).pack(anchor="w", pady=(4, 0))
        self.load()

    @staticmethod
    def _label(a):
        return f"{a['label']} ({a['github']['login']})"

    def account(self):
        return next(a for a in self.accounts if self._label(a) == self.v_acc.get())

    def load(self):
        self.status.set("Loading repositories from GitHub…")
        self.lb.delete(0, "end")

        def done(repos):
            self.repos = repos
            self.fill()
            self.status.set(f"{len(repos)} repositories")

        def fail(e):
            self.status.set("")
            messagebox.showerror("GitHub", str(e), parent=self)

        run_bg(self, lambda: core.list_github_repos(self.account()["id"]), done, fail)

    def fill(self):
        q = self.v_q.get().lower()
        self.shown = [r for r in self.repos if q in r["full_name"].lower()]
        self.lb.delete(0, "end")
        for r in self.shown:
            self.lb.insert("end", f"{r['full_name']}    {'private' if r['private'] else 'public'}"
                                  f"    last push {r['pushed_at'] or '-'}")

    def add(self):
        picked = [self.shown[i]["full_name"] for i in self.lb.curselection()]
        if not picked:
            messagebox.showinfo("Add GitHub repos", "Select at least one repository.", parent=self)
            return
        name = self.v_name.get().strip() or picked[0].split("/")[-1]
        try:
            core.add_github_project(name, self.account()["id"], picked)
        except ValueError as e:
            messagebox.showerror("Add GitHub repos", str(e), parent=self)
            return
        self.destroy()
        self.app.load_projects()
        self.app.update_now()


# ====================================================================== main window
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(core.APP_NAME)
        self.geometry("1150x720")
        self.minsize(950, 580)
        style = ttk.Style(self)
        try:
            style.theme_use("vista")
        except tk.TclError:
            pass
        style.configure("Treeview", rowheight=24)
        style.configure("H.TLabel", font=("Segoe UI", 11, "bold"), foreground=NAVY)

        # ---- projects
        top = ttk.Frame(self, padding=10)
        top.pack(fill="x")
        ttk.Label(top, text="Projects", style="H.TLabel").pack(anchor="w")
        pf = ttk.Frame(top)
        pf.pack(fill="x", pady=4)
        self.projects = ttk.Treeview(pf, columns=("source", "counting"), height=5)
        self.projects.heading("#0", text="Name")
        self.projects.heading("source", text="Source")
        self.projects.heading("counting", text="Counting commits by")
        self.projects.column("#0", width=150)
        self.projects.column("source", width=440)
        self.projects.column("counting", width=330)
        self.projects.pack(side="left", fill="x", expand=True)
        btns = ttk.Frame(pf)
        btns.pack(side="left", padx=8, fill="y")
        for text, cmd in [("Git accounts…", self.accounts), ("Add local project…", self.add_project),
                          ("Add GitHub repos…", self.add_github), ("Edit emails…", self.edit_emails),
                          ("Remove", self.remove_project), ("Update now", self.update_now)]:
            ttk.Button(btns, text=text, command=cmd, width=18).pack(pady=1)

        # ---- export / filter
        mid = ttk.Frame(self, padding=(10, 0))
        mid.pack(fill="x")
        ttk.Label(mid, text="Timesheet", style="H.TLabel").grid(row=0, column=0, sticky="w", columnspan=8)
        today = dt.date.today()
        self.v_from = tk.StringVar(value=today.replace(day=1).isoformat())
        self.v_to = tk.StringVar(value=today.isoformat())
        self.v_proj = tk.StringVar(value="All projects")
        ttk.Label(mid, text="From").grid(row=1, column=0, padx=(0, 4))
        ttk.Entry(mid, textvariable=self.v_from, width=12).grid(row=1, column=1)
        ttk.Label(mid, text="To").grid(row=1, column=2, padx=(10, 4))
        ttk.Entry(mid, textvariable=self.v_to, width=12).grid(row=1, column=3)
        ttk.Label(mid, text="Project").grid(row=1, column=4, padx=(10, 4))
        self.proj_box = ttk.Combobox(mid, textvariable=self.v_proj, state="readonly", width=20)
        self.proj_box.grid(row=1, column=5)
        ttk.Button(mid, text="Show", command=self.refresh_preview).grid(row=1, column=6, padx=8)
        ttk.Button(mid, text="Download Excel…", command=self.download).grid(row=1, column=7)
        quick = ttk.Frame(mid)
        quick.grid(row=2, column=0, columnspan=8, sticky="w", pady=4)
        for label, fn in [("This month", self.q_this_month), ("Last month", self.q_last_month),
                          ("This week", self.q_this_week), ("Everything", self.q_all)]:
            ttk.Button(quick, text=label, command=fn).pack(side="left", padx=(0, 4))

        # ---- preview
        bot = ttk.Frame(self, padding=10)
        bot.pack(fill="both", expand=True)
        cols = ("day", "project", "time", "span", "commits", "work")
        self.preview = ttk.Treeview(bot, columns=cols, show="tree headings")
        self.preview.heading("#0", text="Date")
        for c, t, w in [("day", "Day", 50), ("project", "Project", 130), ("time", "First – Last", 110),
                        ("span", "Span h", 60), ("commits", "Commits", 65), ("work", "Work done", 580)]:
            self.preview.heading(c, text=t)
            self.preview.column(c, width=w, anchor="w")
        self.preview.column("#0", width=100)
        sb = ttk.Scrollbar(bot, orient="vertical", command=self.preview.yview)
        self.preview.configure(yscrollcommand=sb.set)
        self.preview.pack(side="left", fill="both", expand=True)
        sb.pack(side="left", fill="y")

        # ---- status
        sf = ttk.Frame(self, padding=(10, 4))
        sf.pack(fill="x")
        self.status = tk.StringVar()
        ttk.Label(sf, textvariable=self.status).pack(side="left")
        ttk.Button(sf, text="Change daily time…", command=self.change_schedule).pack(side="right")
        ttk.Button(sf, text="Change start date…", command=self.change_since).pack(side="right", padx=6)
        ttk.Button(sf, text="Open data folder", command=lambda: os.startfile(core.HOME)).pack(side="right")

        self.load_projects()
        if not core.load_config()["projects"]:
            self.after(300, self.welcome)
        self.update_now()

    # ------------------------------------------------ helpers
    def set_status(self, extra=""):
        cfg = core.load_config()
        self.status.set(f"Last update: {core.last_update()}   ·   Daily auto-update at {cfg['daily_time']}"
                        f"   ·   Since {cfg['since']}" + (f"   ·   {extra}" if extra else ""))

    def load_projects(self):
        self.projects.delete(*self.projects.get_children())
        cfg = core.load_config()
        for p in cfg["projects"]:
            self.projects.insert("", "end", text=p["name"],
                                 values=(core.project_source(p, cfg), core.project_counting(p, cfg)))
        self.proj_box["values"] = ["All projects"] + [p["name"] for p in cfg["projects"]]
        if self.v_proj.get() not in self.proj_box["values"]:
            self.v_proj.set("All projects")
        self.set_status()

    def welcome(self):
        cfg = core.load_config()
        acc = cfg["accounts"]
        who = (f"Your git account '{acc[0]['label']}' ({', '.join(acc[0]['emails'])}) was detected."
               if acc else "No git identity was found - add one with 'Git accounts…'.")
        messagebox.showinfo(f"Welcome to {core.APP_NAME}",
                            f"{who}\n\n1. 'Git accounts…' - add more identities or connect GitHub.\n"
                            "2. 'Add local project…' or 'Add GitHub repos…'.\n"
                            "3. Pick dates and click 'Download Excel…'.", parent=self)

    def selected_project(self):
        sel = self.projects.selection()
        if not sel:
            messagebox.showinfo(core.APP_NAME, "Select a project first.")
            return None
        return self.projects.item(sel[0], "text")

    def range(self):
        try:
            f = dt.date.fromisoformat(self.v_from.get().strip())
            t = dt.date.fromisoformat(self.v_to.get().strip())
        except ValueError:
            messagebox.showerror(core.APP_NAME, "Dates must be YYYY-MM-DD, e.g. 2026-11-01")
            return None
        if f > t:
            messagebox.showerror(core.APP_NAME, "'From' is after 'To'.")
            return None
        proj = None if self.v_proj.get() == "All projects" else self.v_proj.get()
        return f.isoformat(), t.isoformat(), proj

    # ------------------------------------------------ project actions
    def accounts(self):
        AccountsDialog(self)

    def add_github(self):
        GitHubReposDialog(self)

    def add_project(self):
        path = filedialog.askdirectory(title="Choose the project folder (git repo or folder containing repos)")
        if not path:
            return
        name = simpledialog.askstring("Project name", "Name for this project:",
                                      initialvalue=os.path.basename(os.path.normpath(path)), parent=self)
        if not name:
            return
        try:
            repos, emails = core.add_project(name, path)
        except ValueError as e:
            messagebox.showerror("Add project", str(e))
            return
        messagebox.showinfo("Project added",
                            f"Found {len(repos)} git repo(s):\n" + "\n".join(repos)
                            + f"\n\nCounting commits by all your accounts: {', '.join(emails) or '(everyone)'}\n"
                              "Use 'Edit emails…' to count only specific emails for this project.")
        self.load_projects()
        self.update_now()

    def edit_emails(self):
        name = self.selected_project()
        if not name:
            return
        p = next(p for p in core.load_config()["projects"] if p["name"] == name)
        if p.get("type") == "github":
            messagebox.showinfo(core.APP_NAME, "GitHub projects count commits by their account's GitHub login "
                                "and emails. Change them in 'Git accounts…'.")
            return
        s = simpledialog.askstring("Count commits by",
                                   "Only count these emails for this project (comma-separated).\n"
                                   "Leave empty to count all your git accounts:",
                                   initialvalue=", ".join(p.get("emails", [])), parent=self)
        if s is None:
            return
        core.update_project_emails(name, s)
        self.load_projects()
        self.update_now()

    def remove_project(self):
        name = self.selected_project()
        if not name:
            return
        ans = messagebox.askyesnocancel(
            "Remove project",
            f"Stop tracking '{name}'?\n\nYes = remove and delete its saved history\n"
            "No = remove but keep saved history (still appears in exports)")
        if ans is None:
            return
        core.remove_project(name, delete_data=ans)
        self.load_projects()
        self.refresh_preview()

    def update_now(self):
        self.set_status("Updating…")

        def done(totals):
            errors = [f"{k}: {v[2]}" for k, v in totals.items() if v[2]]
            self.set_status("  ".join(f"{k}: {v[0]} commits" for k, v in totals.items() if not v[2]))
            self.refresh_preview()
            if errors:
                messagebox.showwarning("Some projects were skipped", "\n".join(errors))

        run_bg(self, lambda: core.update_all(log=lambda m: None), done)

    # ------------------------------------------------ timesheet
    def refresh_preview(self):
        r = self.range()
        if not r:
            return
        self.preview.delete(*self.preview.get_children())
        rows = core.daily_rows(*r)
        for row in rows:
            d = dt.date.fromisoformat(row["date"])
            self.preview.insert("", "end", text=d.strftime("%d-%b-%Y"),
                                values=(d.strftime("%a"), row["project"], f"{row['first']} – {row['last']}",
                                        row["span"], row["commits"], " | ".join(row["subjects"])[:400]))
        self.set_status(f"{len(rows)} day-rows shown")

    def download(self):
        r = self.range()
        if not r:
            return
        path = filedialog.asksaveasfilename(
            title="Save timesheet", defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")],
            initialdir=os.path.join(os.path.expanduser("~"), "Downloads"),
            initialfile=f"timesheet_{r[2] or 'all'}_{r[0]}_to_{r[1]}.xlsx")
        if not path:
            return
        try:
            n = core.export_excel(path, *r)
        except PermissionError:
            messagebox.showerror("Download", "That file is open in Excel. Close it and try again.")
            return
        if messagebox.askyesno("Download", f"Saved {n} day-rows to\n{path}\n\nOpen it now?"):
            os.startfile(path)

    def _set(self, f, t):
        self.v_from.set(f.isoformat())
        self.v_to.set(t.isoformat())
        self.refresh_preview()

    def q_this_month(self):
        t = dt.date.today()
        self._set(t.replace(day=1), t)

    def q_last_month(self):
        end = dt.date.today().replace(day=1) - dt.timedelta(days=1)
        self._set(end.replace(day=1), end)

    def q_this_week(self):
        t = dt.date.today()
        self._set(t - dt.timedelta(days=t.weekday()), t)

    def q_all(self):
        self._set(dt.date.fromisoformat(core.load_config()["since"]), dt.date.today())

    def change_since(self):
        s = simpledialog.askstring("Start date", "Collect commits made on or after (YYYY-MM-DD):",
                                   initialvalue=core.load_config()["since"], parent=self)
        if not s:
            return
        try:
            core.set_since(s.strip())
        except ValueError:
            messagebox.showerror(core.APP_NAME, "Use the format YYYY-MM-DD, e.g. 2026-01-01")
            return
        self.update_now()

    def change_schedule(self):
        s = simpledialog.askstring("Daily update", "Run the automatic update every day at (HH:MM, 24h):",
                                   initialvalue=core.load_config()["daily_time"], parent=self)
        if not s:
            return
        try:
            core.set_schedule(s.strip())
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("Daily update", str(e))
            return
        self.set_status()


if __name__ == "__main__":
    App().mainloop()

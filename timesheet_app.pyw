"""VJ-TimesheetTracker - desktop window.

Projects: add / remove / edit author emails.
Update now: rescan all git repos (also runs daily via Task Scheduler).
Download: Excel timesheet for a date range and project.
"""
import os
import threading
import datetime as dt
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

import tracker_core as core

NAVY = "#162F6A"


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(core.APP_NAME)
        self.geometry("1100x700")
        self.minsize(900, 560)
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
        self.projects = ttk.Treeview(pf, columns=("path", "emails"), height=4)
        self.projects.heading("#0", text="Name")
        self.projects.heading("path", text="Folder")
        self.projects.heading("emails", text="Your git emails (commits counted)")
        self.projects.column("#0", width=160)
        self.projects.column("path", width=380)
        self.projects.column("emails", width=380)
        self.projects.pack(side="left", fill="x", expand=True)
        btns = ttk.Frame(pf)
        btns.pack(side="left", padx=8, fill="y")
        for text, cmd in [("Add project…", self.add_project), ("Edit emails…", self.edit_emails),
                          ("Remove", self.remove_project), ("Update now", self.update_now)]:
            ttk.Button(btns, text=text, command=cmd, width=16).pack(pady=2)

        # ---- export / filter
        mid = ttk.Frame(self, padding=(10, 0))
        mid.pack(fill="x")
        ttk.Label(mid, text="Timesheet", style="H.TLabel").grid(row=0, column=0, sticky="w", columnspan=8)
        today = dt.date.today()
        cfg = core.load_config()
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
                          ("This week", self.q_this_week), (f"All (since {cfg['since']})", self.q_all)]:
            ttk.Button(quick, text=label, command=fn).pack(side="left", padx=(0, 4))

        # ---- preview
        bot = ttk.Frame(self, padding=10)
        bot.pack(fill="both", expand=True)
        cols = ("day", "project", "time", "span", "commits", "work")
        self.preview = ttk.Treeview(bot, columns=cols, show="tree headings")
        self.preview.heading("#0", text="Date")
        for c, t, w in [("day", "Day", 50), ("project", "Project", 130), ("time", "First – Last", 110),
                        ("span", "Span h", 60), ("commits", "Commits", 65), ("work", "Work done", 560)]:
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
        ttk.Button(sf, text="Open data folder", command=lambda: os.startfile(core.HOME)).pack(side="right", padx=6)

        self.load_projects()
        self.update_now(silent=True)

    # ------------------------------------------------ helpers
    def set_status(self, extra=""):
        cfg = core.load_config()
        self.status.set(f"Last update: {core.last_update()}   ·   Daily auto-update at {cfg['daily_time']}"
                        + (f"   ·   {extra}" if extra else ""))

    def load_projects(self):
        self.projects.delete(*self.projects.get_children())
        cfg = core.load_config()
        for p in cfg["projects"]:
            self.projects.insert("", "end", text=p["name"], values=(p["path"], ", ".join(p.get("emails", []))))
        self.proj_box["values"] = ["All projects"] + [p["name"] for p in cfg["projects"]]
        if self.v_proj.get() not in self.proj_box["values"]:
            self.v_proj.set("All projects")
        self.set_status()

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
    def add_project(self):
        path = filedialog.askdirectory(title="Choose the project folder (git repo or folder containing repos)")
        if not path:
            return
        name = simpledialog.askstring("Project name", "Name for this project:",
                                      initialvalue=os.path.basename(os.path.normpath(path)), parent=self)
        if not name:
            return
        try:
            repos, emails = core.add_project(name.strip(), path)
        except ValueError as e:
            messagebox.showerror("Add project", str(e))
            return
        messagebox.showinfo("Project added",
                            f"Found {len(repos)} git repo(s):\n" + "\n".join(repos)
                            + f"\n\nCounting commits by: {', '.join(emails) or '(everyone)'}\n"
                              "Use 'Edit emails…' if you also commit with another address.")
        self.load_projects()
        self.update_now()

    def edit_emails(self):
        name = self.selected_project()
        if not name:
            return
        p = next(p for p in core.load_config()["projects"] if p["name"] == name)
        s = simpledialog.askstring("Your git emails",
                                   "Comma-separated emails whose commits count as yours\n(empty = everyone):",
                                   initialvalue=", ".join(p.get("emails", [])), parent=self)
        if s is None:
            return
        core.update_project_emails(name, [e.strip().lower() for e in s.split(",") if e.strip()])
        with core.db() as c:
            c.execute("DELETE FROM commits WHERE project=?", (name,))
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

    def update_now(self, silent=False):
        self.set_status("Updating…")

        def work():
            try:
                totals = core.update_all(log=lambda m: None)
                msg = "  ".join(f"{k}: {c} commits" for k, (c, _) in totals.items())
                self.after(0, lambda: (self.set_status(msg), self.refresh_preview()))
            except Exception as e:
                self.after(0, lambda: messagebox.showerror("Update failed", str(e)))

        threading.Thread(target=work, daemon=True).start()

    # ------------------------------------------------ timesheet
    def refresh_preview(self):
        r = self.range()
        if not r:
            return
        self.preview.delete(*self.preview.get_children())
        total = 0
        for row in core.daily_rows(*r):
            d = dt.date.fromisoformat(row["date"])
            work = " | ".join(row["subjects"])[:400]
            self.preview.insert("", "end", text=d.strftime("%d-%b-%Y"),
                                values=(d.strftime("%a"), row["project"], f"{row['first']} – {row['last']}",
                                        row["span"], row["commits"], work))
            total += 1
        self.set_status(f"{total} day-rows shown")

    def download(self):
        r = self.range()
        if not r:
            return
        suffix = r[2] or "all"
        path = filedialog.asksaveasfilename(
            title="Save timesheet", defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")],
            initialdir=os.path.join(os.path.expanduser("~"), "Downloads"),
            initialfile=f"timesheet_{suffix}_{r[0]}_to_{r[1]}.xlsx")
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
        except Exception as e:
            messagebox.showerror("Daily update", str(e))
            return
        self.set_status()


if __name__ == "__main__":
    App().mainloop()

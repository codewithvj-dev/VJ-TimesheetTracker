# VJ-TimesheetTracker

Fill your timesheet from your git history. VJ-TimesheetTracker is a small Windows desktop app. It reads the commits **you** made across any number of projects, updates itself every day, and downloads a ready-to-fill Excel timesheet for any date range.

- **Any number of projects.** Point it at a folder. It finds every git repo in that folder and up to two levels below it, so a project made of `backend/` and `frontend/` repos counts as one project.
- **Only your work.** It counts commits by your git email(s), detected automatically and editable per project.
- **Daily auto-update.** A Windows scheduled task rescans every day (19:00 by default). If the PC was off at that time, the task runs at next startup.
- **Keeps push history.** Git forgets pushes after about 90 days, but the tracker saves every push it sees, so the history stays complete.
- **Excel export.** The workbook has four sheets:
  - **Timesheet**: one row per day per project, with first/last commit time, the commit span, an empty **Hours worked** column for you to fill in, the repos, and what you did (commit messages).
  - **Monthly Summary**
  - **All Commits**
  - **Pushes**
- **Local only.** Nothing is uploaded and the app never contacts your git server. Data stays in `%APPDATA%\VJ-TimesheetTracker`.

## Requirements

- Windows 10 or 11
- [Python 3.9+](https://www.python.org/downloads/) with tkinter (included in the standard installer)
- [Git for Windows](https://git-scm.com/download/win) on `PATH`

## Install

```powershell
git clone https://github.com/codewithvj-dev/VJ-TimesheetTracker-.git
cd VJ-TimesheetTracker-
powershell -ExecutionPolicy Bypass -File install.ps1            # daily update at 19:00
powershell -ExecutionPolicy Bypass -File install.ps1 -Time 18:30 # or pick a time
```

The installer:
- installs `openpyxl`
- registers the per-user scheduled task **"VJ-TimesheetTracker Daily Update"** (no admin rights needed)
- adds a **VJ-TimesheetTracker** shortcut to your Desktop and Start Menu

Keep the cloned folder where it is, because the shortcut and the task run from it. To upgrade, `git pull`.

## Use

1. Open **VJ-TimesheetTracker** and click **Add project…**, then choose a project folder.
2. If you commit under more than one address (work and personal), select the project and use **Edit emails…** to list them all.
3. Choose a range (**This month**, **Last month**, **This week**, or type From/To as `YYYY-MM-DD`) and a project, or **All projects**.
4. Click **Show** to preview, or **Download Excel…** to save the timesheet.

At the bottom of the window:
- **Change start date…** sets how far back commits are collected. The default is 1 January of last year.
- **Change daily time…** moves the scheduled update.
- **Open data folder** shows where your data and logs are stored.

### Command line

```powershell
python tracker_core.py add C:\path\to\project --name MyProject [--email me@example.com]
python tracker_core.py list
python tracker_core.py update
python tracker_core.py export timesheet.xlsx --from 2026-11-01 [--to 2026-11-30] [--project MyProject]
python tracker_core.py schedule 19:00
python tracker_core.py unschedule
```

## Good to know

- **"Commit span" isn't hours worked.** It is the time from your first commit to your last commit that day, so treat it as a hint. That's why **Hours worked** is left for you to fill in.
- **Only commits in your local clones are counted.** Run `git fetch` in a repo first if you also commit from another machine.
- **Push history only covers what is still in your local reflog.** Git keeps pushes for about 90 days, so after installing you get roughly the last 90 days of pushes, and every push after that.
- **You can move the data folder.** Set the `VJTT_HOME` environment variable to keep data somewhere other than `%APPDATA%\VJ-TimesheetTracker`.

## Uninstall

```powershell
powershell -ExecutionPolicy Bypass -File uninstall.ps1              # keeps your data
powershell -ExecutionPolicy Bypass -File uninstall.ps1 -RemoveData  # deletes it too
```

## License

[MIT](LICENSE)

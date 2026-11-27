# VJ-TimesheetTracker

**Fill your timesheet from your git history.** VJ-TimesheetTracker is a small Windows desktop app. It collects the commits and pushes **you** made, across as many projects and git accounts as you have, and turns them into a ready-to-fill Excel timesheet for any date range. It updates itself every day, so the history is always current.

```
┌ VJ-TimesheetTracker ──────────────────────────────────────────────────────────────┐
│ Projects                                                       [Git accounts…    ] │
│  Name        Source                              Counting by    [Add local project…] │
│  my-app      C:\Work\my-app                      All my accounts [Add GitHub repos… ] │
│  client-x    GitHub (Work): acme/api, acme/web   jdoe + 3 emails [Edit emails…      ] │
│                                                                 [Remove] [Update now] │
│ Timesheet  From [2026-11-01] To [2026-11-30] Project [All ▾] [Show] [Download Excel…] │
│            [This month] [Last month] [This week] [Everything]                         │
│  Date         Day  Project   First – Last    Span h  Commits  Work done              │
│  03-Nov-2026  Mon  my-app    10:12 – 17:40   7.5     6        Fix login | Add export… │
│ Last update: 2026-11-27 19:00 · Daily auto-update at 19:00 · Since 2026-01-01        │
└────────────────────────────────────────────────────────────────────────────────────┘
```

## Contents

1. [Features](#features)
2. [Requirements](#requirements)
3. [Install](#install)
4. [Quick start (5 minutes)](#quick-start-5-minutes)
5. [Git accounts: single or multiple](#git-accounts-single-or-multiple)
6. [Connecting GitHub](#connecting-github)
7. [Adding projects](#adding-projects)
8. [Downloading your timesheet](#downloading-your-timesheet)
9. [Daily automatic update](#daily-automatic-update)
10. [Command line](#command-line)
11. [Where your data is & privacy](#where-your-data-is--privacy)
12. [Troubleshooting](#troubleshooting)
13. [Update / uninstall](#update--uninstall)
14. [License](#license)

## Features

- **One or many git accounts.** Add each identity you commit with (Work, Personal, Client…). Every account can have several emails.
- **Local projects.** Point the app at a folder. It finds every git repo in that folder and up to two levels below it, so `backend/` and `frontend/` repos count as one project. The local server you push to can be any host (GitHub, GitLab, Bitbucket, Azure DevOps or self-hosted), because the app only reads the repos on your disk.
- **GitHub projects without cloning.** Connect a GitHub account with a token and pick repos from a list. Commits and pushes are read through the GitHub API.
- **Daily auto-update.** A Windows scheduled task runs at 19:00 by default. If the PC was off at that time, it runs at the next startup.
- **Keeps push history.** Git forgets pushes after about 90 days, but the app saves every push it sees.
- **Excel export.** The workbook has four sheets:
  - **Timesheet**: one row per day per project, with first/last commit time, the commit span, an empty **Hours worked** column, the repos, and the work done.
  - **Monthly Summary**
  - **All Commits**
  - **Pushes**
- **Local and private.** Data stays on your PC, and GitHub tokens are encrypted with Windows DPAPI.

## Requirements

| What | Why |
|---|---|
| Windows 10 / 11 | Uses Task Scheduler and Windows encryption (DPAPI) |
| [Python 3.9+](https://www.python.org/downloads/) | Runs the app. Tick **"Add python.exe to PATH"** in the installer. tkinter is included by default |
| [Git for Windows](https://git-scm.com/download/win) | Reads local repos |

## Install

Open **PowerShell** and run:

```powershell
git clone https://github.com/codewithvj-dev/VJ-TimesheetTracker-.git "$HOME\VJ-TimesheetTracker"
cd "$HOME\VJ-TimesheetTracker"
powershell -ExecutionPolicy Bypass -File install.ps1
```

To run the daily update at a different time, use for example `install.ps1 -Time 18:30`.

The installer, which needs no admin rights:
1. checks Python, tkinter and git
2. installs the one dependency, `openpyxl`
3. creates the scheduled task **"VJ-TimesheetTracker Daily Update"**
4. adds a **VJ-TimesheetTracker** shortcut to your **Desktop** and **Start Menu**

> Keep the cloned folder where it is. The shortcut and the daily task run from it.

## Quick start (5 minutes)

1. **Open the app** from the Desktop shortcut. On first start it reads your git identity (`git config --global user.email`) and creates your first account from it.
2. **Add a project.** Click **Add local project…** and choose the folder of a project you work on, or use **Add GitHub repos…** (see [Connecting GitHub](#connecting-github)).
3. **Check the start date.** Click **Change start date…** at the bottom if you need history from before 1 January of last year.
4. **Download.** Click **This month** (or type From/To dates), then **Download Excel…**. Fill in the yellow **Hours worked** column and submit.

From then on the app updates itself daily. Just open it and download.

## Git accounts: single or multiple

Click **Git accounts…** to see and manage your identities.

### Single account (most people)

Nothing to do. The account created on first start holds your global git email. Every local project counts commits made with that email.

Check which email your commits use by running this inside any repo:

```powershell
git config user.email
```

If a repo uses a different email than your global one, add it to the account with **Git accounts… → Edit…**.

### Multiple accounts (work + personal, several clients…)

1. **Git accounts… → Add account…** and give it a name such as `Work`.
2. Enter the emails you commit with on that account, comma-separated:
   `jdoe@company.com, jdoe@users.noreply.github.com`
3. Repeat for each identity. Optionally **Connect GitHub…** for each one (see below). Each account can connect its **own** GitHub login.

How commits are counted:

| Project type | Counts commits by |
|---|---|
| Local project, default | **All my accounts**: every email of every account |
| Local project with **Edit emails…** set | Only the emails you list for that project, e.g. only your work email for a client repo |
| GitHub project | The GitHub login **and** the emails of the account it was added with |

> **Tip:** a local project with the wrong emails shows 0 commits. Open **All Commits** in an export, or run `git log --format=%ae | sort -u` in the repo, to see which emails actually appear.

## Connecting GitHub

Connect a GitHub account when you want to track repos **you haven't cloned** on this PC, or to pick up work pushed from other machines.

### 1. Create a token

**Fine-grained token (recommended):**
1. GitHub → **Settings → Developer settings → Personal access tokens → Fine-grained tokens → Generate new token**. The app's **Create a token on GitHub…** button opens this page.
2. **Resource owner**: your user, or the organisation that owns the repos.
3. **Repository access**: *All repositories*, or select the ones you want.
4. **Permissions → Repository**: **Contents: Read-only**. Metadata is added automatically.
5. Choose an expiry, generate the token, and copy it.

**Classic token:** select the `repo` scope for private repos (none needed for public ones).

> Organisation repos may need the token approved by an org admin, or **Configure SSO → Authorize** if the org uses SAML single sign-on.

### 2. Connect

**Git accounts…** → select the account → **Connect GitHub…** → paste the token → **OK**.

- The app checks the token with GitHub and shows the login it connected to.
- It adds GitHub's no-reply addresses (`…@users.noreply.github.com`) to the account, so commits made in the GitHub web UI are counted too.
- The token is stored encrypted for your Windows user only, never in plain text.

**GitHub Enterprise Server:** change the API URL in the same dialog to `https://<your-host>/api/v3`.

### 3. Add repos

Click **Add GitHub repos…**, choose the account, then search and select one or more repos (Ctrl/Shift-click). Give the project a name and click **Add project**. Several repos in one project show up as one line per day in the timesheet.

To remove a connection, use **Git accounts… → Disconnect GitHub**. You can also revoke the token on GitHub at any time.

## Adding projects

| | **Add local project…** | **Add GitHub repos…** |
|---|---|---|
| Needs | The repo cloned on this PC | A connected GitHub account |
| Git host | Any (GitHub, GitLab, Bitbucket, self-hosted…) | GitHub / GitHub Enterprise |
| Branches | All local and remote-tracking branches | The repo's **default branch** |
| Pushes | From this PC's reflog (about 90 days back, then kept forever) | From GitHub's activity feed (about 90 days, max 300 events, then kept forever) |
| Works offline | Yes | No (skipped with a warning) |

**Edit emails…** (local projects only) limits which emails count for that one project. **Remove** asks whether to keep or delete the project's saved history.

> Don't add the same repo both locally and through GitHub, or its commits are counted twice.

## Downloading your timesheet

1. Pick **From / To**: type `YYYY-MM-DD` or use **This month / Last month / This week / Everything**.
2. Pick a **Project**, or **All projects**.
3. Click **Show** to preview, then **Download Excel…** to save (Downloads by default).

| Sheet | Contents |
|---|---|
| **Timesheet** | Date, Day, Project, First commit, Last commit, Commit span (h), **Hours worked** (yellow, for you to fill), Commits, Repos, Work done (bullet list of commit messages) |
| **Monthly Summary** | Active days and commits per month per project, with a total |
| **All Commits** | Every commit: date, time, project, repo, short hash, message |
| **Pushes** | Every push: date, time, project, repo, branch, commit |

> **Commit span ≠ hours worked.** The span is the time from your first commit to your last commit that day. Use it as a hint and enter your real hours.

## Daily automatic update

- The update runs every day at the time shown in the status bar, and again whenever you open the app or click **Update now**.
- To move it, click **Change daily time…**.
- Each run is logged in `%APPDATA%\VJ-TimesheetTracker\logs\YYYY-MM.log`. Click **Open data folder** to find it.
- You can see or run the task manually in **Task Scheduler → VJ-TimesheetTracker Daily Update**.

## Command line

Everything in the window can also be done from PowerShell in the install folder:

```powershell
# accounts
python tracker_core.py accounts
python tracker_core.py account-add Work --email jdoe@company.com --email jdoe@users.noreply.github.com
python tracker_core.py github-connect <account-id>          # asks for the token (hidden input)
python tracker_core.py github-repos <account-id>
python tracker_core.py github-disconnect <account-id>
python tracker_core.py account-remove <account-id>

# projects
python tracker_core.py add C:\Work\my-app --name my-app [--email only-this@company.com]
python tracker_core.py add-github client-x --account <account-id> --repo acme/api --repo acme/web
python tracker_core.py list
python tracker_core.py remove my-app [--delete-data]

# data & export
python tracker_core.py update
python tracker_core.py since 2026-01-01
python tracker_core.py export nov.xlsx --from 2026-11-01 --to 2026-11-30 [--project my-app]

# schedule
python tracker_core.py schedule 19:00
python tracker_core.py unschedule
```

## Where your data is & privacy

| | |
|---|---|
| Program | The folder you cloned (contains no personal data) |
| Your data | `%APPDATA%\VJ-TimesheetTracker\`: `config.json` (accounts, projects, settings), `data.db` (commits & pushes), `logs\` |
| Move the data | Set the environment variable `VJTT_HOME` to another folder |
| GitHub tokens | Encrypted with Windows DPAPI for your user, stored in `config.json`. Unreadable by other users or on another PC |
| Network | Only `api.github.com` (or your Enterprise URL), and only for connected accounts. Local projects never touch the network |

## Troubleshooting

| Problem | Fix |
|---|---|
| A project shows **0 commits** | Your commits use an email the app doesn't know. Run `git log -5 --format=%ae` in the repo and add that email in **Git accounts… → Edit…** |
| Commits from my other PC are missing | Local projects read only this clone, so run `git fetch` first. Or track the repo with **Add GitHub repos…** |
| History before a date is missing | **Change start date…** to an earlier date. The app rescans automatically |
| **"GitHub rejected the token (401)"** | The token expired or was revoked. Create a new one and **Connect GitHub…** again |
| **"GitHub refused access (403)"** | The token lacks *Contents: Read*, the org requires SSO authorisation, or you hit the rate limit (wait an hour) |
| A GitHub repo is missing from the list | The token doesn't include that repo or organisation. Edit the token's repository access |
| "Some projects were skipped" | The message names the project and the reason (folder moved, no internet, token problem). The others still updated |
| "That file is open in Excel" | Close the earlier export and download again |
| The daily update didn't run | Check `logs\` and **Task Scheduler → VJ-TimesheetTracker Daily Update → History**. Run `install.ps1` again to re-create the task |

## Update / uninstall

**Update** to the latest version. Your data is untouched:

```powershell
cd "$HOME\VJ-TimesheetTracker"
git pull
```

**Uninstall:**

```powershell
powershell -ExecutionPolicy Bypass -File uninstall.ps1              # removes task + shortcuts, keeps data
powershell -ExecutionPolicy Bypass -File uninstall.ps1 -RemoveData  # also deletes your data and tokens
```

## License

[MIT](LICENSE): free to use, modify and share.

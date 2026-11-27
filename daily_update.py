"""Run by Windows Task Scheduler once a day: rescans every project into data.db."""
import os
import datetime as dt
import traceback
import tracker_core as core

os.makedirs(core.LOGS, exist_ok=True)
log_path = os.path.join(core.LOGS, dt.date.today().strftime("%Y-%m") + ".log")
with open(log_path, "a", encoding="utf-8") as f:
    stamp = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    try:
        core.update_all(log=lambda m: f.write(f"{stamp} {m}\n"))
        f.write(f"{stamp} update OK\n")
    except Exception:
        f.write(f"{stamp} update FAILED\n{traceback.format_exc()}\n")

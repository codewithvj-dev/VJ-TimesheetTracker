"""Run by Windows Task Scheduler once a day (source install): rescans every project into data.db."""
import tracker_core as core

core.run_daily_update()

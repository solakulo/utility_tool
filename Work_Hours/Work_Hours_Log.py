import os
import win32evtlog
from datetime import datetime

LOG_DIR = "logs"

if not os.path.exists(LOG_DIR):
    os.makedirs(LOG_DIR)


def extract_logs():
    server = None
    log_type = "System"
    flags = win32evtlog.EVENTLOG_BACKWARDS_READ | win32evtlog.EVENTLOG_SEQUENTIAL_READ

    handle = win32evtlog.OpenEventLog(server, log_type)
    daily_logs = {}
    today = datetime.now().strftime("%Y-%m-%d")

    while True:
        events = win32evtlog.ReadEventLog(handle, flags, 0)
        if not events:
            break
        for event in events:
            if event.EventID in [30, 7002]:  # 开机 (30) 和 关机 (7002)
                date = event.TimeGenerated.strftime("%Y-%m-%d")
                if date == today:  # 跳过当天的日志
                    continue
                time = event.TimeGenerated.strftime("%H:%M:%S")
                if date not in daily_logs:
                    daily_logs[date] = {"first": None, "last": None}

                if event.EventID == 30:  # 开机
                    if (
                        daily_logs[date]["first"] is None
                        or time < daily_logs[date]["first"]
                    ):
                        daily_logs[date]["first"] = time
                elif event.EventID == 7002:  # 关机
                    if (
                        daily_logs[date]["last"] is None
                        or time > daily_logs[date]["last"]
                    ):
                        daily_logs[date]["last"] = time

    win32evtlog.CloseEventLog(handle)
    return daily_logs


def save_logs(daily_logs):
    yearly_logs = {}
    for date, times in daily_logs.items():
        year = date.split("-")[0]
        if year not in yearly_logs:
            yearly_logs[year] = []
        yearly_logs[year].append(
            f"{date}, {times['first'] or '无记录'}, {times['last'] or '无记录'}"
        )

    for year, logs in yearly_logs.items():
        log_file = os.path.join(LOG_DIR, f"{year}.log")
        with open(log_file, "w", encoding="utf-8") as f:
            f.write("\n".join(logs))
        print(f"日志已保存到 {log_file}")


if __name__ == "__main__":
    logs = extract_logs()
    save_logs(logs)

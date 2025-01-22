import os
import json
import openpyxl
from datetime import datetime, timedelta
import sys

CONFIG_FILE = "work_schedule_config.json"
LOG_DIR = "logs"


def load_config():
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        return json.load(f)


def read_logs(year):
    log_file = os.path.join(LOG_DIR, f"{year}.log")
    if not os.path.exists(log_file):
        raise FileNotFoundError(f"日志文件 {log_file} 不存在！")

    logs = []
    with open(log_file, "r", encoding="utf-8") as f:
        for line in f:
            parts = line.strip().split(", ")
            if len(parts) == 3:
                date, first_time, last_time = parts
                logs.append(
                    {"date": date, "first_time": first_time, "last_time": last_time}
                )
    return logs


def calculate_rest_duration(work_start, work_end, rest_periods):
    total_rest_duration = timedelta()
    for period in rest_periods:
        start, end = period.split("-")
        rest_start = datetime.strptime(start, "%H:%M")
        rest_end = datetime.strptime(end, "%H:%M")

        # 1.1 完整包括在上下班时间段之间
        if rest_start >= work_start and rest_end <= work_end:
            total_rest_duration += rest_end - rest_start

        # 1.2 完整在上下班时间之外 (不计算)
        elif rest_end <= work_start or rest_start >= work_end:
            continue

        # 1.3 与上下班时间段有交集
        else:
            overlap_start = max(work_start, rest_start)
            overlap_end = min(work_end, rest_end)
            if overlap_end > overlap_start:
                total_rest_duration += overlap_end - overlap_start

    return total_rest_duration


def calculate_work_hours(log, config):
    work_start = datetime.strptime(log["first_time"], "%H:%M:%S")
    work_end = datetime.strptime(log["last_time"], "%H:%M:%S")

    # 计算包含所有休息时间的上班时长
    total_work_duration = work_end - work_start

    # 计算当天有效的休息时间段
    rest_periods = config.get("rest_periods", [])
    total_rest_duration = calculate_rest_duration(work_start, work_end, rest_periods)

    # 去除休息的上班时长
    effective_work_duration = total_work_duration - total_rest_duration

    # 标准工作时长
    standard_duration = timedelta(
        hours=int(config["standard_work_hours"].split(":")[0]),
        minutes=int(config["standard_work_hours"].split(":")[1]),
    )

    # 计算加班时长
    overtime = max(timedelta(0), effective_work_duration - standard_duration)

    return total_work_duration, effective_work_duration, overtime


def save_to_excel(logs, config, year):
    output_file = f"system_logs_{year}.xlsx"
    wb = openpyxl.Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)

    for month in range(1, 13):
        ws = wb.create_sheet(title=f"{month:02}")
        ws.append(
            [
                "日期",
                "最早开机时间",
                "最晚关机时间",
                "包含休息上班时长",
                "去除休息上班时长",
                "加班时长",
            ]
        )

    for log in logs:
        date = log["date"]
        month = int(date.split("-")[1])
        ws = wb[f"{month:02}"]

        total, effective, overtime = calculate_work_hours(log, config)
        ws.append(
            [
                date,
                log["first_time"],
                log["last_time"],
                str(total),
                str(effective),
                str(overtime),
            ]
        )

    wb.save(output_file)
    print(f"生成的图表已保存到 {output_file}")


if __name__ == "__main__":
    year = sys.argv[1]
    config = load_config()
    logs = read_logs(year)
    save_to_excel(logs, config, year)

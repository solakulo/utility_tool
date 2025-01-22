import os
import json
from tkinter import (
    Tk,
    Label,
    Entry,
    Button,
    StringVar,
    messagebox,
    filedialog,
    simpledialog,
)

# 配置文件路径
CONFIG_FILE = "work_schedule_config.json"


def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"standard_work_hours": "7:45", "rest_periods": []}


def save_config(config):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config, f, indent=4, ensure_ascii=False)


class WorkScheduleUI:
    def __init__(self, root):
        self.root = root
        self.root.title("工作时间设置")
        self.config = load_config()

        self.init_ui()

    def init_ui(self):
        Label(self.root, text="标准上班时长:").grid(row=0, column=0, sticky="w")
        self.work_hours_var = StringVar(value=self.config["standard_work_hours"])
        Entry(self.root, textvariable=self.work_hours_var).grid(row=0, column=1)

        Label(self.root, text="休息时间段 (用逗号分隔):").grid(
            row=1, column=0, sticky="w"
        )
        self.rest_periods_var = StringVar(value=",".join(self.config["rest_periods"]))
        Entry(self.root, textvariable=self.rest_periods_var).grid(row=1, column=1)

        Button(self.root, text="保存", command=self.save_config).grid(row=2, column=0)
        Button(self.root, text="关闭", command=self.on_close).grid(row=2, column=1)

        Button(self.root, text="开关机情报取得", command=self.run_log_extraction).grid(
            row=3, column=0
        )
        Button(self.root, text="生成图表", command=self.run_chart_generation).grid(
            row=3, column=1
        )

    def save_config(self):
        self.config["standard_work_hours"] = self.work_hours_var.get()
        self.config["rest_periods"] = [
            x.strip() for x in self.rest_periods_var.get().split(",") if x.strip()
        ]
        save_config(self.config)
        messagebox.showinfo("提示", "配置已保存！")

    def on_close(self):
        if messagebox.askyesno("确认", "是否保存修改？"):
            self.save_config()
        self.root.quit()

    def run_log_extraction(self):
        os.system("python Work_Hours_Log.py")

    def run_chart_generation(self):
        year = simpledialog.askstring("选择年份", "请输入要生成图表的年份:")
        if year:
            os.system(f"python Calculate_Work_Hours.py {year}")


if __name__ == "__main__":
    root = Tk()
    app = WorkScheduleUI(root)
    root.mainloop()

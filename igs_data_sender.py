import pandas as pd
import time
from opcua import Client
from opcua.ua import DataValue, Variant, VariantType
import os
import threading
import tkinter as tk
from tkinter import scrolledtext
from tkinter import messagebox

stop_flag = False
pause_flag = False
running = True
auto_scroll = True


def on_text_scroll(*args):
    global auto_scroll
    current_pos = text_area.yview()[1]
    auto_scroll = current_pos >= 0.95


def log(msg):
    if running:

        def update():
            text_area.insert(tk.END, msg + "\n")

            lines = text_area.get("1.0", tk.END).count("\n")
            if lines > 1000:
                text_area.delete("1.0", f"{lines - 1000 + 1}.0")

            if auto_scroll:
                text_area.see(tk.END)

        root.after(0, update)


def check_control_files():
    global stop_flag
    while running:
        time.sleep(1)
        if os.path.exists("stop.txt"):
            stop_flag = True
            log("检测到 stop.txt，程序将停止...")


def run_task():
    global START_ROW, stop_flag, pause_flag, running, client, data_df, tag_columns

    try:
        log(f"数据列: {tag_columns}")
        log(f"总行数: {len(data_df)}")
        log(f"从第 {START_ROW + 1} 行开始输入\n")
        log("控制方式: stop.txt=停止, pause.txt=暂停\n")

        control_thread = threading.Thread(target=check_control_files, daemon=True)
        control_thread.start()

        log("控制方式: stop.txt=停止, 暂停按钮=暂停\n")

        while running:
            if stop_flag:
                log("程序已停止")
                break

            while pause_flag and running:
                time.sleep(1)
                if stop_flag:
                    break

            if stop_flag:
                break

            interval = int(entry_interval.get())
            log(f"写入间隔: {interval}秒")

            if START_ROW >= len(data_df):
                log("数据已全部输入完成")
                break

            row = data_df.iloc[START_ROW]

            for tag in tag_columns:
                value = row[tag]
                node_id = f"ns=2;s={CHANNEL_NAME}.{DEVICE_NAME}.{tag}"

                try:
                    if tag in ["IsWork_1"]:
                        vtype = VariantType.Boolean
                        value = bool(value)
                    elif isinstance(value, (int, float)) and value == int(value):
                        vtype = VariantType.Int32
                        value = int(value)
                    else:
                        vtype = VariantType.Float
                        value = float(value)

                    node = client.get_node(node_id)
                    dv = DataValue(Variant(value, vtype))
                    node.set_value(dv)
                    log(f"[{START_ROW + 1}] {tag} = {value}")
                except Exception as e:
                    log(f"写入 {tag} 失败: {e}")

            log(f"--- 第 {START_ROW + 1} 行输入完成 ---\n")

            START_ROW += 1

            for _ in range(interval):
                if stop_flag or pause_flag or not running:
                    break
                time.sleep(1)

        if client:
            client.disconnect()
            log("已断开连接")
    except Exception as e:
        log(f"错误: {e}")
    finally:
        log("程序结束")
        btn_start.config(state=tk.NORMAL)
        btn_pause.config(state=tk.DISABLED)
        btn_resume.config(state=tk.DISABLED)


def on_closing():
    global running, stop_flag
    if messagebox.askokcancel("退出", "确定要停止程序并退出吗?"):
        running = False
        stop_flag = True
        root.destroy()


START_ROW = 0
client = None
data_df = None
tag_columns = []

OPC_UA_SERVER_URL = "opc.tcp://127.0.0.1:49310"
CHANNEL_NAME = "模拟数据"
DEVICE_NAME = "PAC"
EXCEL_FILE = "PAC数据.xlsx"


def init_connection():
    global client, data_df, tag_columns
    try:
        log("正在读取Excel文件...")
        data_df = pd.read_excel(EXCEL_FILE)
        tag_columns = [col for col in data_df.columns if col not in ["id", "datetime"]]
        log(f"Excel读取完成，共 {len(data_df)} 行")

        log("正在连接OPC服务器...")
        client = Client(OPC_UA_SERVER_URL)
        client.connect()
        log(f"已连接到 {OPC_UA_SERVER_URL}")

        btn_start.config(state=tk.NORMAL)
        lbl_status.config(text="状态: 已就绪", fg="green")
        log("连接成功，可以开始写入")
    except Exception as e:
        log(f"初始化失败: {e}")
        lbl_status.config(text="状态: 连接失败", fg="red")


root = tk.Tk()
root.title("IGS数据发送工具")
root.geometry("600x550")
root.protocol("ON_CLOSE", on_closing)

text_area = scrolledtext.ScrolledText(root, width=70, height=22)
text_area.pack(padx=10, pady=10)
text_area.config(yscrollcommand=on_text_scroll)

frame = tk.Frame(root)
frame.pack(pady=5)

lbl_interval = tk.Label(frame, text="写入间隔(秒):")
lbl_interval.pack(side=tk.LEFT, padx=5)

entry_interval = tk.Entry(frame, width=6)
entry_interval.insert(0, "60")
entry_interval.pack(side=tk.LEFT, padx=5)


def on_pause():
    global pause_flag
    pause_flag = True
    btn_pause.config(state=tk.DISABLED)
    btn_resume.config(state=tk.NORMAL)
    log("已暂停，可修改间隔后点击继续")


def on_resume():
    global pause_flag
    pause_flag = False
    btn_resume.config(state=tk.DISABLED)
    btn_pause.config(state=tk.NORMAL)
    log("继续运行")


btn_start = tk.Button(
    frame,
    text="开始运行",
    state=tk.DISABLED,
    command=lambda: [
        btn_start.config(state=tk.DISABLED),
        btn_pause.config(state=tk.NORMAL),
        threading.Thread(target=run_task, daemon=True).start(),
    ],
)
btn_start.pack(side=tk.LEFT, padx=5)

btn_pause = tk.Button(frame, text="暂停", state=tk.DISABLED, command=on_pause)
btn_pause.pack(side=tk.LEFT, padx=5)

btn_resume = tk.Button(frame, text="继续", state=tk.DISABLED, command=on_resume)
btn_resume.pack(side=tk.LEFT, padx=5)

lbl_info = tk.Label(frame, text="控制: stop.txt=停止")
lbl_info.pack(side=tk.LEFT, padx=20)

lbl_status = tk.Label(frame, text="状态: 等待连接...", fg="orange")
lbl_status.pack(side=tk.LEFT, padx=20)

root.update()

log("正在连接OPC服务器...")
threading.Thread(target=init_connection, daemon=True).start()

root.mainloop()

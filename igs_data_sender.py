import pandas as pd
import time
import sys
from opcua import Client
from opcua.ua import DataValue, Variant, VariantType
import os
import threading
import tkinter as tk
from tkinter import scrolledtext, filedialog, messagebox

stop_flag = False
pause_flag = False
running = True
auto_scroll = True

if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))


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
    global START_ROW, stop_flag, pause_flag

    channel = entry_channel.get().strip() or CHANNEL_NAME
    device = entry_device.get().strip() or DEVICE_NAME

    try:
        log(f"数据列: {tag_columns}")
        log(f"总行数: {len(data_df)}")
        log(f"从第 {START_ROW + 1} 行开始输入\n")
        log("控制方式: stop.txt=停止, 暂停按钮=暂停\n")

        control_thread = threading.Thread(target=check_control_files, daemon=True)
        control_thread.start()

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
                node_id = f"ns=2;s={channel}.{device}.{tag}"

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
CHANNEL_NAME = "PAC"
DEVICE_NAME = "PLC"
SELECTED_FILE = ""


def select_excel():
    global SELECTED_FILE
    file_path = filedialog.askopenfilename(
        title="选择Excel文件",
        initialdir=BASE_DIR,
        filetypes=[("Excel文件", "*.xlsx *.xls")],
    )
    if file_path:
        SELECTED_FILE = file_path
        lbl_file.config(text=os.path.basename(file_path))
        log(f"已选择文件: {file_path}")
        threading.Thread(target=init_connection, daemon=True).start()


def init_connection():
    global client, data_df, tag_columns
    try:
        log("正在读取Excel文件...")
        data_df = pd.read_excel(SELECTED_FILE)
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
root.title("IGS数据发送工具 v1.0.0")
root.geometry("700x580")
root.protocol("ON_CLOSE", on_closing)

text_area = scrolledtext.ScrolledText(root, width=80, height=22)
text_area.pack(padx=10, pady=10)
text_area.config(yscrollcommand=on_text_scroll)

file_frame = tk.Frame(root)
file_frame.pack(pady=5)

lbl_file_title = tk.Label(file_frame, text="Excel文件:")
lbl_file_title.pack(side=tk.LEFT, padx=5)

lbl_file = tk.Label(file_frame, text="未选择", fg="gray")
lbl_file.pack(side=tk.LEFT, padx=5)

btn_select = tk.Button(file_frame, text="选择文件", command=select_excel)
btn_select.pack(side=tk.LEFT, padx=5)

frame = tk.Frame(root)
frame.pack(pady=5)

lbl_channel = tk.Label(frame, text="通道名:")
lbl_channel.pack(side=tk.LEFT, padx=5)
entry_channel = tk.Entry(frame, width=10)
entry_channel.insert(0, CHANNEL_NAME)
entry_channel.pack(side=tk.LEFT, padx=5)

lbl_device = tk.Label(frame, text="设备名:")
lbl_device.pack(side=tk.LEFT, padx=5)
entry_device = tk.Entry(frame, width=10)
entry_device.insert(0, DEVICE_NAME)
entry_device.pack(side=tk.LEFT, padx=5)

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

lbl_status = tk.Label(frame, text="状态: 请选择Excel文件", fg="gray")
lbl_status.pack(side=tk.LEFT, padx=20)

log("请选择Excel文件以开始")

root.mainloop()

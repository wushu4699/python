import warnings
warnings.filterwarnings("ignore", category=DeprecationWarning, module="telnetlib")

"""
网络设备自动化巡检系统
版本：5.31（迪普设备分页提示过滤版）
功能：
1. 新增迪普设备`terminal line 0`命令禁用分页，彻底避免分页提示
2. 增强输出清理逻辑，过滤所有分页相关冗余内容
3. 保留多线程巡检、日志记录等全部原有功能
"""

import logging
import re
import os
import time
import platform
import pandas as pd
import tkinter as tk
from tkinter import messagebox, filedialog
from netmiko import ConnectHandler, NetmikoAuthenticationException, NetmikoTimeoutException
from netmiko.base_connection import BaseConnection
from netmiko.ssh_dispatcher import CLASS_MAPPER
from concurrent.futures import ThreadPoolExecutor, as_completed
import shutil
import telnetlib
import threading

# 全局配置（保持不变）
DEVICE_TYPES = {
    "华为": "huawei",
    "思科": "cisco_ios",
    "华三": "hp_comware",
    "锐捷": "ruijie_os",
    "中兴": "zte_zxros",
    "迪普": "dptech_os"
}
PROTOCOL_MAP = {
    "ssh": "",
    "telnet": "_telnet"
}
DEFAULT_PORTS = {
    "ssh": 22,
    "telnet": 23
}
SUPPORTED_DEVICES = set(CLASS_MAPPER.keys())
RETRY_CONFIG = {
    "max_retries": 5,
    "retry_interval": 0.5
}

# 自定义迪普设备Telnet连接类（新增禁用分页命令）
class DPTechTelnet(BaseConnection):
    def session_preparation(self):
        self._test_channel_read()
        self.set_base_prompt()
        # 发送禁用分页命令（支持terminal line 0）
        self.write_channel(b'terminal line 0\r\n')
        time.sleep(0.5)
        self._test_channel_read()  # 清除命令执行后的输出

    def set_base_prompt(self):
        prompt = super().set_base_prompt()
        return re.sub(r'\n+|\r+', '', prompt).strip()

CLASS_MAPPER["dptech_os_telnet"] = DPTechTelnet

# 配置日志
def setup_logging():
    logging.basicConfig(
        level=logging.DEBUG,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler('network_inspection.log', encoding='utf-8'),
            logging.StreamHandler()
        ]
    )

# 清理旧文件
def clean_old_files(result_dir, keep_files=10):
    if not os.path.exists(result_dir):
        return
    try:
        dirs = sorted(os.listdir(result_dir), key=lambda x: os.path.getmtime(os.path.join(result_dir, x)), reverse=True)
        for dirname in dirs[keep_files:]:
            shutil.rmtree(os.path.join(result_dir, dirname))
    except Exception as e:
        logging.warning(f"清理文件失败: {dirname} - {str(e)}")

# 播放提示音
def play_sound():
    try:
        if platform.system() == "Windows":
            import winsound
            winsound.Beep(1000, 1000)
        elif platform.system() == "Linux":
            os.system('play -nq -t alsa synth 0.2 sine 440')
    except Exception as e:
        logging.error(f"播放提示音失败: {str(e)}")

# 生成Excel模板
def generate_excel_template():
    required_columns = ["IP地址", "设备品牌", "密码"]
    try:
        with pd.ExcelWriter("devices_info_template.xlsx") as writer:
            pd.DataFrame(columns=required_columns).to_excel(writer, sheet_name="devices", index=False)
            pd.DataFrame(columns=["设备品牌", "巡检命令"]).to_excel(writer, sheet_name="巡检命令", index=False)
    except Exception as e:
        logging.error(f"生成Excel模板失败: {str(e)}")

# 读取巡检命令选项卡的数据
def load_inspection_commands(excel_path="devices_info.xlsx"):
    try:
        if not excel_path.endswith('.xlsx'):
            logging.error(f"文件格式错误，仅支持 .xlsx 文件: {excel_path}")
            return {}
        df = pd.read_excel(excel_path, sheet_name="巡检命令").fillna("")
        command_map = {}
        for idx, row in df.iterrows():
            brand = row.get("设备品牌")
            commands = parse_commands(row.get("巡检命令", ""))
            if not brand:
                logging.warning(f"第{idx + 2}行 '设备品牌' 字段为空，请检查 '巡检命令' 工作表。")
            if not commands:
                logging.warning(f"第{idx + 2}行 '巡检命令' 字段为空，请检查 '巡检命令' 工作表。")
            if brand and commands:
                command_map[brand] = commands
        return command_map
    except FileNotFoundError:
        logging.error(f"未找到 Excel 文件: {excel_path}")
        return {}
    except PermissionError:
        logging.error(f"Excel 文件被占用，无法读取: {excel_path}")
        return {}
    except Exception as e:
        logging.error(f"读取巡检命令失败: {str(e)}")
        return {}

# 解析命令字符串
def parse_commands(raw_str):
    return [cmd.strip() for cmd in re.split(r'[;,\n|]', raw_str) if cmd.strip()] if raw_str else []

# 数据处理模块
def load_devices(excel_path="devices_info.xlsx"):
    if not os.path.exists("devices_info_template.xlsx"):
        generate_excel_template()
    try:
        if not excel_path.endswith('.xlsx'):
            logging.error(f"文件格式错误，仅支持 .xlsx 文件: {excel_path}")
            return []
        df = pd.read_excel(excel_path, sheet_name="devices").fillna("")
        devices = []
        valid_brands = list(DEVICE_TYPES.keys())
        command_map = load_inspection_commands(excel_path)
        if '是否加载批量巡检命令' not in df.columns:
            logging.warning("未找到 '是否加载批量巡检命令' 列，将不加载批量巡检命令。")
        for idx, row in df.iterrows():
            missing = [f for f in ["IP地址", "设备品牌", "密码"] if not row[f]]
            if missing:
                logging.warning(f"第{idx + 2}行缺失字段: {', '.join(missing)}")
                continue
            if row["设备品牌"] not in valid_brands:
                logging.warning(f"第{idx + 2}行无效品牌: {row['设备品牌']}")
                continue
            protocol = str(row.get("登录协议", "ssh")).lower()
            if protocol not in ["ssh", "telnet"]:
                logging.warning(f"第{idx + 2}行无效协议: {protocol}，使用SSH")
                protocol = "ssh"
            try:
                port = int(row["端口"]) if row["端口"] else DEFAULT_PORTS[protocol]
            except ValueError:
                port = DEFAULT_PORTS[protocol]
                logging.warning(f"第{idx + 2}行无效端口: {row['端口']}，使用{port}")
            try:
                timeout = int(row["超时时间"]) if row["超时时间"] else 30
            except ValueError:
                timeout = 30
                logging.warning(f"第{idx + 2}行超时时间值无效，使用默认超时时间: 30")
            load_batch_commands = row.get("是否加载批量巡检命令", "否").strip().lower() == "是"
            brand_commands = command_map.get(row["设备品牌"], []) if load_batch_commands else []
            special_commands = parse_commands(row.get("特殊命令", ""))
            brand = row["设备品牌"]
            logging.info(f"设备 {row['IP地址']} 的品牌识别为: {brand}")
            base_type = DEVICE_TYPES.get(brand, "autodetect")
            protocol_suffix = PROTOCOL_MAP.get(protocol, "")
            device_type = f"{base_type}{protocol_suffix}"
            # 处理迪普设备 Telnet 连接
            if brand == "迪普" and protocol == "telnet":
                if "dptech_os_telnet" in SUPPORTED_DEVICES:
                    device_type = "dptech_os_telnet"
                else:
                    device_type = "custom_dptech_telnet"
            elif protocol == "telnet":
                if f"{base_type}_telnet" in SUPPORTED_DEVICES:
                    device_type = f"{base_type}_telnet"
                else:
                    device_type = "autodetect_telnet"
            else:
                if device_type not in SUPPORTED_DEVICES:
                    device_type = "autodetect"
            device = {
                "host": row["IP地址"],
                "username": row["用户名"] or None,
                "password": row["密码"],
                "secret": row["特权密码"] or None,
                "port": port,
                "device_type": device_type,
                "login_protocol": protocol,
                "timeout": timeout,
                "commands": brand_commands + special_commands
            }
            devices.append(device)
        return devices
    except FileNotFoundError:
        logging.error(f"未找到 Excel 文件: {excel_path}")
        return []
    except PermissionError:
        logging.error(f"Excel 文件被占用，无法读取: {excel_path}")
        return []
    except Exception as e:
        import traceback
        logging.error(f"设备数据加载失败: {str(e)}\n{traceback.format_exc()}")
        return []

# 设备连接模块
def connect_device(device):
    retries = 0
    while retries < RETRY_CONFIG["max_retries"]:
        try:
            if device["device_type"] == "custom_dptech_telnet":
                logging.debug(f"尝试使用 Telnet 连接迪普设备 {device['host']}")
                tn = telnetlib.Telnet(device["host"], device["port"], device["timeout"])
                logging.debug(f"Telnet 连接建立，等待 Login 提示...")
                output = tn.read_until(b"Login: ", timeout=5).decode("ascii")
                logging.debug(f"收到 Login 提示: {output}")
                username = device["username"] if device["username"] else ""
                tn.write((username + "\r\n").encode("ascii"))
                logging.debug(f"已发送用户名: {username}")
                output = tn.read_until(b"Password: ", timeout=5).decode("ascii")
                logging.debug(f"收到 Password 提示: {output}")
                password = device["password"] if device["password"] else ""
                tn.write((password + "\r\n").encode("ascii"))
                logging.debug(f"已发送密码: {password}")
                # 处理二次认证（如特权密码）
                if device.get("secret"):
                    output = tn.read_until(b"Secondary Password: ", timeout=5).decode("ascii")
                    logging.debug(f"收到 Secondary Password 提示: {output}")
                    tn.write((device["secret"] + "\r\n").encode("ascii"))
                    logging.debug(f"已发送二次密码: {device['secret']}")
                output = tn.read_until(b'>', timeout=5).decode("ascii")
                logging.debug(f"登录后收到输出: {output}")
                if ">" in output:
                    logging.info(f"成功连接到迪普设备 {device['host']}")
                    return tn
                elif "Authentication failed" in output:
                    raise Exception("迪普设备认证失败：用户名或密码错误")
                else:
                    raise Exception("迪普设备 Telnet 登录失败")
            else:
                # 动态构建 ConnectHandler 的参数
                params = {
                    "device_type": device["device_type"],
                    "host": device["host"],
                    "port": device["port"],
                    "password": device["password"],
                    "timeout": device["timeout"]
                }
                if device["username"]:
                    params["username"] = device["username"]
                if device["secret"]:
                    params["secret"] = device["secret"]
                # 调试输出
                logging.info(f"尝试连接设备 {device['host']}，使用 device_type: {device['device_type']}")
                conn = ConnectHandler(**params)
                prompt = conn.find_prompt()
                # 跳过迪普设备的特权模式处理
                if device["device_type"] not in ["dptech_os_telnet", "custom_dptech_telnet"]:
                    if prompt.endswith(">") and device["secret"]:
                        if "cisco" in device["device_type"]:
                            try:
                                conn.enable()
                                logging.info(f"{device['host']} 特权模式无密码进入成功")
                            except Exception:
                                conn.enable(password=device["secret"])
                                logging.info(f"{device['host']} 特权模式密码验证成功")
                        else:
                            conn.enable(password=device["secret"])
                return conn
        except NetmikoAuthenticationException as e:
            logging.error(f"{device['host']} 认证失败: {str(e)}，第 {retries + 1} 次尝试")
        except NetmikoTimeoutException:
            logging.error(f"{device['host']} 连接超时，第 {retries + 1} 次尝试")
        except Exception as e:
            logging.error(f"{device['host']} 连接失败: {str(e)}，第 {retries + 1} 次尝试")
        retries += 1
        time.sleep(RETRY_CONFIG["retry_interval"])
    logging.error(f"{device['host']} 经过 {RETRY_CONFIG['max_retries']} 次尝试后仍无法连接")
    return None

# 巡检执行模块（关键优化：分页提示过滤）
def execute_inspection(device, result_dir):
    host = device["host"]
    timestamp = time.strftime("%Y%m%d-%H%M%S")
    inspection_time = time.strftime("%Y-%m-%d %H:%M:%S")
    conn = connect_device(device)
    if not conn:
        error_path = os.path.join(result_dir, "errors")
        try:
            os.makedirs(error_path, exist_ok=True)
            error_filename = f"{host}_{timestamp}.error.log"
            with open(os.path.join(error_path, error_filename), "w", encoding="utf-8") as f:
                f.write(f"设备连接失败，无法获取设备名称。")
            logging.error(f"设备 {host} 连接失败，跳过巡检")
        except Exception as e:
            logging.error(f"创建错误日志文件失败: {str(e)}")
        return False

    try:
        command_outputs = []
        if isinstance(conn, telnetlib.Telnet):
            tn = conn
            tn.write(b'\r\n')
            time.sleep(0.5)
            output = tn.read_until(b'>', timeout=10).decode('ascii', errors='replace')
            device_name = re.search(r'<\s*(\S+)\s*>', output).group(1) or "未知设备"

            for command in device["commands"]:
                tn.write((command + "\r\n").encode("ascii"))
                time.sleep(0.3)
                output_buffer = []
                # 禁用分页后无需循环翻页，直接读取完整输出
                current_output = tn.read_until(b'>', timeout=30)  # 延长单次读取超时
                output_buffer.append(current_output)
                command_output = b''.join(output_buffer).decode('ascii', errors='replace')
                # 增强清理规则：去除分页提示及多余空格
                command_output = re.sub(r'--More\(CTRL\+C break\)--.*?\r?\n?', '', command_output, flags=re.DOTALL)
                command_output = re.sub(r'\x08|\x1b\[.*?m', '', command_output)  # 保留原有控制符清理
                command_output = re.sub(rf'^{re.escape(command)}\s*\r?\n', '', command_output)
                command_output = re.sub(rf'{re.escape(device_name)}[>#]\s*$', '', command_output).strip()
                command_outputs.append((command, command_output))
        else:
            # 非迪普设备保持原有逻辑
            prompt = conn.find_prompt()
            device_name = re.search(r'^(\S+)[>#]', prompt).group(1) or "未知设备"
            if "hp_comware" in device["device_type"]:
                try:
                    output = conn.send_command("display current-configuration | include sysname", read_timeout=10)
                    match = re.search(r"sysname (\S+)", output)
                    if match:
                        device_name = match.group(1)
                except Exception:
                    pass
            for cmd in device["commands"]:
                output = conn.send_command(cmd)
                output = re.sub(rf'^{re.escape(cmd)}\s*\r?\n', '', output)
                output = re.sub(rf'{re.escape(device_name)}[>#]\s*$', '', output).strip()
                command_outputs.append((cmd, output))

        # 保存巡检结果
        safe_device_name = re.sub(r'[\\/*?:"<>|]', '_', device_name)
        device_dir = os.path.join(result_dir, f"{host}__{safe_device_name}")
        os.makedirs(device_dir, exist_ok=True)
        report_filename = os.path.join(device_dir, f"{timestamp}.txt")
        with open(report_filename, "w", encoding="utf-8") as f:
            f.write(f"=== 设备巡检报告 ===\n")
            f.write(f"设备 IP: {host}\n")
            f.write(f"设备名称: {device_name}\n")
            f.write(f"巡检时间: {inspection_time}\n")
            f.write(f"登录协议: {device['login_protocol']}\n")
            f.write(f"=== 巡检命令输出 ===\n\n")
            separator = '#' * 40 + '\n'
            for command, output in command_outputs:
                f.write(separator)
                f.write(f"--- 命令: {command} ---\n")
                f.write(f"{output}\n\n")
            f.write(separator)

        logging.info(f"设备 {host} 巡检完成，报告已保存到 {report_filename}")
        return True
    except Exception as e:
        error_path = os.path.join(result_dir, "errors")
        os.makedirs(error_path, exist_ok=True)
        error_filename = f"{host}_{timestamp}.error.log"
        with open(os.path.join(error_path, error_filename), "w", encoding="utf-8") as f:
            f.write(f"巡检过程中发生错误: {str(e)}")
        logging.error(f"设备 {host} 巡检失败: {str(e)}")
        return False
    finally:
        if conn:
            if isinstance(conn, telnetlib.Telnet):
                conn.close()
            else:
                conn.disconnect()

# 开始巡检
def start_inspection(input_file, result_dir):
    setup_logging()
    clean_old_files(result_dir)
    devices = load_devices(input_file)
    if not devices:
        messagebox.showerror("错误", "未找到有效的设备信息，请检查输入文件。")
        return

    with ThreadPoolExecutor() as executor:
        futures = [executor.submit(execute_inspection, device, result_dir) for device in devices]
        for future in as_completed(futures):
            future.result()

    play_sound()
    messagebox.showinfo("完成", "巡检任务已完成。")

# 浏览文件
def browse_file(entry):
    file_path = filedialog.askopenfilename(filetypes=[("Excel 文件", "*.xlsx")])
    if file_path:
        entry.delete(0, tk.END)
        entry.insert(0, file_path)

# 主程序入口
if __name__ == "__main__":
    root = tk.Tk()
    root.title("网络设备自动化巡检系统")

    frame = tk.Frame(root, padx=20, pady=20)
    frame.pack()

    input_label = tk.Label(frame, text="输入文件:")
    input_label.grid(row=0, column=0, padx=5, pady=5)

    input_entry = tk.Entry(frame, width=50)
    input_entry.grid(row=0, column=1, padx=5, pady=5)

    browse_button = tk.Button(frame, text="浏览", command=lambda: browse_file(input_entry))
    browse_button.grid(row=0, column=2, padx=5, pady=5)

    start_button = tk.Button(frame, text="开始运行", command=lambda: start_inspection(input_entry.get(), "巡检结果"))
    start_button.grid(row=1, column=1, padx=5, pady=20)

    root.mainloop()
    

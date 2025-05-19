import openpyxl
from netmiko import ConnectHandler, exceptions
from concurrent.futures import ThreadPoolExecutor, as_completed
import datetime
import os
import sys
from tqdm import tqdm
import time
import threading
import colorama  # 添加颜色支持

# 初始化colorama以支持Windows终端颜色
colorama.init()

# 定义颜色常量
COLORS = {
    'GREEN': colorama.Fore.GREEN,
    'RED': colorama.Fore.RED,
    'YELLOW': colorama.Fore.YELLOW,
    'BLUE': colorama.Fore.BLUE,
    'RESET': colorama.Fore.RESET,
    'CYAN': colorama.Fore.CYAN,
}

# 状态显示锁
status_lock = threading.Lock()

# 设备状态字典
device_status = {}

def format_timedelta(seconds):
    """将秒数转换为HH:MM:SS格式"""
    return str(datetime.timedelta(seconds=int(seconds)))

def update_status(ip, status, message="", color=None):
    """更新设备状态显示"""
    with status_lock:
        device_status[ip] = (status, message, color)
        # 清屏并重新显示所有状态
        os.system('cls' if os.name == 'nt' else 'clear')
        print("\n======= 命令批量下发工具1.2 - 执行状态 =======\n")
        
        # 显示设备状态表格
        print(f"{COLORS['CYAN']}{'IP地址':<15} {'状态':<10} {'详细信息':<40}{COLORS['RESET']}")
        print("-" * 65)
        
        for dev_ip, (dev_status, dev_msg, dev_color) in sorted(device_status.items()):
            color_code = dev_color if dev_color else COLORS['RESET']
            print(f"{color_code}{dev_ip:<15} {dev_status:<10} {dev_msg:<40}{COLORS['RESET']}")

def get_user_input():
    print("\n======= NetCBDTool 1.2 =======")
    print("   by:wangtz1 Mail:zvrz@163.com \n")
    
    filename = input("1.请输入文件名，回车确定。不指定文件名将执行（批量配置模板.xlsx）: ").strip()
    if not filename:
        filename = "批量配置模板.xlsx"

    while True:
        threads = input("2.输入需要多少线程并行执行任务。输入回车代表按照单线程顺序执行: ").strip()
        if not threads:
            return filename, 1
        try:
            return filename, min(200, int(threads))
        except ValueError:
            print("错误：请输入有效数字")

def safe_strip(value):
    return str(value).strip() if value is not None else ''

def read_command_file(file_path):
    """读取命令文件内容"""
    try:
        # 如果不是绝对路径，则视为相对于当前目录的路径
        if not os.path.isabs(file_path):
            file_path = os.path.join(os.getcwd(), file_path)
            
        with open(file_path, 'r', encoding='utf-8') as f:
            return [line.strip() for line in f if line.strip()]
    except Exception as e:
        print(f"读取命令文件 {file_path} 失败：{str(e)}")
        return []

def read_devices(filename):
    try:
        wb = openpyxl.load_workbook(filename)
        ws = wb.active
        devices = []
        valid_count = 0  
        error_count = 0
        
        for row_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
            try:
                ip = safe_strip(row[0].value)
                if not ip:
                    continue

                port_value = safe_strip(row[1].value)
                try:
                    port = int(port_value) if port_value else 22
                except ValueError:
                    raise ValueError(f"无效端口号：'{port_value}'")

                username = safe_strip(row[2].value)
                password = safe_strip(row[3].value)
                if not username or not password:
                    raise ValueError("用户名或密码为空")

                device_type = safe_strip(row[5].value)
                if not device_type:
                    raise ValueError("设备类型为空")

                raw_commands = safe_strip(row[6].value)
                commands = []
                
                # 处理每一行的命令
                for line in raw_commands.split('\n'):
                    line = line.strip()
                    if not line:
                        continue
                    
                    # 如果行以.txt结尾，视为文件路径
                    if line.lower().endswith('.txt'):
                        file_commands = read_command_file(line)
                        commands.extend(file_commands)
                    else:
                        commands.append(line)

                devices.append({
                    'ip': ip,
                    'port': port,
                    'username': username,
                    'password': password,
                    'secret': safe_strip(row[4].value),
                    'device_type': device_type,
                    'commands': commands,
                })
                valid_count +=1
                
            except Exception as e:
                error_count += 1
                print(f"第{row_idx}行错误：{str(e)}")

        print(f"\n[设备统计] 检测到有效设备{valid_count}台")  # 新输出格式
        if error_count > 0:
            print(f"[数据警告] 发现{error_count}行格式错误数据")
            
        return devices if devices else None
    
    except FileNotFoundError:
        print(f"错误：文件 {filename} 不存在")
        return None
    except Exception as e:
        print(f"文件读取失败：{str(e)}")
        return None

def worker(device, thread_num, log_folder, timestamp):
    ip_clean = device['ip']
    log_filename = f"线程{thread_num}_{ip_clean}_{timestamp}.txt"  # 文件名格式调整
    log_path = os.path.join(log_folder, log_filename)
    log_content = []
    error_info = None  # 新增错误信息记录
    
    # 初始化设备状态
    update_status(device['ip'], "准备中", "等待连接...", COLORS['BLUE'])

    try:
        # 更新状态为连接中
        update_status(device['ip'], "连接中", f"尝试连接 {device['ip']}:{device['port']}...", COLORS['YELLOW'])
        
        conn = ConnectHandler(
            ip=device['ip'],
            port=device['port'],
            username=device['username'],
            password=device['password'],
            secret=device['secret'] or None,
            device_type=device['device_type'],
            conn_timeout=30
        )

        update_status(device['ip'], "已连接", "认证成功，准备执行命令", COLORS['GREEN'])

        if device['secret']:
            conn.enable()

        # 执行命令
        for i, cmd in enumerate(device['commands']):
            cmd_display = (cmd[:37] + '...') if len(cmd) > 40 else cmd
            update_status(device['ip'], "执行中", f"命令 {i+1}/{len(device['commands'])}: {cmd_display}", COLORS['CYAN'])
            
            output = conn.send_command_timing(cmd, strip_command=False)
            log_content.append(f"{conn.find_prompt()} {cmd}\n{output}\n{'='*60}\n")

        # 发送确认命令
        update_status(device['ip'], "完成中", "发送确认命令...", COLORS['CYAN'])
        for _ in range(3):
            output = conn.send_command_timing("")
            log_content.append(output)

        conn.disconnect()
        log_content.insert(0, f"[Success] {device['ip']} 配置成功\n\n")
        update_status(device['ip'], "成功", "所有命令执行完成", COLORS['GREEN'])
        
    except exceptions.NetmikoAuthenticationException as e:
        error_info = {
            'ip': device['ip'],
            'port': device['port'],
            'username': device['username'],
            'device_type': device['device_type'],
            'error_type': '认证失败',
            'error_msg': str(e)
        }
        log_content = [f"[Error] {device['ip']} 认证失败：用户名/密码错误\n"]
    except exceptions.NetmikoTimeoutException as e:
        error_info = {
            'ip': device['ip'],
            'port': device['port'],
            'username': device['username'],
            'device_type': device['device_type'],
            'error_type': '连接超时',
            'error_msg': str(e)
        }
        log_content = [f"[Error] {device['ip']} 连接超时：端口{device['port']}不可达\n"]
    except Exception as e:
        error_info = {
            'ip': device['ip'],
            'port': device['port'],
            'username': device['username'],
            'device_type': device['device_type'],
            'error_type': '未知错误',
            'error_msg': str(e)
        }
        log_content = [f"[Error] {device['ip']} 发生未知错误：{str(e)}\n"]

    with open(log_path, 'w', encoding='utf-8') as f:
        f.writelines(log_content)

    return thread_num, log_filename, error_info

def main_loop():
    while True:
        # 清空设备状态
        device_status.clear()
        
        filename, max_workers = get_user_input()
        
        # 读取设备列表
        devices = read_devices(filename)
        if not devices:
            print("未读取到有效设备信息！")
            choice = input("\n🔄 是否继续执行？（Y/n）: ").lower()
            if choice == 'n':
                break
            continue

        # 创建日志目录
        timestamp = datetime.datetime.now().strftime("%y%m%d_%H%M%S")
        log_folder = f"执行结果_{timestamp}"
        os.makedirs(log_folder, exist_ok=True)

        # 执行任务
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            start_time = time.time()
            completed = 0
            total = len(devices)
            
            # 优化进度条配置
            with tqdm(
                total=total,
                desc="总体进度",
                unit="台",
                bar_format="{l_bar}{bar}| {n_fmt}/{total_fmt} [已用:{elapsed} 剩余预估:{remaining}]",
                dynamic_ncols=True,
                smoothing=0.9,
                position=0,  # 固定进度条位置
                leave=True   # 保留进度条
            ) as progress:
                
                # 提交任务
                futures = {executor.submit(worker, device, i+1, log_folder, timestamp): i 
                          for i, device in enumerate(devices)}
                
                # 处理结果
                log_files = []
                error_devices = []
                time_records = []
                
                for future in as_completed(futures):
                    thread_num, log_filename, error_info = future.result()
                    exec_time = time.time() - start_time  # 使用全局时间简化计算
                    
                    # 更新统计
                    completed += 1
                    time_records.append(exec_time)
                    avg_time = sum(time_records)/completed if completed else 0
                    remaining = format_timedelta(avg_time * (total - completed)) if completed else "00:00:00"
                    
                    # 动态更新进度条
                    progress.set_postfix_str(f"成功:{completed-len(error_devices)} 失败:{len(error_devices)} 剩余预估:{remaining}")
                    progress.update(1)
                    
                    # 处理日志
                    log_files.append(log_filename)
                    if error_info:
                        error_devices.append(error_info)

            # 清理资源
            progress.close()

        # 清屏并显示最终结果
        os.system('cls' if os.name == 'nt' else 'clear')
        print("\n======= 命令批量下发工具1.2 - 执行结果 =======\n")
        
        # 显示设备最终状态表格
        print(f"{COLORS['CYAN']}{'IP地址':<15} {'状态':<10} {'详细信息':<40}{COLORS['RESET']}")
        print("-" * 65)
        
        for dev_ip, (dev_status, dev_msg, dev_color) in sorted(device_status.items()):
            color_code = dev_color if dev_color else COLORS['RESET']
            print(f"{color_code}{dev_ip:<15} {dev_status:<10} {dev_msg:<40}{COLORS['RESET']}")

        # 生成登录失败汇总日志
        if error_devices:
            summary_log = os.path.join(log_folder, f"登录失败汇总_{timestamp}.txt")
            with open(summary_log, 'w', encoding='utf-8') as f:
                f.write("=== 登录失败设备汇总 ===\n\n")
                f.write(f"总失败设备数：{len(error_devices)}\n\n")
                for idx, dev in enumerate(error_devices, 1):
                    f.write(f"{idx}. IP地址：{dev['ip']}\n")
                    f.write(f"   端口：{dev['port']}\n")
                    f.write(f"   用户名：{dev['username']}\n")
                    f.write(f"   设备类型：{dev['device_type']}\n")
                    f.write(f"   错误类型：{dev['error_type']}\n")
                    f.write(f"   错误信息：{dev['error_msg']}\n")
                    f.write("-"*60 + "\n\n")

        # 生成聚合日志
        aggregate_log = os.path.join(log_folder, f"聚合日志_{timestamp}.txt")
        with open(aggregate_log, 'w', encoding='utf-8') as agg_f:
            for log_file in log_files:
                file_path = os.path.join(log_folder, log_file)
                with open(file_path, 'r', encoding='utf-8') as f:
                    agg_f.write(f"[[ {log_file} ]]\n")
                    agg_f.writelines(f.readlines())
                    agg_f.write("\n\n")

        # 修改后的完成提示
        print(f"\n{COLORS['GREEN']}操作完成！成功下发设备：{len(devices)-len(error_devices)}台{COLORS['RESET']}")
        if error_devices:
            summary_log = os.path.join(log_folder, f"登录失败汇总_{timestamp}.txt")
            print(f"{COLORS['RED']}登录失败设备：{len(error_devices)}台，详见：{os.path.abspath(summary_log)}{COLORS['RESET']}")
        print(f"日志目录：{os.path.abspath(log_folder)}")

        # 交互继续
        choice = input("\n🔄 是否继续执行？（Y/n）: ").lower()
        if choice == 'n':
            print("程序退出")
            break
        os.system('cls' if os.name == 'nt' else 'clear')

if __name__ == "__main__":
    if getattr(sys, 'frozen', False):
        os.chdir(os.path.dirname(sys.executable))
    main_loop()
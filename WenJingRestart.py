import os
import sys
import threading
import time
from tkinter import messagebox, filedialog
import customtkinter as ctk
import pyautogui
import psutil
import configparser

# 兼容打包后的路径获取
if getattr(sys, 'frozen', False):
    # 如果是打包后的 exe 运行
    SCRIPT_DIR = os.path.dirname(sys.executable)
else:
    # 如果是源代码运行
    SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

CONFIG_FILE = os.path.join(SCRIPT_DIR, "wenJingRestartConfig.ini")

# 性能优化：关闭 pyautogui 的失败安全提示（避免弹窗阻塞挂机）
pyautogui.FAILSAFE = False 
# 优化：鼠标移动速度设为最快（0秒），减少操作耗时
pyautogui.PAUSE = 0.1 

class ProcessGuardian:
    def __init__(self):
        self.config = configparser.ConfigParser()
        self.config.optionxform = str  # 保持键名大小写
        self.load_config()
        self.running = False
        self.main_thread = None
        self.stop_event = threading.Event()

    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    if not f.read().strip():
                        self.create_default_config()
                        return
                self.config.read(CONFIG_FILE, encoding='utf-8')
                if 'Target' not in self.config:
                    self.create_default_config()
            except:
                self.create_default_config()
        else:
            self.create_default_config()

    def create_default_config(self):
        self.config['Target'] = {
            'process_name': '',  # 用户填写的进程名关键词
            'exe_path': ''       # 完整启动路径
        }
        self.config['Coordinates'] = {
            'login_x': '0', 'login_y': '0',
            'switch_x': '0', 'switch_y': '0',
            'select_all_x': '0', 'select_all_y': '0',
            'product_x': '0', 'product_y': '0'
        }
        self.config['Settings'] = {
            'interval_sec': '60',
            'startup_wait': '15'   # 启动后等待时间
        }
        self.save_config()

    def save_config(self):
        try:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                self.config.write(f)
        except Exception as e:
            print(f"保存配置失败：{e}")

    def get_config_safe(self, section, option, default):
        try:
            if self.config.has_section(section) and self.config.has_option(section, option):
                val = self.config.get(section, option)
                return val if val else default
            return default
        except:
            return default

    def is_process_running(self, keyword):
        """
        高性能进程检测：
        只要进程名中包含 keyword (不区分大小写)，即视为运行中。
        """
        if not keyword:
            return False
        
        target = keyword.lower()
        try:
            for proc in psutil.process_iter(['name']):
                try:
                    name = proc.info['name']
                    if name and target in name.lower():
                        return True
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    continue
        except Exception:
            pass
        return False

    def kill_process(self, keyword):
        """清理残留进程"""
        if not keyword:
            return
        target = keyword.lower()
        killed = False
        try:
            for proc in psutil.process_iter(['name', 'pid']):
                try:
                    name = proc.info['name']
                    if name and target in name.lower():
                        proc.kill()
                        killed = True
                        print(f"已终止残留进程：{name}")
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    continue
        except Exception:
            pass
        if killed:
            time.sleep(1) # 给系统一点时间释放资源

    def start_app(self, exe_path):
        """启动应用程序"""
        if not exe_path or not os.path.exists(exe_path):
            print(f"❌ 启动失败：文件不存在 -> {exe_path}")
            return False
        try:
            os.startfile(exe_path)
            print(f"🚀 正在启动程序：{os.path.basename(exe_path)}")
            return True
        except Exception as e:
            print(f"❌ 启动异常：{e}")
            return False

    def execute_flow(self):
        """执行全套点击流程 (绝对坐标)"""
        coords = self.config['Coordinates']
        steps = [
            ("登录", coords.get('login_x'), coords.get('login_y')),
            ("切换列表", coords.get('switch_x'), coords.get('switch_y')),
            ("全选", coords.get('select_all_x'), coords.get('select_all_y')),
            ("成品", coords.get('product_x'), coords.get('product_y'))
        ]
        
        # 获取等待时间
        wait_time = 15
        try:
            wait_time = int(self.get_config_safe('Settings', 'startup_wait', '15'))
        except:
            pass
        print("--- 开始执行初始化流程 ---")
        
        # 第一步：等待软件启动完成
        print(f"⏳ 等待软件初始化 ({wait_time}秒)...")
        # 使用分段等待，以便随时响应停止信号
        for _ in range(wait_time * 10):
            if self.stop_event.is_set():
                return
            time.sleep(0.1)

        # 后续步骤
        actions = [
            (steps[0], 15), # 登录后等待久一点
            (steps[1], 1),
            (steps[2], 1),
            (steps[3], 1)
        ]

        for (name, x, y), delay in actions:
            if self.stop_event.is_set():
                print("⚠️ 流程被用户中断")
                return
            
            if not x or not y:
                print(f"⚠️ 跳过 {name} (坐标未设置)")
                continue
            try:
                ix, iy = int(x), int(y)
                print(f"🖱️ 点击 [{name}] -> 绝对坐标：({ix}, {iy})")
                
                # 1. 移动鼠标
                pyautogui.moveTo(ix, iy, duration=0.3) 
                
                # 2. 【新增】移动到位后，强制等待 1 秒，防止电脑卡顿导致点击位置偏差
                time.sleep(1) 
                
                # 3. 执行点击
                pyautogui.click()
                
                # 分段等待 (步骤之间的等待)
                for _ in range(delay * 10):
                    if self.stop_event.is_set():
                        return
                    time.sleep(0.1)
            except Exception as e:
                print(f"❌ 执行 {name} 时出错：{e}")
        
        print("--- 初始化流程结束，进入监控状态 ---")

    def start_monitoring(self):
        if self.running:
            return
        self.running = True
        self.stop_event.clear()
        self.main_thread = threading.Thread(target=self._monitor_loop, daemon=True)
        self.main_thread.start()

    def stop_monitoring(self):
        self.running = False
        self.stop_event.set()
        if self.main_thread and self.main_thread.is_alive():
            self.main_thread.join(timeout=2)

    def _monitor_loop(self):
        """核心监控循环"""
        keyword = self.get_config_safe('Target', 'process_name', '')
        exe_path = self.get_config_safe('Target', 'exe_path', '')
        interval_str = self.get_config_safe('Settings', 'interval_sec', '60')
        
        try:
            interval = int(interval_str)
        except:
            interval = 60
            
        if not keyword:
            print("❌ 错误：未设置进程名关键词，监控无法启动。")
            return
            
        print(f"✅ 监控已启动 | 目标进程关键词：'{keyword}' | 检测间隔：{interval}秒")
        
        while self.running and not self.stop_event.is_set():
            # 1. 检查进程
            if not self.is_process_running(keyword):
                print(f"\n[{time.strftime('%H:%M:%S')}] ⚠️ 检测到进程消失！判定为闪退。")
                
                # 2. 清理残留
                self.kill_process(keyword)
                
                # 3. 重启
                if self.start_app(exe_path):
                    # 4. 执行流程
                    self.execute_flow()
                else:
                    print("⚠️ 启动失败，将在下一个周期重试...")
                    time.sleep(5)
            else:
                # 进程正常，静默等待
                # 采用分段睡眠，确保停止信号能即时响应
                for _ in range(interval * 10):
                    if self.stop_event.is_set():
                        break
                    time.sleep(0.1)

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("文镜挂机助手")
        self.geometry("800x650")
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        self.guardian = ProcessGuardian()
        self.setup_ui()
        self.update_display()

    def setup_ui(self):
        main = ctk.CTkScrollableFrame(self)
        main.pack(fill="both", expand=True, padx=20, pady=20)

        # 标题
        ctk.CTkLabel(main, text="更新文镜后务必核对取点坐标！！！", font=ctk.CTkFont(size=24, weight="bold")).pack(pady=(0, 5))
        ctk.CTkLabel(main, text="💡 原理：仅监控进程存活 + 绝对坐标点击 | 极低占用 | 支持最小化挂机", 
                     text_color="#00ff00", font=ctk.CTkFont(size=12)).pack(pady=(0, 20))

        # --- 配置区域 ---
        config_frame = ctk.CTkFrame(main)
        config_frame.pack(fill="x", pady=10)

        # 进程名
        f1 = ctk.CTkFrame(config_frame, fg_color="transparent")
        f1.pack(fill="x", padx=10, pady=8)
        ctk.CTkLabel(f1, text="进程名关键词:", width=120).pack(side="left")
        self.proc_entry = ctk.CTkEntry(f1, placeholder_text="例：wenjing 或 main (无需.exe)")
        self.proc_entry.pack(side="left", fill="x", expand=True, padx=10)
        ctk.CTkButton(f1, text="测试检测", command=self.test_process, width=100).pack(side="right")
        ctk.CTkLabel(main, text="🔍 支持模糊匹配，只要文件名包含该词即可", text_color="gray", font=ctk.CTkFont(size=10)).pack(anchor="w", padx=135)

        # 程序路径
        f2 = ctk.CTkFrame(config_frame, fg_color="transparent")
        f2.pack(fill="x", padx=10, pady=8)
        ctk.CTkLabel(f2, text="程序完整路径:", width=120).pack(side="left")
        self.path_entry = ctk.CTkEntry(f2, placeholder_text="选择 .exe 文件用于重启")
        self.path_entry.pack(side="left", fill="x", expand=True, padx=10)
        ctk.CTkButton(f2, text="浏览...", command=self.browse_path, width=80).pack(side="right")

        # 设置行
        f3 = ctk.CTkFrame(config_frame, fg_color="transparent")
        f3.pack(fill="x", padx=10, pady=8)
        ctk.CTkLabel(f3, text="监控间隔(秒):", width=120).pack(side="left")
        self.interval_entry = ctk.CTkEntry(f3, width=80)
        self.interval_entry.pack(side="left", padx=10)
        
        ctk.CTkLabel(f3, text="启动等待(秒):", width=100).pack(side="left", padx=(20, 0))
        self.wait_entry = ctk.CTkEntry(f3, width=80)
        self.wait_entry.pack(side="left", padx=10)

        # --- 坐标区域 ---
        ctk.CTkLabel(main, text="⚙️ 绝对坐标设置 (请在软件默认位置获取)", font=ctk.CTkFont(weight="bold", size=16)).pack(pady=(20, 10))
        ctk.CTkLabel(main, text="注意：软件重启后会自动回到默认位置，请确保在此位置下获取坐标", 
                     text_color="orange", font=ctk.CTkFont(size=11)).pack(pady=(0, 10))
        
        coord_frame = ctk.CTkFrame(main)
        coord_frame.pack(fill="both", expand=True, pady=5)
        self.coord_entries = {}
        labels = [("login", "登录"), ("switch", "切换列表"), ("select_all", "全选"), ("product", "成品")]
        
        for key, label in labels:
            row = ctk.CTkFrame(coord_frame, fg_color="transparent")
            row.pack(fill="x", padx=20, pady=4)
            
            ctk.CTkLabel(row, text=label, width=100).pack(side="left")
            
            x_ent = ctk.CTkEntry(row, width=100, placeholder_text="X")
            x_ent.pack(side="left", padx=10)
            y_ent = ctk.CTkEntry(row, width=100, placeholder_text="Y")
            y_ent.pack(side="left", padx=10)
            
            btn = ctk.CTkButton(row, text="📍 获取当前鼠标坐标", width=160, 
                                command=lambda k=key: self.capture_coord(k))
            btn.pack(side="left", padx=10)
            
            self.coord_entries[key] = {'x': x_ent, 'y': y_ent}

        # --- 控制区域 ---
        ctrl_frame = ctk.CTkFrame(main)
        ctrl_frame.pack(fill="x", pady=20)
        
        self.start_btn = ctk.CTkButton(ctrl_frame, text="▶ 开始守护", command=self.start_guardian, 
                                       fg_color="green", hover_color="darkgreen", height=40, font=ctk.CTkFont(size=16))
        self.start_btn.pack(side="left", padx=20)
        
        self.stop_btn = ctk.CTkButton(ctrl_frame, text="⏹ 停止", command=self.stop_guardian, 
                                      state="disabled", fg_color="red", hover_color="darkred", height=40, font=ctk.CTkFont(size=16))
        self.stop_btn.pack(side="left", padx=20)
        
        self.status_label = ctk.CTkLabel(ctrl_frame, text="● 空闲", font=ctk.CTkFont(size=14, weight="bold"))
        self.status_label.pack(side="left", padx=20)

        # --- 日志区域 ---
        ctk.CTkLabel(main, text="📋 运行日志", anchor="w").pack(fill="x", padx=5)
        self.log_box = ctk.CTkTextbox(main, height=150)
        self.log_box.pack(fill="both", expand=True, pady=5)
        
        # 重定向打印
        class LogRedirector:
            def __init__(self, textbox):
                self.textbox = textbox
            def write(self, msg):
                if msg.strip():
                    self.textbox.insert("end", msg + "\n")
                    self.textbox.see("end")
            def flush(self):
                pass
        sys.stdout = LogRedirector(self.log_box)

    def update_display(self):
        cfg = self.guardian
        self.proc_entry.delete(0, 'end')
        self.proc_entry.insert(0, cfg.get_config_safe('Target', 'process_name', ''))
        
        self.path_entry.delete(0, 'end')
        self.path_entry.insert(0, cfg.get_config_safe('Target', 'exe_path', ''))
        
        self.interval_entry.delete(0, 'end')
        self.interval_entry.insert(0, cfg.get_config_safe('Settings', 'interval_sec', '60'))
        
        self.wait_entry.delete(0, 'end')
        self.wait_entry.insert(0, cfg.get_config_safe('Settings', 'startup_wait', '15'))
        
        for key, ents in self.coord_entries.items():
            ents['x'].delete(0, 'end')
            ents['x'].insert(0, cfg.get_config_safe('Coordinates', f'{key}_x', '0'))
            ents['y'].delete(0, 'end')
            ents['y'].insert(0, cfg.get_config_safe('Coordinates', f'{key}_y', '0'))

    def test_process(self):
        kw = self.proc_entry.get().strip()
        if not kw:
            messagebox.showwarning("提示", "请输入进程名关键词")
            return
        
        print(f"🔍 正在搜索包含 '{kw}' 的进程...")
        if self.guardian.is_process_running(kw):
            print(f"✅ 成功！检测到目标进程正在运行。")
        else:
            print(f"❌ 未找到包含 '{kw}' 的进程。请确保软件已启动。")
        
        # 【新增】测试成功后，将当前进程名关键词立即保存到配置文件
        if kw:  # 确保输入不为空
            if not self.guardian.config.has_section('Target'):
                self.guardian.config.add_section('Target')
            
            self.guardian.config.set('Target', 'process_name', kw)
            
            try:
                with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                    self.guardian.config.write(f)
                print(f"✅ 进程名关键词已自动保存至配置文件: {kw}")
            except Exception as e:
                print(f"❌ 保存配置文件失败: {e}")

    def browse_path(self):
        path = filedialog.askopenfilename(title="选择程序主程序", filetypes=[("Executable", "*.exe")])
        if path:
            # 1. 更新界面显示
            self.path_entry.delete(0, 'end')
            self.path_entry.insert(0, path)
            
            # 2. 【新增】立即更新配置对象并保存到文件
            # 确保配置对象已加载 (通常 __init__ 里已经 load 了)
            if not self.guardian.config.has_section('Target'):
                self.guardian.config.add_section('Target')
            
            # 将新路径写入内存中的 config 对象
            self.guardian.config.set('Target', 'exe_path', path)
            
            # 立即写入磁盘文件
            try:
                with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                    self.guardian.config.write(f)
                print(f"✅ 路径已自动保存至配置文件: {path}")
            except Exception as e:
                print(f"❌ 保存配置文件失败: {e}")
            
            # 自动填充进程名建议
            suggested_name = os.path.basename(path).replace('.exe', '')
            if not self.proc_entry.get():
                self.proc_entry.delete(0, 'end')
                self.proc_entry.insert(0, suggested_name)
                print(f"💡 已自动填入进程名关键词：{suggested_name}")

    def capture_coord(self, key):
        print(f"👆 请将鼠标移动到目标位置，3秒后自动记录 {key} 坐标...")
        countdown = 3
        for i in range(countdown, 0, -1):
            print(f"   {i}...")
            time.sleep(1)
        
        x, y = pyautogui.position()
        
        # 1. 更新界面显示
        self.coord_entries[key]['x'].delete(0, 'end')
        self.coord_entries[key]['x'].insert(0, str(x))
        self.coord_entries[key]['y'].delete(0, 'end')
        self.coord_entries[key]['y'].insert(0, str(y))
        
        # 2. 【新增】立即更新配置对象并保存到文件
        if not self.guardian.config.has_section('Coordinates'):
            self.guardian.config.add_section('Coordinates')
        
        self.guardian.config.set('Coordinates', f'{key}_x', str(x))
        self.guardian.config.set('Coordinates', f'{key}_y', str(y))
        
        try:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                self.guardian.config.write(f)
            print(f"✅ 已记录并保存 {key} 坐标：({x}, {y})")
        except Exception as e:
            print(f"❌ 保存配置文件失败: {e}")

    def start_guardian(self):
        # 保存配置
        cfg = self.guardian.config
        proc_name = self.proc_entry.get().strip()
        if not proc_name:
            messagebox.showerror("错误", "必须填写进程名关键词！")
            return
            
        cfg.set('Target', 'process_name', proc_name)
        cfg.set('Target', 'exe_path', self.path_entry.get().strip())
        cfg.set('Settings', 'interval_sec', self.interval_entry.get().strip() or '60')
        cfg.set('Settings', 'startup_wait', self.wait_entry.get().strip() or '15')
        
        for key, ents in self.coord_entries.items():
            cfg.set('Coordinates', f'{key}_x', ents['x'].get().strip() or '0')
            cfg.set('Coordinates', f'{key}_y', ents['y'].get().strip() or '0')
            
        self.guardian.save_config()
        
        # 更新 UI
        self.start_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        self.status_label.configure(text="● 运行中 (监控进程)", text_color="#00ff00")
        
        self.guardian.start_monitoring()
        print("\n" + "="*30)
        print("🛡️ 守护进程已启动")
        print("="*30)

    def stop_guardian(self):
        self.guardian.stop_monitoring()
        self.start_btn.configure(state="normal")
        self.stop_btn.configure(state="disabled")
        self.status_label.configure(text="● 已停止", text_color="red")
        print("\n⏹ 守护已停止")

    def destroy(self):
        self.guardian.stop_monitoring()
        sys.stdout = sys.__stdout__
        super().destroy()

if __name__ == "__main__":
    app = App()
    app.mainloop() 

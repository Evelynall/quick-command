import json
import os
import time
import tkinter as tk
from tkinter import ttk, Frame, messagebox
from ttkthemes import ThemedTk
import pystray
import pyautogui
import keyboard
import threading
from PIL import Image
import pyperclip

CONFIG_FILE = "button_config.json"
HOTKEY_CONFIG = "hotkey_config.json"


class DraggableButton(ttk.Button):
    """支持拖动排序的按钮组件（保持主题一致性）"""

    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.drag_start_pos = (0, 0)
        self.is_dragging = False
        self.click_time = 0
        self._is_closing = False  # 新增关闭状态标志
        self.valid_click = True
        self.data_index = -1  # 新增属性用于存储对应配置索引


class MainApplication:
    def __init__(self):
        # 使用支持主题的窗口
        self.root = ThemedTk(theme="arc")
        self.root.title("快捷指令-Evelynal")

        # 初始化热键相关变量
        self.hotkey = "shift+e"
        self.hotkey_handler = None
        self.load_hotkey_config()

        # 设置主窗口居中
        window_width = 300
        window_height = 450
        self.center_window(self.root, window_width, window_height)

        icon_path = "icon.ico"
        icon = tk.PhotoImage(file=icon_path)
        self.root.iconphoto(True, icon)

        # 设置全局按键延迟
        pyautogui.PAUSE = 0

        # 初始化设置窗口引用
        self.settings_window = None

        # 初始化样式系统
        self.init_styles()

        # 初始化数据存储
        self.button_data = []
        self.page_scrollable_frames = []
        self.page_canvas = []
        self.load_config()

        # 拖动状态
        self.drag_source = None
        self.drag_placeholder = None
        self.drag_switch_var = tk.BooleanVar(value=False)

        # 系统托盘
        self.setup_tray()

        # 界面构建
        self.create_scrollable_ui()
        self.create_action_buttons()
        self.refresh_current_page_buttons()

        # 事件绑定
        self.root.protocol('WM_DELETE_WINDOW', self.hide_to_tray)
        self.register_hotkey()
        self.setup_tab_context_menu()  # 添加这行初始化右键菜单

    def safe_tkinter_operation(func):
        """防止在组件销毁后执行UI操作的装饰器"""

        def wrapper(self, *args, **kwargs):
            if not hasattr(self, 'root') or not self.root.winfo_exists():
                return
            try:
                return func(self, *args, **kwargs)
            except tk.TclError as e:
                if "bad window path name" in str(e):
                    return
                raise
        return wrapper

    def async_safe(func):
        """确保异步操作中不访问已销毁的组件"""

        def wrapper(self, *args, **kwargs):
            if not hasattr(self, '_is_closing') or self._is_closing:
                return
            return func(self, *args, **kwargs)
        return wrapper

    def center_window(self, window, width, height):
        """窗口居中函数"""
        screen_width = window.winfo_screenwidth()
        screen_height = window.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        window.geometry(f"{width}x{height}+{x}+{y}")

    def init_styles(self):
        """初始化所有组件样式"""
        self.style = ttk.Style()
        self.style.configure(".", background="#f5f6f8", font=('微软雅黑', 9))
        self.style.configure("TButton", padding=6, width=12)
        self.style.configure("TScrollbar", gripcount=0)
        self.style.configure(
            "Dragging.TButton", background="#e0e0e0", relief="sunken", borderwidth=2)
        self.style.map("TButton",
                       background=[('active', '#e2e6ea'),
                                   ('pressed', '#dae0e5')],
                       relief=[('pressed', 'sunken'), ('!pressed', 'flat')])

    def setup_tray(self):
        """系统托盘设置"""
        menu = pystray.Menu(
            pystray.MenuItem('显示', self.show_main_window),
            pystray.MenuItem('设置', self.show_settings),
            pystray.MenuItem('退出', self.exit_app)
        )
        icon_path = "icon.ico"
        image = Image.open(icon_path)
        self.tray_icon = pystray.Icon("name", image, "快捷指令-Evelynal", menu)
        threading.Thread(target=self.tray_icon.run, daemon=True).start()

    def create_scrollable_ui(self):
        """创建分页的可滚动界面"""
        main_container = ttk.Frame(self.root, padding=10)
        main_container.pack(fill=tk.BOTH, expand=True)

        # 创建Notebook分页组件
        self.notebook = ttk.Notebook(main_container)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        # 根据加载的配置创建页面
        for page in self.button_data:
            self.add_page_ui(page["page_name"])

        # 如果没有页面则创建默认页
        if not self.button_data:
            self.button_data.append({"page_name": "默认页", "buttons": []})
            self.add_page_ui("默认页")

        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_changed)

    def add_page_ui(self, page_name):
        """添加新页面的UI组件"""
        page_frame = ttk.Frame(self.notebook)
        self.notebook.add(page_frame, text=page_name)

        canvas = tk.Canvas(page_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(
            page_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        scrollable_frame_id = canvas.create_window(
            (0, 0), window=scrollable_frame, anchor="nw",
            width=canvas.winfo_width()
        )

        def on_canvas_configure(event):
            canvas.itemconfigure(scrollable_frame_id, width=event.width)
            self.refresh_current_page_buttons()

        canvas.bind('<Configure>', on_canvas_configure)
        scrollable_frame.bind("<Configure>",
                              lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        self.page_scrollable_frames.append(scrollable_frame)
        self.page_canvas.append(canvas)

    @async_safe
    def on_tab_changed(self, event):
        """标签页切换事件处理（增加窗口存在性检查）"""
        if not self.root.winfo_exists():
            return
        self.refresh_current_page_buttons()

    def create_action_buttons(self):
        """创建底部操作按钮（移除了管理页面按钮）"""
        bottom_frame = ttk.Frame(self.root)
        bottom_frame.pack(side='bottom', fill=tk.X, padx=10, pady=5)

        ttk.Checkbutton(
            bottom_frame,
            text="启用拖动排序",
            variable=self.drag_switch_var,
            style="Toggle.TCheckbutton"
        ).pack(side='left', padx=5)

        ttk.Button(
            bottom_frame,
            text="+ 添加按钮",
            style="Accent.TButton",
            command=self.show_add_dialog
        ).pack(side='right', padx=5)

    def setup_tab_context_menu(self):
        """设置分页标签右键菜单"""
        self.notebook.bind("<Button-3>", self.on_tab_right_click)

    def on_tab_right_click(self, event):
        """分页标签右键点击事件处理"""
        # 获取点击的标签索引
        try:
            tab_index = self.notebook.index(f"@{event.x},{event.y}")
        except tk.TclError:
            return

        if tab_index >= 0:
            self.show_tab_context_menu(event.x_root, event.y_root, tab_index)

    def show_tab_context_menu(self, x, y, tab_index):
        """显示分页右键菜单"""
        menu = tk.Menu(self.root, tearoff=0)

        # 当前页面操作
        menu.add_command(
            label="重命名页面",
            command=lambda: self.show_rename_page_dialog(tab_index)
        )
        menu.add_command(
            label="删除页面",
            command=lambda: self.delete_page(tab_index)
        )

        # 全局操作
        menu.add_separator()
        menu.add_command(
            label="添加新页面",
            command=self.show_add_page_dialog
        )

        menu.tk_popup(x, y)

    def show_page_management(self):
        """显示页面管理对话框"""
        dialog = tk.Toplevel(self.root)
        dialog.title("管理页面")
        self.center_window(dialog, 250, 200)

        ttk.Button(dialog, text="添加新页面",
                   command=lambda: self.show_add_page_dialog(dialog)).pack(pady=5)
        ttk.Button(dialog, text="删除当前页面",
                   command=lambda: self.delete_current_page(dialog)).pack(pady=5)
        ttk.Button(dialog, text="重命名当前页面",
                   command=lambda: self.show_rename_page_dialog(dialog)).pack(pady=5)

    def show_add_page_dialog(self):
        """显示添加新页面对话框"""
        dialog = tk.Toplevel(self.root)
        dialog.title("添加页面")
        self.center_window(dialog, 250, 100)

        ttk.Label(dialog, text="页面名称：").pack(pady=5)
        name_entry = ttk.Entry(dialog)
        name_entry.pack(pady=5)

        def add_page():
            page_name = name_entry.get().strip()
            if not page_name:
                messagebox.showwarning("错误", "页面名称不能为空")
                return
            if any(page["page_name"] == page_name for page in self.button_data):
                messagebox.showwarning("错误", "页面名称已存在")
                return
            self.button_data.append({"page_name": page_name, "buttons": []})
            self.add_page_ui(page_name)
            self.save_config()
            dialog.destroy()

        ttk.Button(dialog, text="添加", command=add_page).pack(pady=5)

    def delete_page(self, tab_index):
        """删除指定页面"""
        if len(self.button_data) <= 1:
            messagebox.showerror("错误", "至少需要保留一个页面")
            return

        page_name = self.button_data[tab_index]["page_name"]
        if messagebox.askyesno("确认删除", f"确定删除页面 '{page_name}' 吗？"):
            # 删除数据和UI组件
            del self.button_data[tab_index]
            self.notebook.forget(tab_index)
            del self.page_scrollable_frames[tab_index]
            del self.page_canvas[tab_index]
            self.save_config()
            self.refresh_current_page_buttons()

    def show_rename_page_dialog(self, tab_index):
        """显示重命名页面对话框"""
        old_name = self.button_data[tab_index]["page_name"]

        dialog = tk.Toplevel(self.root)
        dialog.title("重命名页面")
        self.center_window(dialog, 250, 100)

        ttk.Label(dialog, text="新页面名称：").pack(pady=5)
        name_entry = ttk.Entry(dialog)
        name_entry.insert(0, old_name)
        name_entry.pack(pady=5)

        def rename():
            new_name = name_entry.get().strip()
            if not new_name:
                messagebox.showwarning("错误", "页面名称不能为空")
                return
            if any(page["page_name"] == new_name for page in self.button_data):
                messagebox.showwarning("错误", "页面名称已存在")
                return
            self.button_data[tab_index]["page_name"] = new_name
            self.notebook.tab(tab_index, text=new_name)
            self.save_config()
            dialog.destroy()

        ttk.Button(dialog, text="重命名", command=rename).pack(pady=5)

    def get_current_page_index(self):
        """获取当前活动页面的索引"""
        try:
            return self.notebook.index("current")
        except:
            return None

    def get_scrollable_frame(self):
        """安全获取当前页面的滚动框架"""
        current_index = self.get_current_page_index()
        if (
            current_index is None or
            current_index >= len(self.page_scrollable_frames) or
            not self.page_scrollable_frames[current_index].winfo_exists()
        ):
            return None
        return self.page_scrollable_frames[current_index]

    @safe_tkinter_operation
    def refresh_current_page_buttons(self):
        """刷新当前活动页面的按钮（增加双重保护检查）"""
        # 先检查主窗口是否有效
        if not self.root.winfo_exists():
            return

        current_index = self.get_current_page_index()
        if current_index is None:
            return

        # 二次检查页面容器是否存在
        try:
            scrollable_frame = self.page_scrollable_frames[current_index]
            if not scrollable_frame.winfo_exists():  # 关键检查
                return
        except (IndexError, tk.TclError):
            return

        # 清除现有按钮
        for widget in scrollable_frame.winfo_children():
            if isinstance(widget, DraggableButton):
                widget.destroy()

        # 获取当前页面的按钮数据
        buttons = self.button_data[current_index]["buttons"]

        # 计算列数
        available_width = scrollable_frame.winfo_width(
        ) or self.page_canvas[current_index].winfo_width()
        columns = max(1, available_width // 100)

        # 重新创建按钮
        for idx, btn_data in enumerate(buttons):
            btn = DraggableButton(
                scrollable_frame,
                text=btn_data["name"],
                command=lambda cmd=btn_data["command"]: self.safe_execute(
                    cmd, btn),
                style="TButton"
            )
            btn.data_index = idx

            # 计算行列位置
            row = idx // columns
            col = idx % columns
            btn.grid(row=row, column=col, padx=4, pady=4, sticky="ew")

            # 绑定事件
            btn.bind("<ButtonPress-1>", lambda e,
                     b=btn: self.on_drag_start(e, b))
            btn.bind("<B1-Motion>", self.on_drag_motion)
            btn.bind("<ButtonRelease-1>", self.on_drag_end)
            btn.bind("<Button-3>", self.on_right_click)

        # 更新布局
        scrollable_frame.update_idletasks()

    @safe_tkinter_operation
    def get_current_page_index(self):
        """获取当前活动页面的索引（增加容错处理）"""
        try:
            # 检查notebook是否已被销毁
            if not hasattr(self, 'notebook') or not self.notebook.winfo_exists():
                return None
            return self.notebook.index("current")
        except (AttributeError, tk.TclError):
            return None

    def on_drag_start(self, event, button):
        if not self.drag_switch_var.get():
            return
        button.click_time = time.time()
        button.drag_start_pos = (event.x_root, event.y_root)
        button.after_id = button.after(200, self.start_dragging, button)

    def start_dragging(self, button):
        button.is_dragging = True
        self.drag_source = button
        self.create_placeholder(button)
        button.configure(style="Dragging.TButton")

    def on_drag_motion(self, event):
        """处理拖动排序时的按钮交换逻辑"""
        if not self.drag_source or not self.drag_source.is_dragging:
            return

        current_index = self.get_current_page_index()
        if current_index is None:
            return

        scrollable_frame = self.page_scrollable_frames[current_index]
        buttons = self.button_data[current_index]["buttons"]

        # 获取所有按钮的位置信息
        button_positions = {
            btn: (btn.grid_info()["row"], btn.grid_info()["column"])
            for btn in scrollable_frame.winfo_children()
            if isinstance(btn, DraggableButton)
        }

        # 计算目标位置
        frame_x = scrollable_frame.winfo_rootx()
        frame_width = scrollable_frame.winfo_width()
        columns = max(1, frame_width // 100)

        if columns == 0 or frame_width == 0:
            return

        column_width = frame_width / columns
        x_in_frame = event.x_root - frame_x
        target_col = int(x_in_frame // column_width)
        target_col = max(0, min(columns - 1, target_col))

        y_in_frame = event.y_root - scrollable_frame.winfo_rooty()
        row_height = self.drag_source.winfo_height() + 4
        target_row = int(y_in_frame // row_height)

        # 找到目标按钮
        target_btn = next(
            (btn for btn, (r, c) in button_positions.items()
             if r == target_row and c == target_col),
            None
        )

        if target_btn and target_btn != self.drag_source:
            # 交换数据和位置
            src_index = self.drag_source.data_index
            tgt_index = target_btn.data_index

            buttons[src_index], buttons[tgt_index] = buttons[tgt_index], buttons[src_index]

            # 更新按钮索引
            self.drag_source.data_index = tgt_index
            target_btn.data_index = src_index

            # 交换布局位置
            src_row, src_col = button_positions[self.drag_source]
            tgt_row, tgt_col = button_positions[target_btn]
            self.drag_source.grid(row=tgt_row, column=tgt_col)
            target_btn.grid(row=src_row, column=src_col)

            self.save_config()

    def on_drag_end(self, event):
        if self.drag_source:
            self.drag_source.after_cancel(
                getattr(self.drag_source, 'after_id', None))
            self.drag_source.configure(style="TButton")
            self.remove_placeholder()
            self.drag_source.is_dragging = False
            self.drag_source = None

    def create_placeholder(self, button):
        self.drag_placeholder = Frame(button.master,
                                      height=button.winfo_height(),
                                      width=button.winfo_width(),
                                      bg="#e0e0e0")
        row = button.grid_info()["row"]
        col = button.grid_info()["column"]
        self.drag_placeholder.grid(row=row, column=col, padx=2, pady=2)

    def remove_placeholder(self):
        if self.drag_placeholder:
            self.drag_placeholder.destroy()
            self.drag_placeholder = None

    def load_config(self):
        """加载配置文件（统一使用UTF-8编码）"""
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    data = json.load(f)

                    # 验证配置文件格式
                    if not isinstance(data, list):
                        raise ValueError("配置文件格式错误")

                    # 清理无效数据
                    valid_data = []
                    for page in data:
                        if "page_name" in page and "buttons" in page:
                            valid_buttons = []
                            for btn in page["buttons"]:
                                if "name" in btn and "command" in btn:
                                    valid_buttons.append({
                                        "name": btn["name"],
                                        "command": btn["command"]
                                    })
                            valid_data.append({
                                "page_name": page["page_name"],
                                "buttons": valid_buttons
                            })

                    if not valid_data:
                        raise ValueError("没有有效页面数据")

                    self.button_data = valid_data
            except Exception as e:
                messagebox.showerror("配置错误",
                                     f"配置文件加载失败，已重置为默认配置\n错误信息：{str(e)}")
                self.button_data = [{"page_name": "默认页", "buttons": []}]
                self.save_config()
        else:
            self.button_data = [{"page_name": "默认页", "buttons": []}]

    def save_config(self):
        """保存配置（统一使用UTF-8编码）"""
        try:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump([{
                    "page_name": page["page_name"],
                    "buttons": [{
                        "name": btn["name"],
                        "command": btn["command"]
                    } for btn in page["buttons"]]
                } for page in self.button_data], f, ensure_ascii=False, indent=2)
        except Exception as e:
            messagebox.showerror("保存失败", f"无法保存配置：{str(e)}")

    def safe_execute(self, command, button):
        if self.drag_switch_var.get():
            return
        if button.valid_click and not button.is_dragging:
            self.execute_command(command)
        button.valid_click = True

    def show_add_dialog(self):
        """显示添加按钮对话框"""
        current_index = self.get_current_page_index()
        if current_index is None:
            return

        dialog = tk.Toplevel(self.root)
        dialog.title("添加新指令")
        self.center_window(dialog, 250, 150)

        ttk.Label(dialog, text="按钮名称：").grid(row=0, column=0, padx=5, pady=5)
        name_entry = ttk.Entry(dialog)
        name_entry.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(dialog, text="执行指令：").grid(row=1, column=0, padx=5, pady=5)
        cmd_entry = ttk.Entry(dialog)
        cmd_entry.grid(row=1, column=1, padx=5, pady=5)

        def add_button():
            name = name_entry.get().strip()
            command = cmd_entry.get().strip()
            if name and command:
                self.button_data[current_index]["buttons"].append({
                    "name": name,
                    "command": command
                })
                self.save_config()
                self.refresh_current_page_buttons()
                dialog.destroy()
            else:
                messagebox.showwarning("输入错误", "按钮名称和执行指令不能为空")

        ttk.Button(dialog, text="确认添加", command=add_button).grid(
            row=2, column=1, pady=10)

    def execute_command(self, command):
        try:
            self.root.withdraw()
            pyautogui.hotkey('esc')
            time.sleep(0.05)
            pyautogui.press('/')
            time.sleep(0.1)
            pyperclip.copy(command)
            pyautogui.hotkey('ctrl', 'v')
            pyautogui.press('enter')
        except Exception as e:
            messagebox.showerror("执行错误", f"指令发送失败：{str(e)}")

    def on_right_click(self, event):
        """右键菜单处理"""
        btn = event.widget
        if not isinstance(btn, DraggableButton):
            return

        current_index = self.get_current_page_index()
        if current_index is None:
            return

        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(
            label="修改", command=lambda: self.edit_button(btn, current_index))
        menu.add_command(
            label="删除", command=lambda: self.delete_button(btn, current_index))
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def edit_button(self, button, page_index):
        """修改按钮"""
        btn_index = button.data_index
        if btn_index < 0 or btn_index >= len(self.button_data[page_index]["buttons"]):
            messagebox.showerror("错误", "找不到对应的按钮配置")
            return

        current_data = self.button_data[page_index]["buttons"][btn_index]

        dialog = tk.Toplevel(self.root)
        dialog.title("修改按钮")
        self.center_window(dialog, 300, 180)

        ttk.Label(dialog, text="新名称：").grid(row=0, column=0, padx=5, pady=5)
        name_entry = ttk.Entry(dialog)
        name_entry.insert(0, current_data["name"])
        name_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        ttk.Label(dialog, text="新指令：").grid(row=1, column=0, padx=5, pady=5)
        cmd_entry = ttk.Entry(dialog)
        cmd_entry.insert(0, current_data["command"])
        cmd_entry.grid(row=1, column=1, padx=5, pady=5, sticky="ew")

        def save_changes():
            new_name = name_entry.get().strip()
            new_cmd = cmd_entry.get().strip()
            if not new_name or not new_cmd:
                messagebox.showwarning("输入错误", "名称和指令都不能为空")
                return
            self.button_data[page_index]["buttons"][btn_index] = {
                "name": new_name,
                "command": new_cmd
            }
            self.save_config()
            self.refresh_current_page_buttons()
            dialog.destroy()

        btn_frame = ttk.Frame(dialog)
        btn_frame.grid(row=2, column=0, columnspan=2, pady=10)
        ttk.Button(btn_frame, text="保存", command=save_changes).pack(
            side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="取消", command=dialog.destroy).pack(
            side=tk.LEFT, padx=5)

    def delete_button(self, button, page_index):
        """删除按钮"""
        btn_index = button.data_index
        if btn_index < 0 or btn_index >= len(self.button_data[page_index]["buttons"]):
            messagebox.showerror("错误", "找不到对应的按钮配置")
            return

        if messagebox.askyesno("确认删除", f"确定要删除按钮 [{button.cget('text')}] 吗？"):
            del self.button_data[page_index]["buttons"][btn_index]
            self.save_config()
            self.refresh_current_page_buttons()

    def show_settings(self):
        self.root.after(0, self._create_settings_window)

    def _create_settings_window(self):
        if self.settings_window and self.settings_window.winfo_exists():
            self.settings_window.lift()
            return

        self.settings_window = tk.Toplevel(self.root)
        self.settings_window.title("设置")
        self.center_window(self.settings_window, 330, 200)

        container = ttk.Frame(self.settings_window, padding=15)
        container.pack(fill=tk.BOTH, expand=True)

        # 热键设置组件
        hotkey_frame = ttk.Frame(container)
        hotkey_frame.pack(fill=tk.X, pady=5)

        ttk.Label(hotkey_frame, text="主窗口快捷键:").pack(side=tk.LEFT)
        self.hotkey_entry = ttk.Entry(hotkey_frame)
        self.hotkey_entry.insert(0, self.hotkey)
        self.hotkey_entry.pack(side=tk.LEFT, padx=5)

        ttk.Button(
            hotkey_frame,
            text="保存",
            command=self.save_hotkey_setting
        ).pack(side=tk.LEFT)

        self.settings_window.protocol(
            "WM_DELETE_WINDOW", self._on_settings_close)

    def save_hotkey_setting(self):
        """保存热键设置（添加未修改判断）"""
        new_hotkey = self.hotkey_entry.get().strip().lower()  # 统一小写并去除空格
        current_hotkey = self.hotkey.lower()

        # 判断是否修改
        if new_hotkey == current_hotkey:
            messagebox.showinfo("提示", "快捷键未修改")
            return

        # 保留原热键用于恢复
        old_hotkey = self.hotkey
        self.hotkey = new_hotkey

        try:
            # 测试新热键有效性
            test_handler = keyboard.add_hotkey(new_hotkey, lambda: None)
            keyboard.remove_hotkey(test_handler)
        except Exception as e:
            # 恢复原有热键
            self.hotkey = old_hotkey
            self.hotkey_entry.delete(0, tk.END)
            self.hotkey_entry.insert(0, old_hotkey)
            messagebox.showerror("无效热键", f"热键设置失败: {str(e)}")
            return

        # 仅当有修改时执行保存和注册
        self.save_hotkey_config()
        self.register_hotkey()
        messagebox.showinfo("保存成功", "热键设置已更新！")

    def _on_settings_close(self):
        if self.settings_window:
            self.settings_window.destroy()
            self.settings_window = None

    def show_main_window(self):
        self.root.deiconify()
        self.root.attributes('-topmost', 1)
        self.root.after_idle(self.root.attributes, '-topmost', 0)

    def hide_to_tray(self):
        self.root.withdraw()

    def exit_app(self):
        """退出程序时增加销毁顺序控制"""
        self._is_closing = True  # 标记正在关闭

        # 先解除事件绑定
        if hasattr(self, 'notebook'):
            self.notebook.unbind("<<NotebookTabChanged>>")

        # 停止系统托盘
        if hasattr(self, 'tray_icon'):
            self.tray_icon.stop()

        # 销毁子组件
        if hasattr(self, 'notebook'):
            for child in self.notebook.winfo_children():
                try:
                    child.destroy()
                except tk.TclError:
                    pass
            try:
                self.notebook.destroy()
            except tk.TclError:
                pass

        # 延迟主窗口销毁
        if hasattr(self, 'root') and self.root.winfo_exists():
            self.root.after(100, self.root.destroy)

    def run(self):
        self.root.mainloop()

    def load_hotkey_config(self):
        """加载热键配置"""
        if os.path.exists(HOTKEY_CONFIG):
            try:
                with open(HOTKEY_CONFIG, 'r') as f:
                    config = json.load(f)
                    self.hotkey = config.get("hotkey", "shift+e")
            except Exception as e:
                messagebox.showerror("加载失败", f"热键配置文件错误：{str(e)}")

    def save_hotkey_config(self):
        """保存热键配置"""
        try:
            with open(HOTKEY_CONFIG, 'w') as f:
                json.dump({"hotkey": self.hotkey}, f, indent=2)
        except Exception as e:
            messagebox.showerror("保存失败", f"无法保存热键配置：{str(e)}")

    def register_hotkey(self):
        """注册全局热键"""
        if self.hotkey_handler:
            keyboard.remove_hotkey(self.hotkey_handler)
        try:
            self.hotkey_handler = keyboard.add_hotkey(
                self.hotkey, self.show_main_window)
        except ValueError as e:
            messagebox.showerror(
                "热键错误", f"无效热键配置: {self.hotkey}，将恢复默认值\n错误信息: {str(e)}")
            self.hotkey = "shift+e"
            self.save_hotkey_config()
            self.hotkey_handler = keyboard.add_hotkey(
                self.hotkey, self.show_main_window)


if __name__ == "__main__":
    app = MainApplication()
    app.run()

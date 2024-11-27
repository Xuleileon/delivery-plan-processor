import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
from main import process_delivery_plan
import ctypes

# 启用DPI感知
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except:
    pass

class StatusFrame(tk.Frame):
    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)
        self.configure(bg=parent["bg"])
        
        # 状态标签
        self.label = tk.Label(
            self,
            text="等待处理文件...",
            font=('Microsoft YaHei UI', 10),
            bg=parent["bg"],
            fg="#666666"
        )
        self.label.pack(pady=10)
        
    def update_status(self, message, is_error=False):
        color = "#d32f2f" if is_error else "#666666"
        self.label.configure(text=message, fg=color)
        
class DeliveryPlanProcessorGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("到货计划处理工具")
        
        # 设置最小窗口大小
        self.root.minsize(500, 400)
        
        # 获取屏幕DPI缩放因子
        try:
            scaling = ctypes.windll.shcore.GetScaleFactorForDevice(0) / 100
        except:
            scaling = self.root.winfo_fpixels('1i') / 72
            
        # 设置适中的窗口大小
        base_width = 600
        base_height = 500
        scaled_width = int(base_width * scaling)
        scaled_height = int(base_height * scaling)
        
        self.root.geometry(f"{scaled_width}x{scaled_height}")
        
        # 设置主题颜色
        self.primary_color = "#2196F3"    # 主色调
        self.bg_color = "#FAFAFA"         # 背景色
        self.text_color = "#333333"       # 文字颜色
        self.secondary_color = "#64B5F6"  # 次要色调
        
        # 设置窗口样式
        self.root.configure(bg=self.bg_color)
        
        # 设置ttk样式
        self.style = ttk.Style()
        self.style.configure(
            'Custom.TButton',
            font=('Microsoft YaHei UI', 11),
            padding=5
        )
        
        # 创建主框架
        self.main_frame = tk.Frame(self.root, bg=self.bg_color)
        self.main_frame.pack(expand=True, fill='both', padx=20, pady=20)
        
        # 初始化变量
        self.output_dir = None
        self.config_path = None
        
        self.create_widgets()
        self.center_window()
        
    def create_widgets(self):
        # 创建标题区域
        title_frame = tk.Frame(self.main_frame, bg=self.bg_color)
        title_frame.pack(fill='x', pady=(0, 15))
        
        # Logo区域
        logo_label = tk.Label(
            title_frame,
            text="",
            font=('Segoe UI Emoji', 32),
            bg=self.bg_color,
            fg=self.primary_color
        )
        logo_label.pack()
        
        # 主标题
        title_label = tk.Label(
            title_frame,
            text="到货计划处理工具",
            font=('Microsoft YaHei UI', 18, 'bold'),
            fg=self.text_color,
            bg=self.bg_color
        )
        title_label.pack(pady=(5, 3))
        
        # 副标题
        subtitle_label = tk.Label(
            title_frame,
            text="快速处理和转换您的Excel文件",
            font=('Microsoft YaHei UI Light', 10),
            fg="#666666",
            bg=self.bg_color
        )
        subtitle_label.pack()
        
        # 创建内容区域
        content_frame = tk.Frame(self.main_frame, bg=self.bg_color)
        content_frame.pack(fill='both', expand=True)
        
        # 创建选项区域
        options_frame = tk.LabelFrame(
            content_frame,
            text="选项设置",
            font=('Microsoft YaHei UI', 10),
            bg=self.bg_color,
            fg=self.text_color
        )
        options_frame.pack(fill='x', padx=10, pady=10)
        
        # 输出目录选择
        output_frame = tk.Frame(options_frame, bg=self.bg_color)
        output_frame.pack(fill='x', padx=10, pady=5)
        
        tk.Label(
            output_frame,
            text="输出目录：",
            font=('Microsoft YaHei UI', 10),
            bg=self.bg_color,
            fg=self.text_color
        ).pack(side='left')
        
        self.output_label = tk.Label(
            output_frame,
            text="(默认)",
            font=('Microsoft YaHei UI', 10),
            bg=self.bg_color,
            fg="#666666"
        )
        self.output_label.pack(side='left', padx=(0, 10))
        
        ttk.Button(
            output_frame,
            text="选择目录",
            command=self.select_output_dir,
            style='Custom.TButton'
        ).pack(side='right')
        
        # 配置文件选择
        config_frame = tk.Frame(options_frame, bg=self.bg_color)
        config_frame.pack(fill='x', padx=10, pady=5)
        
        tk.Label(
            config_frame,
            text="配置文件：",
            font=('Microsoft YaHei UI', 10),
            bg=self.bg_color,
            fg=self.text_color
        ).pack(side='left')
        
        self.config_label = tk.Label(
            config_frame,
            text="(默认)",
            font=('Microsoft YaHei UI', 10),
            bg=self.bg_color,
            fg="#666666"
        )
        self.config_label.pack(side='left', padx=(0, 10))
        
        ttk.Button(
            config_frame,
            text="选择配置",
            command=self.select_config_file,
            style='Custom.TButton'
        ).pack(side='right')
        
        # 数据源选择区域
        source_frame = tk.LabelFrame(
            content_frame,
            text="数据来源",
            font=('Microsoft YaHei UI', 10),
            bg=self.bg_color,
            fg=self.text_color
        )
        source_frame.pack(fill='x', padx=10, pady=10)
        
        # 按钮区域
        button_frame = tk.Frame(source_frame, bg=self.bg_color)
        button_frame.pack(pady=15)
        
        ttk.Button(
            button_frame,
            text="选择Excel文件",
            command=self.process_local_file,
            style='Custom.TButton',
            width=20
        ).pack(side='left', padx=10)
        
        ttk.Button(
            button_frame,
            text="从飞书获取",
            command=self.process_feishu_data,
            style='Custom.TButton',
            width=20
        ).pack(side='left', padx=10)
        
        # 状态显示区域
        self.status_frame = StatusFrame(content_frame, bg=self.bg_color)
        self.status_frame.pack(fill='x', pady=(10, 0))
        
        # 添加底部信息
        footer_label = tk.Label(
            self.main_frame,
            text=" 2024 powered by Rien",
            font=('Microsoft YaHei UI Light', 9),
            fg="#999999",
            bg=self.bg_color
        )
        footer_label.pack(side='bottom', pady=5)
        
    def select_output_dir(self):
        """选择输出目录"""
        dir_path = filedialog.askdirectory(title="选择输出目录")
        if dir_path:
            self.output_dir = dir_path
            self.output_label.config(text=Path(dir_path).name)
            
    def select_config_file(self):
        """选择配置文件"""
        file_path = filedialog.askopenfilename(
            title="选择配置文件",
            filetypes=[("YAML文件", "*.yaml"), ("所有文件", "*.*")]
        )
        if file_path:
            self.config_path = file_path
            self.config_label.config(text=Path(file_path).name)
            
    def process_local_file(self):
        """处理本地Excel文件"""
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
        )
        
        if not file_path:
            return
            
        self.process_file(file_path)
        
    def process_feishu_data(self):
        """从飞书获取并处理数据"""
        try:
            self.update_status("正在从飞书获取数据，请稍候...")
            self.process_file(None, is_feishu=True)
        except Exception as e:
            messagebox.showerror("错误", f"从飞书获取数据失败：{str(e)}")
            self.update_status(f"获取失败：{str(e)}", True)
            
    def process_file(self, file_path=None, is_feishu=False):
        """处理文件的主函数"""
        try:
            self.update_status("正在处理，请稍候...")
            
            # 准备参数
            input_source = None if is_feishu else file_path
            
            # 调用处理函数
            result = process_delivery_plan(
                input_source=input_source,
                output_dir=self.output_dir,
                config_path=self.config_path
            )
            
            if result['success']:
                message = (
                    f"处理成功！\n\n"
                    f"输出文件位置：\n"
                    f"1. 预处理文件：{Path(result['data']['preprocessed_file']).name}\n"
                    f"2. 转换后文件：{Path(result['data']['transformed_file']).name}\n"
                    f"3. 最终文件：{Path(result['data']['final_file']).name}\n\n"
                    f"所有文件已保存在{'选定的输出目录' if self.output_dir else '原始文件夹'}中。"
                )
                messagebox.showinfo("处理成功", message)
                self.update_status("处理完成")
            else:
                messagebox.showerror("处理失败", f"错误：{result['message']}")
                self.update_status(f"处理失败：{result['message']}", True)
                
        except Exception as e:
            messagebox.showerror("错误", f"处理过程中发生错误：{str(e)}")
            self.update_status(f"发生错误：{str(e)}", True)
            
    def center_window(self):
        """使窗口在屏幕中居中"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
        
    def update_status(self, message, is_error=False):
        """更新状态信息"""
        self.status_frame.update_status(message, is_error)
        
    def run(self):
        """运行GUI程序"""
        self.root.mainloop()

if __name__ == "__main__":
    app = DeliveryPlanProcessorGUI()
    app.run()

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
        self.root.minsize(400, 300)
        
        # 获取屏幕DPI缩放因子
        try:
            scaling = ctypes.windll.shcore.GetScaleFactorForDevice(0) / 100
        except:
            scaling = self.root.winfo_fpixels('1i') / 72
            
        # 设置适中的窗口大小
        base_width = 500
        base_height = 400
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
        
        # 文件选择按钮
        button_frame = tk.Frame(content_frame, bg=self.bg_color)
        button_frame.pack(expand=True)
        
        self.select_button = ttk.Button(
            button_frame,
            text="选择Excel文件",
            command=self.process_file,
            style='Custom.TButton',
            width=20
        )
        self.select_button.pack(pady=15)
        
        # 状态显示区域
        self.status_frame = StatusFrame(content_frame, bg=self.bg_color)
        self.status_frame.pack(fill='x', pady=(10, 0))
        
        # 添加底部信息
        footer_label = tk.Label(
            self.main_frame,
            text=" 2024 ",
            font=('Microsoft YaHei UI Light', 9),
            fg="#999999",
            bg=self.bg_color
        )
        footer_label.pack(side='bottom', pady=5)
        
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
        
    def process_file(self):
        """处理文件的主函数"""
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
        )
        
        if not file_path:
            return
            
        try:
            self.update_status("正在处理文件，请稍候...")
            
            result = process_delivery_plan(file_path)
            
            if result['success']:
                message = (
                    f"处理成功！\n\n"
                    f"输出文件位置：\n"
                    f"1. 预处理文件：{Path(result['data']['preprocessed_file']).name}\n"
                    f"2. 转换后文件：{Path(result['data']['transformed_file']).name}\n"
                    f"3. 最终文件：{Path(result['data']['final_file']).name}\n\n"
                    f"所有文件已保存在原始文件夹中。"
                )
                messagebox.showinfo("处理成功", message)
                self.update_status("处理完成")
            else:
                messagebox.showerror("处理失败", f"错误：{result['message']}")
                self.update_status(f"处理失败：{result['message']}", True)
                
        except Exception as e:
            messagebox.showerror("错误", f"处理过程中发生错误：{str(e)}")
            self.update_status(f"发生错误：{str(e)}", True)
            
    def run(self):
        """运行GUI程序"""
        self.root.mainloop()

if __name__ == "__main__":
    app = DeliveryPlanProcessorGUI()
    app.run()

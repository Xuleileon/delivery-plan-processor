import sys
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
from main import process_delivery_plan

class DeliveryPlanProcessorGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("到货计划处理工具")
        self.root.geometry("500x300")
        
        # 设置窗口样式
        self.root.configure(bg='#f0f0f0')
        
        # 创建主框架
        main_frame = tk.Frame(self.root, bg='#f0f0f0')
        main_frame.pack(expand=True, fill='both', padx=20, pady=20)
        
        # 标题
        title_label = tk.Label(
            main_frame,
            text="到货计划处理工具",
            font=('Arial', 16, 'bold'),
            bg='#f0f0f0'
        )
        title_label.pack(pady=20)
        
        # 说明文字
        instruction_label = tk.Label(
            main_frame,
            text="请选择要处理的Excel文件",
            font=('Arial', 10),
            bg='#f0f0f0'
        )
        instruction_label.pack(pady=10)
        
        # 选择文件按钮
        self.select_button = tk.Button(
            main_frame,
            text="选择文件",
            command=self.process_file,
            font=('Arial', 10),
            width=20,
            bg='#4CAF50',
            fg='white',
            relief=tk.RAISED
        )
        self.select_button.pack(pady=20)
        
        # 状态标签
        self.status_label = tk.Label(
            main_frame,
            text="",
            font=('Arial', 9),
            bg='#f0f0f0',
            wraplength=400
        )
        self.status_label.pack(pady=10)
        
    def process_file(self):
        """处理文件的主函数"""
        # 打开文件选择对话框
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
        )
        
        if not file_path:
            return
            
        try:
            # 更新状态
            self.status_label.config(text="正在处理文件，请稍候...", fg='#666666')
            self.root.update()
            
            # 处理文件
            result = process_delivery_plan(file_path)
            
            if result['success']:
                # 显示成功消息
                message = (
                    f"处理成功！\n\n"
                    f"输出文件位置：\n"
                    f"1. 预处理文件：{Path(result['data']['preprocessed_file']).name}\n"
                    f"2. 转换后文件：{Path(result['data']['transformed_file']).name}\n"
                    f"3. 最终文件：{Path(result['data']['final_file']).name}\n\n"
                    f"所有文件已保存在原始文件夹中。"
                )
                messagebox.showinfo("处理成功", message)
                self.status_label.config(text="处理完成", fg='green')
            else:
                # 显示错误消息
                messagebox.showerror("处理失败", f"错误：{result['message']}")
                self.status_label.config(text=f"处理失败：{result['message']}", fg='red')
                
        except Exception as e:
            # 显示错误消息
            messagebox.showerror("错误", f"处理过程中发生错误：{str(e)}")
            self.status_label.config(text=f"发生错误：{str(e)}", fg='red')
            
    def run(self):
        """运行GUI程序"""
        self.root.mainloop()

if __name__ == "__main__":
    app = DeliveryPlanProcessorGUI()
    app.run()

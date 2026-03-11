import os
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox

class DocGenGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("合同生成工具")
        self.root.geometry("400x300")
        
        # 设置默认的zip文件保存路径
        self.zip_output_path = os.getcwd()
        
        # 创建界面元素
        self.create_widgets()
    
    def create_widgets(self):
        # 标题
        title_label = tk.Label(self.root, text="合同生成工具", font=("Arial", 16, "bold"))
        title_label.pack(pady=20)
        
        # 按钮框架
        button_frame = tk.Frame(self.root)
        button_frame.pack(pady=10)
        
        # 修改word模板按钮
        modify_template_btn = tk.Button(button_frame, text="修改word模板", width=20, command=self.open_template_folder)
        modify_template_btn.pack(pady=5)
        
        # 填写合同信息按钮
        fill_info_btn = tk.Button(button_frame, text="填写合同信息", width=20, command=self.open_excel_file)
        fill_info_btn.pack(pady=5)
        
        # 生成合同按钮
        generate_contract_btn = tk.Button(button_frame, text="生成合同", width=20, command=self.generate_contract)
        generate_contract_btn.pack(pady=5)
        
        # 路径设置框架
        path_frame = tk.Frame(self.root)
        path_frame.pack(pady=20)
        
        # 路径标签
        path_label = tk.Label(path_frame, text="ZIP输出路径:")
        path_label.pack(side=tk.LEFT, padx=10)
        
        # 路径显示框
        self.path_var = tk.StringVar(value=self.zip_output_path)
        path_entry = tk.Entry(path_frame, textvariable=self.path_var, width=30)
        path_entry.pack(side=tk.LEFT, padx=10)
        
        # 浏览按钮
        browse_btn = tk.Button(path_frame, text="浏览", command=self.browse_output_path)
        browse_btn.pack(side=tk.LEFT, padx=10)
    
    def open_template_folder(self):
        # 打开包含所有docx文件的文件夹
        template_folder = os.getcwd()
        subprocess.Popen(f"explorer {template_folder}")
    
    def open_excel_file(self):
        # 打开input.xlsx文件
        excel_path = os.path.join(os.getcwd(), "input.xlsx")
        if os.path.exists(excel_path):
            os.startfile(excel_path)
        else:
            messagebox.showerror("错误", "input.xlsx文件不存在")
    
    def generate_contract(self):
        # 运行template_filler.py脚本
        script_path = os.path.join(os.getcwd(), "template_filler.py")
        excel_path = os.path.join(os.getcwd(), "input.xlsx")
        
        if not os.path.exists(script_path):
            messagebox.showerror("错误", "template_filler.py文件不存在")
            return
        
        if not os.path.exists(excel_path):
            messagebox.showerror("错误", "input.xlsx文件不存在")
            return
        
        # 构建命令
        command = f"python {script_path} --excel {excel_path} --zip-output {self.zip_output_path}"
        
        # 执行命令
        try:
            # 保存当前工作目录
            current_dir = os.getcwd()
            # 切换到zip输出路径
            os.chdir(self.zip_output_path)
            # 执行命令
            result = subprocess.run(command, shell=True, capture_output=True, text=True)
            # 切换回原目录
            os.chdir(current_dir)
            
            if result.returncode == 0:
                messagebox.showinfo("成功", "合同生成成功！")
            else:
                messagebox.showerror("错误", f"合同生成失败：{result.stderr}")
        except Exception as e:
            messagebox.showerror("错误", f"执行命令时出错：{str(e)}")
    
    def browse_output_path(self):
        # 浏览选择zip输出路径
        selected_path = filedialog.askdirectory(initialdir=self.zip_output_path)
        if selected_path:
            self.zip_output_path = selected_path
            self.path_var.set(selected_path)

if __name__ == "__main__":
    root = tk.Tk()
    app = DocGenGUI(root)
    root.mainloop()
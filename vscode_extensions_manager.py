import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import subprocess
import os
import json
from pathlib import Path
import winreg

class VSCodeExtensionManager:
    def __init__(self, root):
        self.root = root
        self.root.title("VS Code 扩展管理器")
        self.root.geometry("800x400")
        
        # 获取 VS Code 路径
        self.vscode_path = self.get_vscode_path()
        if not self.vscode_path:
            messagebox.showerror("错误", "未能找到 VS Code，请确保已安装 VS Code")
            root.destroy()
            return
            
        # 设置字体样式
        self.default_font = ('"Noto Sans SC"', 12)
        self.title_font = ('"Noto Sans SC"', 14)
        
        # 配置默认样式
        style = ttk.Style()
        style.configure("TLabel", font=self.default_font)
        style.configure("TButton", font=self.default_font)
        style.configure("TCheckbutton", font=self.default_font)
        style.configure("TEntry", font=self.default_font)
        
        # 创建主框架
        self.main_frame = ttk.Frame(root, padding="10")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 导出扩展部分
        ttk.Label(self.main_frame, text="导出扩展列表", font=self.title_font).grid(row=0, column=0, columnspan=2, pady=10, sticky=tk.W)
        
        ttk.Label(self.main_frame, text="导出路径:").grid(row=1, column=0, sticky=tk.W)
        self.export_path = ttk.Entry(self.main_frame, width=50, font=self.default_font)
        self.export_path.grid(row=1, column=1, padx=5)
        ttk.Button(self.main_frame, text="浏览", command=self.browse_export_path).grid(row=1, column=2)
        ttk.Button(self.main_frame, text="导出扩展列表", command=self.export_extensions).grid(row=1, column=3, padx=5)
        
        # 安装扩展部分
        ttk.Label(self.main_frame, text="安装扩展", font=self.title_font).grid(row=2, column=0, columnspan=2, pady=(20,10), sticky=tk.W)
        
        ttk.Label(self.main_frame, text="扩展列表文件:").grid(row=3, column=0, sticky=tk.W)
        self.install_path = ttk.Entry(self.main_frame, width=50, font=self.default_font)
        self.install_path.grid(row=3, column=1, padx=5)
        ttk.Button(self.main_frame, text="浏览", command=self.browse_install_path).grid(row=3, column=2)
        ttk.Button(self.main_frame, text="安装扩展", command=self.install_extensions).grid(row=3, column=3, padx=5)
        
        # 创建折叠面板
        self.create_collapsible_panel()
        
    def create_collapsible_panel(self):
        # 创建折叠面板框架
        self.panel_frame = ttk.Frame(self.root)
        self.panel_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 创建展开/折叠按钮
        self.is_expanded = tk.BooleanVar(value=False)
        self.toggle_btn = ttk.Checkbutton(
            self.panel_frame,
            text="查看命令行操作方法",
            variable=self.is_expanded,
            command=self.toggle_panel,
            style="TCheckbutton"
        )
        self.toggle_btn.grid(row=0, column=0, sticky=tk.W, padx=10)
        
        # 创建文本框
        self.text_area = tk.Text(self.panel_frame, height=10, width=70, font=self.default_font)
        self.text_area.grid(row=1, column=0, padx=10, pady=5)
        self.text_area.grid_remove()  # 初始状态隐藏
        
        # 添加说明文本
        help_text = """# VS Code 扩展管理命令行操作方法

## 导出已安装的扩展列表
```bash
code --list-extensions > extensions.txt
```

## 安装扩展列表中的所有扩展
```bash
for /f %i in (extensions.txt) do code --install-extension %i
```

注意：在批处理文件(.bat)中使用时，需要使用 %%i 而不是 %i
"""
        self.text_area.insert(tk.END, help_text)
        self.text_area.config(state='disabled')  # 设置为只读
        
    def toggle_panel(self):
        if self.is_expanded.get():
            self.text_area.grid()
        else:
            self.text_area.grid_remove()
            
    def browse_export_path(self):
        directory = filedialog.askdirectory()
        if directory:
            self.export_path.delete(0, tk.END)
            self.export_path.insert(0, os.path.join(directory, "extensions.txt"))
            
    def browse_install_path(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        if file_path:
            self.install_path.delete(0, tk.END)
            self.install_path.insert(0, file_path)
            
    def get_vscode_path(self):
        """获取 VS Code 可执行文件路径"""
        try:
            # 尝试从注册表获取 VS Code 安装路径
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Code.exe", 0, winreg.KEY_READ) as key:
                return winreg.QueryValue(key, None)
        except WindowsError:
            # 如果注册表中没有，尝试常见安装路径
            common_paths = [
                os.path.join(os.environ.get('LOCALAPPDATA', ''), 'Programs', 'Microsoft VS Code', 'bin', 'code.cmd'),
                os.path.join(os.environ.get('ProgramFiles', ''), 'Microsoft VS Code', 'bin', 'code.cmd'),
                os.path.join(os.environ.get('ProgramFiles(x86)', ''), 'Microsoft VS Code', 'bin', 'code.cmd'),
            ]
            
            for path in common_paths:
                if os.path.exists(path):
                    return path
                    
            return None
            
    def export_extensions(self):
        export_path = self.export_path.get()
        if not export_path:
            messagebox.showerror("错误", "请选择导出路径")
            return
            
        try:
            result = subprocess.run(
                [self.vscode_path, "--list-extensions"],
                capture_output=True,
                text=True
            )
            
            if result.returncode == 0:
                # 直接写入命令输出，保持原始格式
                with open(export_path, 'w', encoding='utf-8', newline='') as f:
                    f.write(result.stdout)
                messagebox.showinfo("成功", f"扩展列表已导出到:\n{export_path}")
            else:
                messagebox.showerror("错误", f"无法获取扩展列表\n错误信息: {result.stderr}")
        except Exception as e:
            error_msg = str(e)
            if "WinError 2" in error_msg:
                messagebox.showerror("错误", "无法找到 VS Code。请确保 VS Code 已正确安装。")
            else:
                messagebox.showerror("错误", f"导出过程中发生错误:\n{error_msg}")
            
    def install_extensions(self):
        install_path = self.install_path.get()
        if not install_path or not os.path.exists(install_path):
            messagebox.showerror("错误", "请选择有效的扩展列表文件")
            return
            
        try:
            # 创建临时批处理文件
            bat_path = os.path.join(os.path.dirname(install_path), "install_extensions.bat")
            with open(bat_path, 'w', encoding='utf-8') as f:
                f.write(f'@echo off\n')
                f.write(f'echo 正在安装 VS Code 扩展...\n')
                f.write(f'for /f "tokens=*" %%i in ({os.path.basename(install_path)}) do (\n')
                f.write(f'    echo 正在安装: %%i\n')
                f.write(f'    "{self.vscode_path}" --install-extension %%i\n')
                f.write(f')\n')
                f.write(f'echo 安装完成\n')
                f.write(f'pause\n')
            
            # 执行批处理文件
            result = subprocess.run(
                [bat_path],
                cwd=os.path.dirname(install_path),
                shell=True,
                text=True,
                stderr=subprocess.PIPE
            )
            
            # 删除临时批处理文件
            try:
                os.remove(bat_path)
            except:
                pass
                
            if result.returncode == 0:
                messagebox.showinfo("完成", "扩展安装已完成")
            else:
                messagebox.showerror("错误", f"安装过程中出现错误\n{result.stderr}")
                
        except Exception as e:
            error_msg = str(e)
            if "WinError 2" in error_msg:
                messagebox.showerror("错误", "无法找到 VS Code。请确保 VS Code 已正确安装。")
            else:
                messagebox.showerror("错误", f"安装过程中发生错误:\n{error_msg}")

if __name__ == "__main__":
    root = tk.Tk()
    app = VSCodeExtensionManager(root)
    root.mainloop()

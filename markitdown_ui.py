import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from pathlib import Path
import threading
import warnings
import re

# 不使用拖拽功能

# 忽略ffmpeg警告
warnings.filterwarnings("ignore", message="Couldn't find ffmpeg or avconv")

try:
    from markitdown import MarkItDown, UnsupportedFormatException, MissingDependencyException
except ImportError as e:
    print(f"错误：无法导入markitdown库")
    print(f"请运行：pip install markitdown[all]")
    print(f"详细错误：{e}")
    exit(1)

class MarkItDownUI:
    def __init__(self, root):
        self.root = root
        self.root.title("MarkItDown 文件转换器")
        self.root.geometry("800x600")
        
        # 初始化MarkItDown
        try:
            self.md = MarkItDown()
            self.setup_ui()
        except Exception as e:
            messagebox.showerror("初始化错误", f"无法初始化MarkItDown: {e}")
            self.root.destroy()
        
    def setup_ui(self):
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 文件选择区域
        file_frame = ttk.LabelFrame(main_frame, text="文件选择", padding="5")
        file_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        self.file_path = tk.StringVar()
        self.file_entry = ttk.Entry(file_frame, textvariable=self.file_path, width=60)
        self.file_entry.grid(row=0, column=0, padx=(0, 5))
        ttk.Button(file_frame, text="浏览", command=self.browse_file).grid(row=0, column=1)
        
        # URL输入区域
        url_frame = ttk.LabelFrame(main_frame, text="或输入URL", padding="5")
        url_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        self.url_path = tk.StringVar()
        ttk.Entry(url_frame, textvariable=self.url_path, width=70).grid(row=0, column=0)
        
        # 转换按钮
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=2, column=0, columnspan=2, pady=(0, 10))
        
        ttk.Button(button_frame, text="转换为Markdown", command=self.convert_file).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_frame, text="保存结果", command=self.save_result).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_frame, text="清空", command=self.clear_result).pack(side=tk.LEFT)
        
        # 进度条（初始状态隐藏）
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        self.progress.grid_remove()  # 初始隐藏
        
        # 结果显示区域
        result_frame = ttk.LabelFrame(main_frame, text="转换结果", padding="5")
        result_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        self.result_text = scrolledtext.ScrolledText(result_frame, wrap=tk.WORD, height=20)
        self.result_text.pack(fill=tk.BOTH, expand=True)
        
        # 状态栏
        self.status_var = tk.StringVar(value="就绪 - 请选择文件或输入URL")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN)
        status_bar.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))
        
        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(4, weight=1)
        
    def browse_file(self):
        filename = filedialog.askopenfilename(
            title="选择要转换的文件",
            filetypes=[
                ("所有支持的文件", "*.pdf;*.docx;*.pptx;*.xlsx;*.csv;*.html;*.epub;*.jpg;*.png"),
                ("PDF文件", "*.pdf"),
                ("Word文档", "*.docx"),
                ("PowerPoint", "*.pptx"),
                ("Excel文件", "*.xlsx;*.xls"),
                ("图像文件", "*.jpg;*.jpeg;*.png;*.gif;*.bmp"),
                ("所有文件", "*.*")
            ]
        )
        if filename:
            self.file_path.set(filename)
            self.url_path.set("")  # 清空URL
            
    def convert_file(self):
        file_path = self.file_path.get().strip()
        url_path = self.url_path.get().strip()
        
        if not file_path and not url_path:
            messagebox.showerror("错误", "请选择文件或输入URL")
            return
            
        # 在后台线程中执行转换
        threading.Thread(target=self._convert_worker, args=(file_path or url_path,), daemon=True).start()
        
    def _convert_worker(self, source):
        try:
            self.root.after(0, self._start_conversion)
            
            # 执行转换
            result = self.md.convert(source)
            
            # 更新UI
            self.root.after(0, self._conversion_complete, result, source)
            
        except UnsupportedFormatException:
            self.root.after(0, self._conversion_error, "不支持的文件格式")
        except MissingDependencyException as e:
            self.root.after(0, self._conversion_error, f"缺少依赖: {e}")
        except Exception as e:
            self.root.after(0, self._conversion_error, f"转换失败: {str(e)}")
            
    def _start_conversion(self):
        self.progress.grid()  # 显示进度条
        self.progress.start()
        self.status_var.set("正在转换...")
        self.result_text.delete(1.0, tk.END)
        
    def _conversion_complete(self, result, source):
        self.progress.stop()
        self.progress.grid_remove()  # 隐藏进度条
        self.status_var.set(f"转换完成: {Path(source).name if not source.startswith('http') else source}")
        
        # 显示结果
        self.result_text.delete(1.0, tk.END)
        self.result_text.insert(1.0, result.markdown)
        
        # 存储结果用于保存
        self.current_result = result.markdown
        
        # 根据源文件生成标题
        if source.startswith('http'):
            self.current_title = result.title or "web_content"
        else:
            # 使用原文件名（不含扩展名）作为标题
            source_path = Path(source)
            self.current_title = source_path.stem  # 文件名不含扩展名
        
    def _sanitize_filename(self, filename):
        """清理文件名中的非法字符"""
        # Windows文件名非法字符
        illegal_chars = r'[<>:"/\\|?*]'
        # 替换非法字符为下划线
        sanitized = re.sub(illegal_chars, '_', filename)
        # 移除多余的空格和点
        sanitized = sanitized.strip('. ')
        # 如果文件名为空，使用默认名称
        if not sanitized:
            sanitized = "converted_document"
        return sanitized

    def save_result(self):
        if not hasattr(self, 'current_result'):
            messagebox.showwarning("警告", "没有可保存的转换结果")
            return
        
        # 清理文件名
        clean_title = self._sanitize_filename(self.current_title)
        
        filename = filedialog.asksaveasfilename(
            title="保存Markdown文件",
            defaultextension=".md",
            filetypes=[("Markdown文件", "*.md"), ("文本文件", "*.txt"), ("所有文件", "*.*")],
            initialfile=f"{clean_title}.md"
        )
        
        if filename:
            try:
                with open(filename, 'w', encoding='utf-8') as f:
                    f.write(self.current_result)
                self.status_var.set(f"已保存: {Path(filename).name}")
                messagebox.showinfo("成功", f"文件已保存到: {filename}")
            except Exception as e:
                messagebox.showerror("保存错误", f"保存文件失败: {str(e)}")
                
    def clear_result(self):
        self.result_text.delete(1.0, tk.END)
        self.file_path.set("")
        self.url_path.set("")
        self.status_var.set("就绪")
        if hasattr(self, 'current_result'):
            delattr(self, 'current_result')


def main():
    root = tk.Tk()  # 使用普通的Tk
    app = MarkItDownUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()




import os
import zipfile
import ctypes
import webbrowser
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# 唤醒高清 DPI，消除边缘锯齿模糊
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass

class PPTExtractorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PPT 资源批量提取器")
        self.geometry("450x330") # 稍微调高了一点点以放下底部链接
        self.resizable(False, False)

        # 调用 Windows 原生主题
        style = ttk.Style(self)
        if 'vista' in style.theme_names():
            style.theme_use('vista')
        elif 'clam' in style.theme_names():
            style.theme_use('clam')

        self.file_paths = []
        self.category_map = {
            '图片': ['.png', '.jpg', '.jpeg', '.gif', '.emf', '.wmf', '.svg', '.bmp', '.tif', '.tiff'],
            '视频': ['.mp4', '.avi', '.mov', '.wmv', '.mkv', '.mpg', '.mpeg'],
            '音频': ['.mp3', '.wav', '.wma', '.m4a', '.aac', '.ogg']
        }

        self.setup_ui()

    def setup_ui(self):
        main_frame = ttk.Frame(self, padding="20 20 20 20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 标题与状态
        ttk.Label(main_frame, text="PPTX 资源批量提取器", font=("微软雅黑", 14, "bold")).pack(pady=(0, 15))
        
        self.path_label = ttk.Label(main_frame, text="等待选择文件...", foreground="#666666", font=("微软雅黑", 10))
        self.path_label.pack(pady=(0, 15))

        # 按钮
        ttk.Button(main_frame, text="1. 选择多个 PPTX 文件", command=self.select_files, width=30).pack(pady=5)

        # 复选框区域
        cb_frame = ttk.Frame(main_frame)
        cb_frame.pack(pady=15)
        
        self.var_img = tk.BooleanVar(value=True)
        self.var_vid = tk.BooleanVar(value=True)
        self.var_aud = tk.BooleanVar(value=True)

        ttk.Checkbutton(cb_frame, text="图片", variable=self.var_img).pack(side=tk.LEFT, padx=15)
        ttk.Checkbutton(cb_frame, text="视频", variable=self.var_vid).pack(side=tk.LEFT, padx=15)
        ttk.Checkbutton(cb_frame, text="音频", variable=self.var_aud).pack(side=tk.LEFT, padx=15)

        ttk.Button(main_frame, text="2. 开始批量提取", command=self.extract_resources, width=30).pack(pady=10)

        # --- 新增：作者与 GitHub 信息区域 ---
        author_frame = ttk.Frame(self)
        author_frame.pack(side=tk.BOTTOM, pady=10)
        
        # 制作一个看起来像超链接的 Label
        link_label = tk.Label(author_frame, text="© 作者: 散仙 | 访问 GitHub 主页", 
                              fg="#0066cc", font=("微软雅黑", 9, "underline"), cursor="hand2")
        link_label.pack()
        
        # 绑定鼠标左键点击事件，打开你的 GitHub 主页
        link_label.bind("<Button-1>", lambda e: webbrowser.open_new("https://github.com/sanxianla"))

    def select_files(self):
        self.file_paths = filedialog.askopenfilenames(
            title="请选择一个或多个 PPTX 文件",
            filetypes=[("PowerPoint 文件", "*.pptx")]
        )
        count = len(self.file_paths)
        if count > 0:
            self.path_label.config(text=f"已选中 {count} 个文件，准备就绪", foreground="#005A9E")

    def extract_resources(self):
        if not self.file_paths:
            messagebox.showwarning("提示", "请先选择至少一个 PPTX 文件！")
            return
        
        wanted = []
        if self.var_img.get(): wanted.append('图片')
        if self.var_vid.get(): wanted.append('视频')
        if self.var_aud.get(): wanted.append('音频')

        if not wanted:
            messagebox.showwarning("提示", "请至少勾选一种类型！")
            return

        total_files_processed = 0
        total_resources_extracted = 0

        for file_path in self.file_paths:
            ppt_dir = os.path.dirname(file_path)
            ppt_name = os.path.splitext(os.path.basename(file_path))[0]
            output_dir = os.path.join(ppt_dir, f"{ppt_name}_提取资源")

            try:
                extracted_for_this = 0
                with zipfile.ZipFile(file_path, 'r') as zip_ref:
                    for file_info in zip_ref.infolist():
                        if file_info.filename.startswith('ppt/media/') and not file_info.is_dir():
                            ext = os.path.splitext(file_info.filename)[1].lower()
                            target_cat = next((k for k, v in self.category_map.items() if ext in v), None)
                            
                            if target_cat and (target_cat in wanted):
                                cat_folder = os.path.join(output_dir, target_cat)
                                os.makedirs(cat_folder, exist_ok=True)
                                target_path = os.path.join(cat_folder, os.path.basename(file_info.filename))
                                
                                with open(target_path, 'wb') as f:
                                    f.write(zip_ref.read(file_info.filename))
                                extracted_for_this += 1
                                total_resources_extracted += 1
                                
                if extracted_for_this > 0:
                    total_files_processed += 1
            except Exception:
                continue

        if total_resources_extracted > 0:
            messagebox.showinfo("批量提取完成", f"任务结束！\n成功处理了 {total_files_processed} 个 PPT 文件。\n共计提取出 {total_resources_extracted} 个资源。")
        else:
            messagebox.showinfo("提示", "没有找到符合您勾选条件的资源。")

if __name__ == "__main__":
    app = PPTExtractorApp()
    app.mainloop()
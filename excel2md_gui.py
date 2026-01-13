"""
Excel转Markdown工具 - GUI版本
使用tkinter创建图形界面
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import pandas as pd
import sys


class Excel2MarkdownGUI:
    """Excel转Markdown GUI界面"""

    def __init__(self, root):
        self.root = root
        self.root.title("Excel转Markdown工具")
        self.root.geometry("700x550")
        self.root.resizable(True, True)

        # 设置窗口图标和样式
        self.setup_style()

        # 数据存储
        self.excel_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.sheets = []
        self.selected_sheets = []

        # 创建界面
        self.create_widgets()

    def setup_style(self):
        """设置界面样式"""
        style = ttk.Style()
        style.theme_use('clam')  # 使用现代主题

        # 设置按钮样式
        style.configure('TButton', padding=6)

    def create_widgets(self):
        """创建界面组件"""
        # 主框架
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)

        row = 0

        # 标题
        title_label = ttk.Label(
            main_frame,
            text="Excel转Markdown工具",
            font=('Arial', 16, 'bold')
        )
        title_label.grid(row=row, column=0, columnspan=3, pady=(0, 20))
        row += 1

        # 1. 选择Excel文件
        ttk.Label(main_frame, text="Excel文件:").grid(row=row, column=0, sticky=tk.W, pady=10)
        excel_entry = ttk.Entry(main_frame, textvariable=self.excel_path, width=50)
        excel_entry.grid(row=row, column=1, sticky=(tk.W, tk.E), pady=10, padx=5)
        excel_btn = ttk.Button(main_frame, text="浏览...", command=self.browse_excel)
        excel_btn.grid(row=row, column=2, pady=10)
        row += 1

        # 2. Sheet选择区域
        ttk.Label(main_frame, text="选择工作表:").grid(row=row, column=0, sticky=(tk.W, tk.N), pady=10)

        # Sheet列表框架
        sheet_frame = ttk.Frame(main_frame)
        sheet_frame.grid(row=row, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=10, padx=5)

        # Sheet选择按钮
        button_frame = ttk.Frame(sheet_frame)
        button_frame.grid(row=0, column=0, sticky=tk.W)

        ttk.Button(button_frame, text="全选", command=self.select_all_sheets, width=8).grid(row=0, column=0, padx=2)
        ttk.Button(button_frame, text="清空", command=self.deselect_all_sheets, width=8).grid(row=0, column=1, padx=2)

        # Sheet列表
        self.sheet_listbox = tk.Listbox(sheet_frame, height=6, selectmode=tk.MULTIPLE, exportselection=False)
        self.sheet_listbox.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(10, 0))

        # 滚动条
        scrollbar = ttk.Scrollbar(sheet_frame, orient=tk.VERTICAL, command=self.sheet_listbox.yview)
        scrollbar.grid(row=1, column=1, sticky=(tk.N, tk.S))
        self.sheet_listbox.config(yscrollcommand=scrollbar.set)

        sheet_frame.columnconfigure(0, weight=1)
        row += 1

        # 3. 输出路径
        ttk.Label(main_frame, text="输出文件:").grid(row=row, column=0, sticky=tk.W, pady=10)
        output_entry = ttk.Entry(main_frame, textvariable=self.output_path, width=50)
        output_entry.grid(row=row, column=1, sticky=(tk.W, tk.E), pady=10, padx=5)
        output_btn = ttk.Button(main_frame, text="浏览...", command=self.browse_output)
        output_btn.grid(row=row, column=2, pady=10)
        row += 1

        # 4. 预览区域
        ttk.Label(main_frame, text="预览:").grid(row=row, column=0, sticky=(tk.W, tk.N), pady=10)

        # 预览文本框
        preview_frame = ttk.Frame(main_frame)
        preview_frame.grid(row=row, column=1, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10, padx=5)

        self.preview_text = tk.Text(preview_frame, height=10, wrap=tk.WORD, font=('Consolas', 9))
        self.preview_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        preview_scrollbar = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL, command=self.preview_text.yview)
        preview_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.preview_text.config(yscrollcommand=preview_scrollbar.set)

        preview_frame.columnconfigure(0, weight=1)
        preview_frame.rowconfigure(0, weight=1)
        main_frame.rowconfigure(row, weight=1)
        row += 1

        # 5. 操作按钮
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=row, column=0, columnspan=3, pady=20)

        ttk.Button(button_frame, text="预览", command=self.preview_conversion).grid(row=0, column=0, padx=5)
        ttk.Button(button_frame, text="转换并保存", command=self.convert_and_save).grid(row=0, column=1, padx=5)
        ttk.Button(button_frame, text="退出", command=self.root.quit).grid(row=0, column=2, padx=5)
        row += 1

        # 状态栏
        self.status_label = ttk.Label(main_frame, text="就绪", relief=tk.SUNKEN, anchor=tk.W)
        self.status_label.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))

    def browse_excel(self):
        """浏览并选择Excel文件"""
        filename = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[
                ("Excel文件", "*.xlsx *.xls"),
                ("所有文件", "*.*")
            ]
        )

        if filename:
            self.excel_path.set(filename)
            # 自动设置输出路径
            output_file = Path(filename).with_suffix('.md')
            self.output_path.set(str(output_file))
            # 加载工作表列表
            self.load_sheets()
            self.update_status(f"已加载: {Path(filename).name}")

    def browse_output(self):
        """浏览并选择输出文件"""
        filename = filedialog.asksaveasfilename(
            title="保存Markdown文件",
            defaultextension=".md",
            filetypes=[
                ("Markdown文件", "*.md"),
                ("所有文件", "*.*")
            ]
        )

        if filename:
            self.output_path.set(filename)

    def load_sheets(self):
        """加载Excel文件中的工作表列表"""
        excel_file = self.excel_path.get()

        if not excel_file:
            return

        try:
            # 读取Excel文件获取所有工作表名称
            xl_file = pd.ExcelFile(excel_file)
            self.sheets = xl_file.sheet_names

            # 更新列表框
            self.sheet_listbox.delete(0, tk.END)
            for sheet in self.sheets:
                self.sheet_listbox.insert(tk.END, sheet)

            # 默认全选
            self.sheet_listbox.selection_set(0, tk.END)

            self.update_status(f"已加载 {len(self.sheets)} 个工作表")

        except Exception as e:
            messagebox.showerror("错误", f"加载工作表失败: {str(e)}")
            self.update_status("加载失败")

    def select_all_sheets(self):
        """全选工作表"""
        self.sheet_listbox.selection_set(0, tk.END)

    def deselect_all_sheets(self):
        """清空工作表选择"""
        self.sheet_listbox.selection_clear(0, tk.END)

    def get_selected_sheets(self):
        """获取选中的工作表"""
        selected_indices = self.sheet_listbox.curselection()
        return [self.sheets[i] for i in selected_indices]

    @staticmethod
    def dataframe_to_markdown(df, table_name=""):
        """将DataFrame转换为Markdown表格"""
        if df.empty:
            return f"## {table_name}\n\n表格为空\n"

        # 处理NaN值
        df = df.fillna("")

        # 生成Markdown
        lines = []

        if table_name:
            lines.append(f"## {table_name}")
            lines.append("")

        # 表头
        headers = df.columns.tolist()
        lines.append("| " + " | ".join(str(h).replace("\n", "<br>").replace("\r", "") for h in headers) + " |")

        # 分隔线
        lines.append("| " + " | ".join(["---"] * len(headers)) + " |")

        # 数据行
        for _, row in df.iterrows():
            # 处理每个单元格的换行符
            processed_row = []
            for val in row.tolist():
                # 将换行符替换为<br>，并去除回车符
                processed_val = str(val).replace("\r\n", "<br>").replace("\n", "<br>").replace("\r", "")
                processed_row.append(processed_val)
            row_data = " | ".join(processed_row)
            lines.append(f"| {row_data} |")

        return "\n".join(lines)

    def preview_conversion(self):
        """预览转换结果"""
        if not self.excel_path.get():
            messagebox.showwarning("警告", "请先选择Excel文件")
            return

        selected_sheets = self.get_selected_sheets()
        if not selected_sheets:
            messagebox.showwarning("警告", "请至少选择一个工作表")
            return

        try:
            self.update_status("正在预览...")
            markdown_sections = []

            for sheet_name in selected_sheets:
                df = pd.read_excel(self.excel_path.get(), sheet_name=sheet_name)
                section = self.dataframe_to_markdown(df, sheet_name)
                markdown_sections.append(section)
                markdown_sections.append("\n\n---\n\n")

            markdown = "\n".join(markdown_sections)

            # 显示预览（限制长度）
            self.preview_text.delete(1.0, tk.END)
            self.preview_text.insert(1.0, markdown)

            self.update_status(f"预览完成 - 共 {len(selected_sheets)} 个工作表")

        except Exception as e:
            messagebox.showerror("错误", f"预览失败: {str(e)}")
            self.update_status("预览失败")

    def convert_and_save(self):
        """转换并保存文件"""
        if not self.excel_path.get():
            messagebox.showwarning("警告", "请先选择Excel文件")
            return

        if not self.output_path.get():
            messagebox.showwarning("警告", "请指定输出文件路径")
            return

        selected_sheets = self.get_selected_sheets()
        if not selected_sheets:
            messagebox.showwarning("警告", "请至少选择一个工作表")
            return

        try:
            self.update_status("正在转换...")

            markdown_sections = []

            for sheet_name in selected_sheets:
                df = pd.read_excel(self.excel_path.get(), sheet_name=sheet_name)
                section = self.dataframe_to_markdown(df, sheet_name)
                markdown_sections.append(section)
                markdown_sections.append("\n\n---\n\n")

            markdown = "\n".join(markdown_sections)

            # 保存文件
            output_file = Path(self.output_path.get())
            output_file.parent.mkdir(parents=True, exist_ok=True)

            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(markdown)

            self.update_status(f"转换成功! 已保存到: {output_file}")
            messagebox.showinfo("成功", f"文件已成功转换并保存到:\n{output_file}")

        except Exception as e:
            messagebox.showerror("错误", f"转换失败: {str(e)}")
            self.update_status("转换失败")

    def update_status(self, message):
        """更新状态栏"""
        self.status_label.config(text=message)


def main():
    """主函数"""
    root = tk.Tk()
    app = Excel2MarkdownGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()

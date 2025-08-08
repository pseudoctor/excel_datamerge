#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel合并工具 - EXE优化版本
作者: Pseudoctor & AI Assistant
版本: 1.2
功能: 合并Excel，支持列名归一化，并能根据唯一标识（如条形码）统一字段（如商品名称）。
"""

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
os.environ['TK_SILENCE_DEPRECATION'] = "1"
import sys
import logging
import traceback
from datetime import datetime

# 设置程序信息
APP_NAME = "Excel合并工具"
APP_VERSION = "1.2"
APP_AUTHOR = "AI Assistant"

# 列名归一化规则
COLUMN_ALIASES = {
    '商品条形码': ['条形码', '条码', 'barcode', 'UPC','国条码'],
    '商品名称': ['名称', '品名', '产品名称', 'name', 'product name'],
    '品牌': ['品牌名称', 'brand', 'Brand Name'],
    '产品型号': ['型号', '产品规格', '规格', 'model', 'sku'],
    '含税销售金额': ['销售', '含税销售金额/元', '收入', 'sales', 'revenue','最终销售金额(销售金额+优惠券金额)'],
    '数量': ['销售数量', 'qty', 'quantity'],
    '日期': ['订单日期', 'date', '订单时间']
}


class ExcelMerger:
    def __init__(self):
        self.file_paths = []
        self.normalized_aliases = {k: [alias.lower() for alias in v] for k, v in COLUMN_ALIASES.items()}
        self.reverse_alias_map = {}
        for standard_name, aliases in self.normalized_aliases.items():
            for alias in aliases:
                self.reverse_alias_map[alias] = standard_name

        self.setup_logging()
        self.setup_gui()
        
    def setup_logging(self):
        try:
            if getattr(sys, 'frozen', False): app_dir = os.path.dirname(sys.executable)
            else: app_dir = os.path.dirname(os.path.abspath(__file__))
            log_file = os.path.join(app_dir, 'excel_merger.log')
            logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s', handlers=[logging.FileHandler(log_file, encoding='utf-8'), logging.StreamHandler(sys.stdout)])
            self.logger = logging.getLogger(__name__)
            self.logger.info(f"{APP_NAME} v{APP_VERSION} 启动")
        except Exception as e:
            logging.basicConfig(level=logging.INFO)
            self.logger = logging.getLogger(__name__)
            self.logger.warning(f"日志设置失败: {e}")
        
    def setup_gui(self):
        self.root = tk.Tk()
        self.root.title(f"{APP_NAME} v{APP_VERSION}")
        self.root.geometry("750x680")
        self.root.minsize(600, 550)
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        main_frame = tk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        title_frame = tk.Frame(main_frame)
        title_frame.pack(fill=tk.X, pady=(0, 10))
        tk.Label(title_frame, text=f"{APP_NAME} v{APP_VERSION}", font=("Helvetica", 16, "bold"), fg="#2196F3").pack()
        tk.Label(title_frame, text="专业级Excel合并与数据清洗工具", font=("Helvetica", 9), fg="gray").pack()
        
        files_frame = tk.LabelFrame(main_frame, text="待合并的Excel文件", font=("Helvetica", 10, "bold"))
        files_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        listbox_frame = tk.Frame(files_frame)
        listbox_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.listbox = tk.Listbox(listbox_frame, width=70, height=12, font=("Consolas", 9), selectmode=tk.EXTENDED)
        scrollbar_y = tk.Scrollbar(listbox_frame, orient=tk.VERTICAL, command=self.listbox.yview)
        scrollbar_x = tk.Scrollbar(listbox_frame, orient=tk.HORIZONTAL, command=self.listbox.xview)
        self.listbox.config(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        button_frame = tk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=5)
        tk.Button(button_frame, text="📁 添加文件", command=self.add_files, bg="#4CAF50", fg="white", font=("Helvetica", 9, "bold")).pack(side=tk.LEFT, padx=2)
        tk.Button(button_frame, text="❌ 移除选中", command=self.remove_selected, bg="#f44336", fg="white", font=("Helvetica", 9, "bold")).pack(side=tk.LEFT, padx=2)
        tk.Button(button_frame, text="🗑️ 清空列表", command=self.clear_list, bg="#ff9800", fg="white", font=("Helvetica", 9, "bold")).pack(side=tk.LEFT, padx=2)
        self.stats_label = tk.Label(button_frame, text="文件数量: 0", fg="gray")
        self.stats_label.pack(side=tk.RIGHT, padx=10)

        output_frame = tk.LabelFrame(main_frame, text="输出设置", font=("Helvetica", 10, "bold"))
        output_frame.pack(fill=tk.X, pady=5)
        path_frame = tk.Frame(output_frame)
        path_frame.pack(fill=tk.X, padx=5, pady=5)
        tk.Label(path_frame, text="保存位置:", font=("Helvetica", 9)).pack(anchor=tk.W)
        path_input_frame = tk.Frame(path_frame)
        path_input_frame.pack(fill=tk.X, pady=2)
        self.output_entry = tk.Entry(path_input_frame, font=("Consolas", 9))
        self.output_entry.pack(fill=tk.X, expand=True, side=tk.LEFT, padx=(0, 5))
        tk.Button(path_input_frame, text="浏览...", command=self.browse_output, bg="#2196F3", fg="white").pack(side=tk.RIGHT)
        
        options_frame = tk.LabelFrame(main_frame, text="合并与清洗选项", font=("Helvetica", 9, "bold"))
        options_frame.pack(fill=tk.X, pady=5)
        options_inner = tk.Frame(options_frame)
        options_inner.pack(fill=tk.X, padx=5, pady=(5,0))

        self.unify_names = tk.BooleanVar(value=True)
        tk.Checkbutton(options_inner, text="根据“商品条形码”统一“商品名称” (采用首次出现名称)", variable=self.unify_names, font=("Helvetica", 9)).pack(anchor=tk.W)
        self.normalize_columns = tk.BooleanVar(value=True)
        tk.Checkbutton(options_inner, text="智能统一相似列名 (如“品牌”和“品牌名称”合并)", variable=self.normalize_columns, font=("Helvetica", 9)).pack(anchor=tk.W)
        self.add_source_info = tk.BooleanVar(value=True)
        tk.Checkbutton(options_inner, text="添加来源信息列（文件名和工作表名）", variable=self.add_source_info, font=("Helvetica", 9)).pack(anchor=tk.W)
        self.skip_empty_sheets = tk.BooleanVar(value=True)
        tk.Checkbutton(options_inner, text="跳过空工作表", variable=self.skip_empty_sheets, font=("Helvetica", 9)).pack(anchor=tk.W)
        self.remove_duplicates = tk.BooleanVar(value=False)
        tk.Checkbutton(options_inner, text="移除完全重复的行", variable=self.remove_duplicates, font=("Helvetica", 9)).pack(anchor=tk.W)

        progress_frame = tk.Frame(main_frame)
        progress_frame.pack(fill=tk.X, pady=5)
        tk.Label(progress_frame, text="处理进度:", font=("Helvetica", 9)).pack(anchor=tk.W)
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=tk.X, pady=2)
        self.status_label = tk.Label(progress_frame, text="就绪", relief=tk.SUNKEN, anchor=tk.W, font=("Helvetica", 9), bg="#f0f0f0")
        self.status_label.pack(fill=tk.X, pady=2)

        merge_frame = tk.Frame(main_frame)
        merge_frame.pack(fill=tk.X, pady=10)
        self.merge_button = tk.Button(merge_frame, text="🚀 开始合并与处理", font=("Helvetica", 14, "bold"), command=self.merge_excel, bg="#2196F3", fg="white", height=2)
        self.merge_button.pack(fill=tk.X, ipady=5)

        default_path = os.path.join(os.path.expanduser("~"), "Desktop", f"合并结果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        self.output_entry.insert(0, default_path)
        self.listbox.bind('<Double-1>', self.show_full_path)
        self.update_stats()

    def show_full_path(self, event):
        selection = self.listbox.curselection()
        if selection:
            index = selection[0]
            if index < len(self.file_paths):
                file_path = self.file_paths[index]
                messagebox.showinfo("文件完整路径", file_path)

    def update_status(self, message):
        self.status_label.config(text=f"状态: {message}")
        self.root.update_idletasks()
        self.logger.info(message)
        
    def update_stats(self):
        self.stats_label.config(text=f"文件数量: {len(self.file_paths)}")
        
    def add_files(self):
        files = filedialog.askopenfilenames(title="选择Excel文件", filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")])
        if not files: return
        added_count = sum(1 for file in files if file not in self.file_paths)
        self.file_paths.extend(file for file in files if file not in self.file_paths)
        self.refresh_listbox()
        self.update_stats()
        self.update_status(f"成功添加 {added_count} 个文件")

    def remove_selected(self):
        selection = self.listbox.curselection()
        if not selection: return
        for index in reversed(selection): self.file_paths.pop(index)
        self.refresh_listbox()
        self.update_stats()
    
    def refresh_listbox(self):
        self.listbox.delete(0, tk.END)
        for i, file_path in enumerate(self.file_paths):
            self.listbox.insert(tk.END, f"{i+1:02d}. {os.path.basename(file_path)}")
    
    def clear_list(self):
        if self.file_paths and messagebox.askyesno("确认清空", f"确定要清空所有 {len(self.file_paths)} 个文件吗？"):
            self.file_paths.clear()
            self.refresh_listbox()
            self.update_stats()
            self.update_status("文件列表已清空")

    def browse_output(self):
        filename = filedialog.asksaveasfilename(
            initialfile=f"合并结果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")],
            title="选择保存位置"
        )
        if filename:
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, filename)

    def merge_excel(self):
        if not self.file_paths:
            messagebox.showwarning("警告", "请先添加要合并的Excel文件！")
            return
        output_path = self.output_entry.get().strip()
        if not output_path:
            messagebox.showwarning("警告", "请输入有效的输出文件路径！")
            return
        try:
            with open(output_path, 'a') as f: pass
        except PermissionError:
            messagebox.showerror("文件占用", f"输出文件被占用，请关闭文件后重试:\n{output_path}")
            return
        except IOError:
            pass # Path is likely valid

        self.merge_button.config(state='disabled', text="处理中...")
        self.root.config(cursor="watch")
        
        all_data_frames = []
        try:
            for i, file_path in enumerate(self.file_paths):
                self.update_status(f"正在读取: {os.path.basename(file_path)} ({i+1}/{len(self.file_paths)})")
                self.progress_var.set((i + 1) / len(self.file_paths) * 80)
                try:
                    sheets_dict = pd.read_excel(file_path, sheet_name=None, engine='openpyxl')
                    for sheet_name, df in sheets_dict.items():
                        if self.skip_empty_sheets.get() and df.empty: continue
                        df = self.clean_dataframe(df)
                        if self.add_source_info.get():
                            df.insert(0, '来源工作表名', sheet_name)
                            df.insert(0, '来源文件名', os.path.basename(file_path))
                        all_data_frames.append(df)
                except Exception as e:
                    self.logger.error(f"处理文件失败 {file_path}: {e}")
                    messagebox.showwarning("文件处理警告", f"处理文件时出错，已跳过:\n{os.path.basename(file_path)}\n错误: {str(e)}")
                    continue
            
            if not all_data_frames:
                messagebox.showinfo("提示", "所有选中的文件中没有找到可合并的数据。")
                return

            self.update_status("正在合并所有数据...")
            merged_df = pd.concat(all_data_frames, ignore_index=True, sort=False)
            
            if self.unify_names.get():
                self.update_status("正在根据条形码统一商品名称...")
                self.progress_var.set(85)
                barcode_col, name_col = '商品条形码', '商品名称'
                if barcode_col in merged_df.columns and name_col in merged_df.columns:
                    cleaned_df = merged_df.dropna(subset=[barcode_col])
                    if not cleaned_df.empty:
                        name_map = cleaned_df.drop_duplicates(subset=[barcode_col]).set_index(barcode_col)[name_col]
                        merged_df[name_col] = merged_df[barcode_col].map(name_map).fillna(merged_df[name_col])
                        self.logger.info("商品名称已根据条形码成功统一。")
                else:
                    self.logger.warning(f"无法统一商品名称，因为未找到'{barcode_col}'或'{name_col}'列。")

            if self.remove_duplicates.get():
                self.update_status("正在移除重复行...")
                merged_df.drop_duplicates(inplace=True)

            self.update_status("正在写入Excel文件...")
            self.progress_var.set(95)
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                merged_df.to_excel(writer, sheet_name="合并结果", index=False)

            self.progress_var.set(100)
            self.update_status("处理完成！")
            messagebox.showinfo("成功", "文件合并与数据清洗已成功完成！")

        except Exception as e:
            messagebox.showerror("合并错误", f"合并过程中发生错误: {e}")
            self.logger.error(f"合并错误: {e}\n{traceback.format_exc()}")
        finally:
            self.root.config(cursor="")
            self.progress_var.set(0)
            self.update_status("就绪")
            self.merge_button.config(state='normal', text="🚀 开始合并与处理")

    def _normalize_column_names(self, columns):
        new_columns = []
        for col in columns:
            if pd.isna(col): continue
            cleaned_col_lower = str(col).strip().lower()
            standard_name = self.reverse_alias_map.get(cleaned_col_lower, str(col).strip())
            new_columns.append(standard_name)
        return new_columns

    def clean_dataframe(self, df):
        try:
            if self.normalize_columns.get():
                df.columns = self._normalize_column_names(df.columns)
            
            df.columns = [str(col).strip().replace('\n', ' ').replace('\r', ' ') for col in df.columns]
            
            cols = df.columns.tolist()
            seen, new_cols = {}, []
            for col in cols:
                if col in seen:
                    seen[col] += 1
                    new_cols.append(f"{col}_{seen[col]}")
                else:
                    seen[col] = 0
                    new_cols.append(col)
            df.columns = new_cols
            return df
        except Exception as e:
            self.logger.warning(f"清理DataFrame时出错: {e}")
            return df

    def on_closing(self):
        if messagebox.askokcancel("退出", "确定要退出程序吗?"):
            self.root.destroy()

def main():
    try:
        app = ExcelMerger()
        app.root.mainloop()
    except Exception as e:
        logging.basicConfig(level=logging.ERROR, filename='app_startup_error.log')
        logging.error("程序启动失败", exc_info=True)
        messagebox.showerror("启动失败", f"程序启动时遇到严重错误: {e}\n详情请见 app_startup_error.log")

if __name__ == "__main__":
    main()
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import os
from datetime import datetime


class ExcelCompareApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel表格项目比对工具")
        self.root.geometry("900x700")
        
        self.file1_path = tk.StringVar()
        self.file2_path = tk.StringVar()
        self.reference_file = tk.StringVar()
        self.compare_folder = tk.StringVar()
        self.fixed_items = ["姓名", "身份证号"]
        self.custom_items = []
        self.selected_items = []
        
        self.setup_ui()
    
    def setup_ui(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(3, weight=1)
        
        ttk.Label(main_frame, text="Excel表格项目比对工具", font=("Arial", 16, "bold")).grid(row=0, column=0, columnspan=2, pady=(0, 20))
        
        ttk.Label(main_frame, text="文件选择模式:", font=("Arial", 10, "bold")).grid(row=1, column=0, sticky=tk.W, pady=(0, 5))
        
        mode_frame = ttk.Frame(main_frame)
        mode_frame.grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=(0, 15))
        
        self.mode_var = tk.StringVar(value="single")
        ttk.Radiobutton(mode_frame, text="单个比对（两个文件）", variable=self.mode_var, value="single", command=self.toggle_mode).pack(side=tk.LEFT, padx=(0, 20))
        ttk.Radiobutton(mode_frame, text="批量比对（参考文件+文件夹）", variable=self.mode_var, value="batch", command=self.toggle_mode).pack(side=tk.LEFT)
        
        self.single_frame = ttk.LabelFrame(main_frame, text="单个比对模式", padding="10")
        self.single_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 15))
        self.single_frame.columnconfigure(1, weight=1)
        
        ttk.Label(self.single_frame, text="正确参考文件:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(self.single_frame, textvariable=self.file1_path).grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)
        ttk.Button(self.single_frame, text="浏览...", command=self.select_file1).grid(row=0, column=2, pady=5)
        
        ttk.Label(self.single_frame, text="待比对文件:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(self.single_frame, textvariable=self.file2_path).grid(row=1, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)
        ttk.Button(self.single_frame, text="浏览...", command=self.select_file2).grid(row=1, column=2, pady=5)
        
        self.batch_frame = ttk.LabelFrame(main_frame, text="批量比对模式", padding="10")
        self.batch_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 15))
        self.batch_frame.columnconfigure(1, weight=1)
        self.batch_frame.grid_remove()
        
        ttk.Label(self.batch_frame, text="正确参考文件:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(self.batch_frame, textvariable=self.reference_file).grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)
        ttk.Button(self.batch_frame, text="浏览...", command=self.select_reference_file).grid(row=0, column=2, pady=5)
        
        ttk.Label(self.batch_frame, text="待比对文件夹:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(self.batch_frame, textvariable=self.compare_folder).grid(row=1, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)
        ttk.Button(self.batch_frame, text="浏览...", command=self.select_compare_folder).grid(row=1, column=2, pady=5)
        
        items_frame = ttk.LabelFrame(main_frame, text="比对项目选择", padding="10")
        items_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 15))
        items_frame.columnconfigure(0, weight=1)
        items_frame.rowconfigure(1, weight=1)
        
        self.items_vars = {}
        self.items_check_frame = ttk.Frame(items_frame)
        self.items_check_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        
        for item in self.fixed_items:
            var = tk.BooleanVar(value=True)
            self.items_vars[item] = var
            ttk.Checkbutton(self.items_check_frame, text=item, variable=var).pack(anchor=tk.W, pady=2)
        
        custom_frame = ttk.Frame(items_frame)
        custom_frame.grid(row=2, column=0, sticky=tk.W, pady=5)
        
        self.new_item_var = tk.StringVar()
        ttk.Entry(custom_frame, textvariable=self.new_item_var, width=30).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(custom_frame, text="添加项目", command=self.add_custom_item).pack(side=tk.LEFT)
        
        self.log_frame = ttk.LabelFrame(main_frame, text="进度日志", padding="10")
        self.log_frame.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 15))
        self.log_frame.columnconfigure(0, weight=1)
        self.log_frame.rowconfigure(0, weight=1)
        
        self.log_text = scrolledtext.ScrolledText(self.log_frame, height=10, width=80, state='disabled')
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        progress_frame = ttk.Frame(main_frame)
        progress_frame.grid(row=6, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        progress_frame.columnconfigure(0, weight=1)
        
        self.progress = ttk.Progressbar(progress_frame, mode='determinate')
        self.progress.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 10))
        self.progress_label = ttk.Label(progress_frame, text="0%")
        self.progress_label.grid(row=0, column=1)
        
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=7, column=0, columnspan=2, pady=10)
        
        ttk.Button(button_frame, text="开始比对", command=self.start_compare, width=15).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="清空日志", command=self.clear_log, width=15).pack(side=tk.LEFT, padx=5)
        
        copyright_label = tk.Label(self.root, text="© 奋 青", font=("Arial", 9), fg="#808080")
        copyright_label.place(relx=1.0, rely=0.99, anchor='se', x=-50)
    
    def toggle_mode(self):
        if self.mode_var.get() == "single":
            self.single_frame.grid()
            self.batch_frame.grid_remove()
        else:
            self.single_frame.grid_remove()
            self.batch_frame.grid()
    
    def select_file1(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel文件", "*.xlsx *.xls")])
        if file_path:
            self.file1_path.set(file_path)
            self.log(f"已选择第一个文件: {file_path}")
    
    def select_file2(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel文件", "*.xlsx *.xls")])
        if file_path:
            self.file2_path.set(file_path)
            self.log(f"已选择第二个文件: {file_path}")
    
    def select_reference_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel文件", "*.xlsx *.xls")])
        if file_path:
            self.reference_file.set(file_path)
            self.log(f"已选择参考文件: {file_path}")
    
    def select_compare_folder(self):
        folder_path = filedialog.askdirectory(title="选择待比对文件夹")
        if folder_path:
            self.compare_folder.set(folder_path)
            self.log(f"已选择待比对文件夹: {folder_path}")
    
    def get_excel_files_from_folder(self, folder_path):
        excel_files = []
        for file in os.listdir(folder_path):
            if file.endswith(('.xlsx', '.xls')) and not file.startswith('~$'):
                excel_files.append(os.path.join(folder_path, file))
        return sorted(excel_files)
    
    def add_custom_item(self):
        item = self.new_item_var.get().strip()
        if item and item not in self.items_vars:
            self.custom_items.append(item)
            var = tk.BooleanVar(value=True)
            self.items_vars[item] = var
            ttk.Checkbutton(self.items_check_frame, text=item, variable=var).pack(anchor=tk.W, pady=2)
            self.new_item_var.set("")
            self.log(f"已添加自定义项目: {item}")
        elif item in self.items_vars:
            messagebox.showwarning("警告", "该项目已存在！")
    
    def log(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.log_text.config(state='disabled')
        self.root.update()
    
    def clear_log(self):
        self.log_text.config(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state='disabled')
    
    def update_progress(self, value, max_value=100):
        percentage = int((value / max_value) * 100)
        self.progress['value'] = percentage
        self.progress_label.config(text=f"{percentage}%")
        self.root.update()
    
    def get_selected_items(self):
        selected = []
        for item, var in self.items_vars.items():
            if var.get():
                selected.append(item)
        return selected
    
    def read_excel(self, file_path):
        try:
            if file_path.endswith('.xls'):
                df = pd.read_excel(file_path, engine='xlrd')
            else:
                df = pd.read_excel(file_path, engine='openpyxl')
            return df
        except Exception as e:
            self.log(f"读取文件失败 {file_path}: {str(e)}")
            return None
    
    def normalize_value(self, value):
        if pd.isna(value) or value is None:
            return ""
        s = str(value).strip()
        if s.lower() in ['nan', 'none', 'null', '']:
            return ""
        return s
    
    def normalize_name(self, value):
        s = self.normalize_value(value)
        return s.replace(" ", "").replace("　", "")
    
    def normalize_id(self, value):
        s = self.normalize_value(value)
        if '.' in s:
            s = s.split('.')[0]
        return s
    
    def is_empty_row(self, row):
        for col, value in row.items():
            normalized = self.normalize_value(value)
            if normalized:
                return False
        return True
    
    def compare_two_files(self, reference_file, compare_file, selected_items, output_dir, save_individual_result=True):
        self.log(f"开始比对: {os.path.basename(reference_file)} 和 {os.path.basename(compare_file)}")
        
        df_ref = self.read_excel(reference_file)
        df_comp = self.read_excel(compare_file)
        
        if df_ref is None or df_comp is None:
            return False, []
        
        available_items = [item for item in selected_items if item in df_ref.columns and item in df_comp.columns]
        missing_items = [item for item in selected_items if item not in df_ref.columns or item not in df_comp.columns]
        
        if missing_items:
            self.log(f"警告: 以下项目在某些文件中不存在: {', '.join(missing_items)}")
        
        if not available_items:
            self.log("错误: 没有可比对的项目！")
            return False, []
        
        if "身份证号" not in available_items or "姓名" not in available_items:
            self.log("警告: 未选择身份证号或姓名，可能无法准确匹配！")
        
        self.log(f"将比对以下项目: {', '.join(available_items)}")
        
        result_df = df_comp.copy() if save_individual_result else None
        inconsistent_records = []
        
        if save_individual_result:
            for item in available_items:
                col_name = f"{item}_比对结果"
                result_df[col_name] = ""
        
        ref_key_map = {}
        for idx, row in df_ref.iterrows():
            id_val = self.normalize_id(row.get("身份证号", ""))
            name_val = self.normalize_name(row.get("姓名", ""))
            if id_val or name_val:
                key = (id_val, name_val)
                ref_key_map[key] = idx
        
        matched_count = 0
        unmatched_count = 0
        id_only_matched = 0
        name_only_matched = 0
        skipped_empty_rows = 0
        
        for comp_idx, comp_row in df_comp.iterrows():
            if self.is_empty_row(comp_row):
                skipped_empty_rows += 1
                continue
            
            comp_id = self.normalize_id(comp_row.get("身份证号", ""))
            comp_name = self.normalize_name(comp_row.get("姓名", ""))
            comp_key = (comp_id, comp_name)
            
            ref_idx = None
            match_type = None
            has_inconsistency = False
            
            if comp_key in ref_key_map:
                ref_idx = ref_key_map[comp_key]
                match_type = "full"
            else:
                found_by_id = None
                for (ref_id, ref_name), idx in ref_key_map.items():
                    if ref_id == comp_id and ref_id:
                        found_by_id = idx
                        id_only_matched += 1
                        match_type = "id_only"
                        break
                
                if found_by_id is not None:
                    ref_idx = found_by_id
                else:
                    found_by_name = None
                    for (ref_id, ref_name), idx in ref_key_map.items():
                        if ref_name == comp_name and ref_name:
                            found_by_name = idx
                            name_only_matched += 1
                            match_type = "name_only"
                            break
                    
                    if found_by_name is not None:
                        ref_idx = found_by_name
            
            item_results = {}
            
            if ref_idx is not None:
                ref_row = df_ref.iloc[ref_idx]
                
                if match_type == "full":
                    matched_count += 1
                else:
                    unmatched_count += 1
                
                for item in available_items:
                    if item == "姓名":
                        val_ref = self.normalize_name(ref_row.get(item, ""))
                        val_comp = self.normalize_name(comp_row.get(item, ""))
                    elif item == "身份证号":
                        val_ref = self.normalize_id(ref_row.get(item, ""))
                        val_comp = self.normalize_id(comp_row.get(item, ""))
                    else:
                        val_ref = self.normalize_value(ref_row.get(item, ""))
                        val_comp = self.normalize_value(comp_row.get(item, ""))
                    
                    is_inconsistent = val_ref != val_comp and not (val_ref == "" and val_comp == "")
                    
                    if is_inconsistent:
                        if save_individual_result:
                            col_name = f"{item}_比对结果"
                            result_df.at[comp_idx, col_name] = "不一致"
                        has_inconsistency = True
                        item_results[item] = "不一致"
                    else:
                        item_results[item] = ""
            else:
                unmatched_count += 1
                has_inconsistency = True
                
                if not comp_id and not comp_name:
                    self.log(f"提示: 待比对文件第 {comp_idx + 1} 行姓名和身份证号均为空，跳过比对")
                else:
                    self.log(f"提示: 待比对文件第 {comp_idx + 1} 行未找到匹配记录（身份证号: {comp_id}, 姓名: {comp_name}）")
                    
                    for item in available_items:
                        if save_individual_result:
                            col_name = f"{item}_比对结果"
                            result_df.at[comp_idx, col_name] = "不一致"
                        item_results[item] = "不一致"
            
            if has_inconsistency:
                record_data = {
                    "来源文件": os.path.basename(compare_file),
                    "行号": comp_idx + 1
                }
                
                for col in df_comp.columns:
                    record_data[col] = comp_row.get(col, "")
                
                for item in available_items:
                    if save_individual_result:
                        col_name = f"{item}_比对结果"
                        record_data[col_name] = result_df.at[comp_idx, col_name]
                    else:
                        record_data[f"{item}_比对结果"] = item_results.get(item, "不一致")
                
                inconsistent_records.append(record_data)
        
        self.log(f"完全匹配（身份证号+姓名）: {matched_count} 条记录")
        if id_only_matched > 0:
            self.log(f"仅身份证号匹配: {id_only_matched} 条记录")
        if name_only_matched > 0:
            self.log(f"仅姓名匹配: {name_only_matched} 条记录")
        if unmatched_count > 0:
            self.log(f"未找到匹配: {unmatched_count - id_only_matched - name_only_matched} 条记录")
        if skipped_empty_rows > 0:
            self.log(f"跳过空行: {skipped_empty_rows} 条")
        
        if save_individual_result:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_file = os.path.join(output_dir, f"{os.path.splitext(os.path.basename(compare_file))[0]}_比对结果_{timestamp}.xlsx")
            
            result_df.to_excel(output_file, index=False, engine='openpyxl')
            
            self.log(f"比对完成！结果已保存到: {output_file}")
        
        self.log(f"发现不一致记录: {len(inconsistent_records)} 条")
        
        return True, inconsistent_records
    
    def compare_batch_files(self, reference_file, compare_folder, selected_items, output_dir):
        if not os.path.exists(reference_file):
            messagebox.showerror("错误", "参考文件不存在！")
            return False
        
        if not os.path.exists(compare_folder):
            messagebox.showerror("错误", "待比对文件夹不存在！")
            return False
        
        excel_files = self.get_excel_files_from_folder(compare_folder)
        
        if not excel_files:
            messagebox.showwarning("警告", "文件夹中没有找到Excel文件！")
            return False
        
        self.log(f"开始批量比对，参考文件: {os.path.basename(reference_file)}")
        self.log(f"待比对文件数量: {len(excel_files)}")
        
        all_inconsistent_records = []
        
        for i, file_path in enumerate(excel_files):
            success, records = self.compare_two_files(reference_file, file_path, selected_items, output_dir, save_individual_result=False)
            if success:
                all_inconsistent_records.extend(records)
            self.update_progress(i + 1, len(excel_files))
        
        if all_inconsistent_records:
            self.generate_summary_report(all_inconsistent_records, output_dir)
        else:
            self.log("所有记录均一致，无需生成汇总表")
        
        self.log("批量比对完成！")
        return True
    
    def generate_summary_report(self, records, output_dir):
        self.log(f"正在生成汇总表，共 {len(records)} 条不一致记录")
        
        if not records:
            return
        
        df_summary = pd.DataFrame(records)
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        summary_file = os.path.join(output_dir, f"比对汇总结果_{timestamp}.xlsx")
        
        df_summary.to_excel(summary_file, index=False, engine='openpyxl')
        
        self.log(f"汇总表已保存到: {summary_file}")
    
    def start_compare(self):
        selected_items = self.get_selected_items()
        
        if not selected_items:
            messagebox.showwarning("警告", "请至少选择一个比对项目！")
            return
        
        output_dir = filedialog.askdirectory(title="选择结果保存目录")
        if not output_dir:
            return
        
        self.progress['value'] = 0
        self.progress_label.config(text="0%")
        
        if self.mode_var.get() == "single":
            ref_file = self.file1_path.get()
            comp_file = self.file2_path.get()
            
            if not ref_file or not comp_file:
                messagebox.showwarning("警告", "请选择两个文件！")
                return
            
            if not os.path.exists(ref_file) or not os.path.exists(comp_file):
                messagebox.showerror("错误", "文件不存在！")
                return
            
            self.update_progress(50)
            success, _ = self.compare_two_files(ref_file, comp_file, selected_items, output_dir)
            self.update_progress(100)
            
            if success:
                messagebox.showinfo("完成", "比对完成！请查看结果文件。")
        else:
            ref_file = self.reference_file.get()
            comp_folder = self.compare_folder.get()
            
            if not ref_file or not comp_folder:
                messagebox.showwarning("警告", "请选择参考文件和待比对文件夹！")
                return
            
            success = self.compare_batch_files(ref_file, comp_folder, selected_items, output_dir)
            self.update_progress(100)
            
            if success:
                messagebox.showinfo("完成", "批量比对完成！请查看结果文件。")


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelCompareApp(root)
    root.mainloop()

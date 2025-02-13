import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from collections import defaultdict

class DataDiffApp:
    def __init__(self, root):
        self.root = root
        self.root.title("当前-历史数据对比工具")
        self.root.geometry("1000x800")

        # 初始化变量
        self.file_path = ""
        self.df = None
        self.result_df = None
        self.column_map = defaultdict(list)  # 列名到索引的映射

        # 创建界面
        self.create_widgets()

    def create_widgets(self):
        # 文件选择区域
        file_frame = ttk.LabelFrame(self.root, text="1. 选择整合后的Excel文件")
        file_frame.pack(pady=10, padx=10, fill="x")

        self.btn_open = ttk.Button(
            file_frame, 
            text="打开Excel文件", 
            command=self.open_file
        )
        self.btn_open.pack(side="left", padx=5)

        # file_frame_2 = ttk.LabelFrame(self.root, text="2. 选择需要处理的列(包括价格)，无需选择历史数据")
        # file_frame_2.pack(pady=20, padx=10, fill="x")
        # 列选择区域（使用Canvas实现滚动）
        self.canvas = tk.Canvas(self.root)
        self.scrollbar = ttk.Scrollbar(
            self.root, 
            orient="vertical", 
            command=self.canvas.yview
        )
        self.col_frame = ttk.Frame(self.canvas)

        self.col_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(
                scrollregion=self.canvas.bbox("all")
            )
        )

        self.canvas.create_window((0,0), window=self.col_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        # 操作按钮
        btn_frame = ttk.Frame(self.root)
        btn_frame.pack(pady=5)

        self.btn_process = ttk.Button(
            btn_frame,
            text="计算结果",
            command=self.calculate_diff
        )
        self.btn_process.pack(side="left", padx=5)

        self.btn_export = ttk.Button(
            btn_frame,
            text="导出结果",
            command=self.export_result,
            state="disabled"
        )
        self.btn_export.pack(side="left", padx=5)

        # 结果显示区域
        result_frame = ttk.LabelFrame(self.root, text="处理结果")
        result_frame.pack(pady=10, padx=10, fill="both", expand=True)

        self.tree = ttk.Treeview(result_frame)
        self.tree.pack(fill="both", expand=True)

    def open_file(self):
        """加载Excel文件并构建列映射"""
        file_path = filedialog.askopenfilename(filetypes=[("Excel文件", "*.xlsx")])
        if file_path:
            try:
                self.file_path = file_path
                self.df = pd.read_excel(file_path)
                
                # 构建列名-索引映射
                self.column_map.clear()
                for idx, col in enumerate(self.df.columns):
                    self.column_map[col].append(idx)
                
                # 动态生成复选框
                self.create_column_checkboxes()
                messagebox.showinfo("成功", "文件加载完成！")
            except Exception as e:
                messagebox.showerror("错误", f"读取失败：{str(e)}")

    def create_column_checkboxes(self):
        """动态生成列选择复选框"""
        # 清空旧组件
        for widget in self.col_frame.winfo_children():
            widget.destroy()

        # 生成新的复选框
        self.selected_vars = {}
        for col in self.df.columns:
            var = tk.BooleanVar()
            chk = ttk.Checkbutton(
                self.col_frame,
                text=col,
                variable=var,
                command=lambda c=col: self.on_checkbox_click(c)
            )
            chk.pack(anchor="w", padx=5, pady=2)
            self.selected_vars[col] = var

    def on_checkbox_click(self, col_name):
        """处理复选框点击事件"""
        current_state = self.selected_vars[col_name].get()
        
        # 自动反选历史列
        if "_历史" in col_name:
            self.selected_vars[col_name].set(False)
            messagebox.showwarning("提示", "请选择当前数据列（非历史列）")
    
    def calculate_diff(self):
        """执行差值计算"""
        if self.df is None:
            messagebox.showwarning("警告", "请先选择文件！")
            return

        selected_cols = [
            col for col, var in self.selected_vars.items() 
            if var.get() and "_历史" not in col
        ]

        if not selected_cols:
            messagebox.showwarning("警告", "请至少选择一个当前数据列！")
            return

        # 准备结果数据
        self.result_df = self.df.copy()
        diff_columns = []

        
        price_col_cur =[] 
        price_col_his = []
        other_cols_cur = []
        other_cols_his = []

        print(f'selected_cols:{selected_cols}')

        for current_col in selected_cols:
            # 查找对应的历史列
            if "单价" in current_col: # 单独处理价格列
                price_col_cur = current_col # 赋值
                hist_cols = [
                    col for col in self.df.columns 
                    if "单价" in col and col not in selected_cols # 通过匹配
                ]
                if not hist_cols:
                    messagebox.showerror("错误", f"未找到 {price_col_cur} 对应的历史列")
                    return

                # 取最后一个出现的作为历史列（假设最后出现的为最新历史数据）
                price_col_his = hist_cols[-1] # 价格的历史列
                
                # 计算差值
                diff_col_name = f"价格差值"
                self.result_df[diff_col_name] = (
                    self.df[price_col_cur] - self.df[price_col_his]
                )
                diff_columns.append(diff_col_name)
            
            else:
                other_cols_cur.append(current_col) # 添加属性的当前列
                hist_cols = [
                    col for col in self.df.columns 
                    if current_col == col[:-2] and col not in selected_cols # 通过匹配
                ]
                if not hist_cols:
                    messagebox.showerror("错误", f"未找到 {current_col} 对应的历史列")
                    return
                hist_col = hist_cols[-1]
                other_cols_his.append(hist_col) # 添加属性的历史列
                # 计算差值
                diff_col_name = f"{current_col}_差值"
                self.result_df[diff_col_name] = (
                    self.df[current_col] - self.df[hist_col] #当前-历史 
                )
                diff_columns.append(diff_col_name)
                # 新增一列，根据差值判断正、零、负
                change_col_name = f"{current_col}_remark"
                self.result_df[change_col_name] = self.result_df[diff_col_name].apply(
                    lambda x: "增加" if x > 0 else ("不变" if x == 0 else("" if pd.isna(x) or x == '' else "减少"))
                )
        # todo: 计算如下属性列
        # - BCU等分别乘价格，得到中间数据
        # - 计算现在和历史的中间数据的差值
        # 遍历 other_cols 中的每一列

        # 计算当前总价
        print(f"price_col_cur:{price_col_cur}")
        print(f"price_col_his:{price_col_his}")
        print(f"other_cols_cur:{other_cols_cur}")
        print(f'other_cols_his:{other_cols_his}')
        for col in other_cols_cur:
            # 计算乘积
            col_name = f"{col}定点总价"
            print(f"col:{col}")
            # print(col ,price_col_cur ,self.df[col] , self.df[price_col_cur])
            self.result_df[col_name] = (
                self.df[col] * self.df[price_col_cur]
            )
            diff_columns.append(col_name)
        
        # 计算历史总价
        for col in other_cols_his:
            # 计算乘积
            col_name = f"{col[:-2]}平台化总价"
            print(f"col:{col}")
            # print(col ,price_col_cur ,self.df[col] , self.df[price_col_cur])
            self.result_df[col_name] = (
                self.df[col] * self.df[price_col_cur]
            )
            diff_columns.append(col_name)

        # 计算BOM Cost 差异，当前-历史
        for col in other_cols_cur: ## 从当前的列中挑，emmm，不知道对不对，还需要匹配吗，把后面几个字换掉就行
            # 计算乘积
            col_name = f"{col} BOM Cost 差异"
            print(f"col:{col}")
            # print(col ,price_col_cur ,self.df[col] , self.df[price_col_cur])
            self.result_df[col_name] = (
                self.result_df[col+"定点总价"] - self.result_df[col+"平台化总价"]
            )
            diff_columns.append(col_name)
        

        # 显示结果
        self.show_results(diff_columns)
        self.btn_export.config(state="normal")

    def show_results(self, diff_columns):
        """在Treeview中显示结果"""
        # 清空旧数据
        self.tree.delete(*self.tree.get_children())
        
        # 设置列
        self.tree["columns"] = list(self.result_df.columns)
        for col in self.result_df.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100, anchor="center")

        # 插入数据
        for _, row in self.result_df.iterrows():
            self.tree.insert("", "end", values=tuple(row))

    def export_result(self):
        """导出结果到Excel"""
        if self.result_df is not None:
            save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel文件", "*.xlsx")]
            )
            if save_path:
                try:
                    self.result_df.to_excel(save_path, index=False)
                    messagebox.showinfo("成功", f"文件已保存到：\n{save_path}")
                except Exception as e:
                    messagebox.showerror("错误", f"保存失败：{str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = DataDiffApp(root)
    root.mainloop()
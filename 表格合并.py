import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import numpy as np

def select_sheet(parent, prompt, sheets):
    """弹窗选择工作表"""
    result = None
    dialog = tk.Toplevel(parent)
    dialog.title(prompt)
    tk.Label(dialog, text=prompt).pack(padx=20, pady=5)
    var = tk.StringVar(value=sheets[0])
    tk.OptionMenu(dialog, var, *sheets).pack(padx=20, pady=5)
    def on_confirm():
        nonlocal result
        result = var.get()
        dialog.destroy()
    tk.Button(dialog, text="确定", command=on_confirm).pack(pady=5)
    dialog.transient(parent)
    dialog.grab_set()
    # print("here")
    parent.wait_window(dialog)
    
    # print("here")
    return result

def get_right_table(raw_df, head_str):
    # 查找"component"位置
    component_pos = raw_df.where(raw_df == head_str).stack().index.tolist()
    if not component_pos:
        raise ValueError(f"未找到包含{head_str}的单元格")
    
    # 取第一个找到的位置（行号，列号）
    header_row, header_col = component_pos[0]
    print(component_pos)
    print(header_row, header_col)
    
    # 获取列名（假设列名在component右侧同一行）
    columns = raw_df.iloc[header_row, header_col:].dropna().tolist()
    
    # 获取数据区域（从component下一行开始）
    data_start = header_row + 1
    data = raw_df.iloc[data_start:, header_col:header_col+len(columns)]
    
    # 构建DataFrame
    df = pd.DataFrame(data.values, columns=columns)
    
    # 清理全空行，但应该不用管，问题不大
    # return df.dropna(how='all')
    return df

def main():
    root = tk.Tk()
    # root.withdraw()
    # 选择Excel文件
    file_path = filedialog.askopenfilename(
        title="选择原始Excel文件",
        filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
    )
    if not file_path:
        messagebox.showinfo("取消", "已取消文件选择")
        return
    # 读取所有工作表名称
    try:
        xl = pd.ExcelFile(file_path)
        sheets = xl.sheet_names
    except Exception as e:
        messagebox.showerror("错误", f"读取文件失败: {e}")
        return
    # 引导用户选择四个工作表
    selections = []
    prompts = [
        "请选择旧价格表",
        "请选择旧BOM表",
        "请选择新价格表",
        "请选择新BOM表"
    ]
    for prompt in prompts:
        sheet = select_sheet(root, prompt, sheets)
        if not sheet:
            messagebox.showinfo("取消", "已取消操作")
            return
        selections.append(sheet)
    # 读取数据
    try:
        df_price_old = pd.read_excel(file_path, sheet_name=selections[0],header=None)
        df_quantity_old = pd.read_excel(file_path, sheet_name=selections[1],header=None)
        df_price_new = pd.read_excel(file_path, sheet_name=selections[2],header=None)
        df_quantity_new = pd.read_excel(file_path, sheet_name=selections[3],header=None)
    except Exception as e:
        messagebox.showerror("错误", f"读取数据失败: {e}")
        return
    # print(df_price_old)
    df_price_new = get_right_table(df_price_new, "Component")
    df_quantity_new = get_right_table(df_quantity_new, "Component")
    df_price_old = get_right_table(df_price_old, "JPN")
    df_quantity_old = get_right_table(df_quantity_old, "JPN")
    # 只修改部分列名
    df_price_old.rename(columns={'JPN': 'Component'}, inplace=True)
    df_quantity_old.rename(columns={'JPN': 'Component'}, inplace=True)

    # 使用 filter 方法保留指定列
    df_price_new = df_price_new.filter(['Component', '单价'])
    df_price_old = df_price_old.filter(['Component', 'Price'])

    # 修改价格列名，方便下一步处理
    df_price_old.rename(columns={'Price': '平台化单价'}, inplace=True)
    df_price_new.rename(columns={'单价': '定点单价'}, inplace=True)


    # 合并原始表
    # 需要提前提取出对应表格，这里直接抽取component即可，平台化的就抽取jpn。反正可调
    try:
        merged_old = pd.merge(
            df_price_old, df_quantity_old,
            on="Component", how="inner"
        )
    except KeyError:
        messagebox.showerror("错误", "原始表缺少'Component'列")
        return

    # 合并新表
    try:
        merged_new = pd.merge(
            df_price_new, df_quantity_new,
            on="Component", how="inner"
        )
    except KeyError:
        messagebox.showerror("错误", "新表缺少'Component'列")
        return
    print("here")
    # 最终合并
    try:
        final_merged = pd.merge(
            merged_new, merged_old,
            on="Component", how="outer",
            suffixes=("", ".1")
        )
    except Exception as e:
        messagebox.showerror("错误", f"最终合并失败: {e}")
        return
    print("here")
    # 保存结果
    save_path = filedialog.asksaveasfilename(
        title="保存结果文件",
        defaultextension=".xlsx",
        filetypes=[("Excel文件", "*.xlsx")]
    )
    if save_path:
        try:
            # 按照 'Component' 列的值进行排序（默认升序）
            final_merged = final_merged.sort_values(by='Component')
            final_merged = final_merged.replace('', np.nan).fillna(0)
            final_merged.to_excel(save_path, index=False)
            messagebox.showinfo("成功", f"文件已保存至: {save_path}")
        except Exception as e:
            messagebox.showerror("错误", f"保存失败: {e}")

if __name__ == "__main__":
    main()
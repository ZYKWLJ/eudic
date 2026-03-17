# import pandas as pd
# import math

# def split_excel_by_rows(input_file, output_prefix, rows_per_sheet=45, max_sheets_per_file=26):
#     """
#     将Excel文件按指定行数拆分到多个sheet中，每个sheet的序号列重新从1开始计数
#     每达到指定sheet数量后，保存为一个新文件，命名为：第i个月.xlsx
    
#     参数:
#     input_file: 输入的Excel文件路径
#     output_prefix: 输出文件前缀（无需带后缀）
#     rows_per_sheet: 每个sheet的最大行数，默认45行
#     max_sheets_per_file: 每个文件最多包含的sheet数量，默认26个
#     """
#     try:
#         # 读取原始Excel文件
#         df = pd.read_excel(input_file, header=0)
        
#         # 获取总行数
#         total_rows = len(df)
#         total_cols = len(df.columns)
        
#         print(f"原始数据总行数: {total_rows}, 总列数: {total_cols}")
        
#         # 计算总sheet数量
#         total_sheets = math.ceil(total_rows / rows_per_sheet)
#         print(f"将拆分为 {total_sheets} 个sheet，每个sheet最多 {rows_per_sheet} 行")
#         print(f"每 {max_sheets_per_file} 个sheet保存为一个文件，命名为：第i个月.xlsx\n")
        
#         # 获取序号列名（第一列）
#         serial_number_col = df.columns[0]
#         print(f"检测到序号列名为: {serial_number_col}")

#         # 全局变量：记录当前已经处理的sheet总数
#         current_sheet_count = 0
#         # 文件编号（第1个月、第2个月...）
#         file_number = 1
#         # 当前文件的Excel写入器
#         writer = None

#         # 遍历所有需要生成的sheet
#         for sheet_num in range(total_sheets):
#             # 每达到26个sheet，关闭当前文件，创建新文件
#             if current_sheet_count % max_sheets_per_file == 0:
#                 # 如果已有打开的writer，先保存关闭
#                 if writer is not None:
#                     writer.close()
#                     print(f"\n✅ 文件保存完成：第{file_number}个月.xlsx")
#                     file_number += 1
                
#                 # 创建新文件
#                 output_filename = f"第{file_number}个月.xlsx"
#                 writer = pd.ExcelWriter(output_filename, engine='openpyxl')

#             # 计算当前sheet的行范围
#             start_row = sheet_num * rows_per_sheet
#             end_row = min((sheet_num + 1) * rows_per_sheet, total_rows)
            
#             # 截取数据
#             sheet_data = df.iloc[start_row:end_row].copy()
            
#             # 重置序号从1开始
#             sheet_data[serial_number_col] = range(1, len(sheet_data) + 1)
            
#             # sheet命名
#             sheet_name = f"Unit_{current_sheet_count + 1:03d}"
            
#             # 写入sheet
#             sheet_data.to_excel(writer, sheet_name=sheet_name, index=False)
            
#             print(f"已创建 {sheet_name}，行 {start_row + 1} 至 {end_row} | 当前文件：第{file_number}个月")
            
#             # 累计sheet计数
#             current_sheet_count += 1

#         # 循环结束后，保存最后一个文件
#         if writer is not None:
#             writer.close()
#             print(f"\n✅ 最后一个文件保存完成：第{file_number}个月.xlsx")

#         print(f"\n🎉 全部拆分完成！共生成 {file_number} 个文件")
#         return True
        
#     except FileNotFoundError:
#         print(f"错误：找不到文件 '{input_file}'")
#         return False
#     except Exception as e:
#         print(f"错误：{str(e)}")
#         return False

# # 主程序执行
# if __name__ == "__main__":
#     # ========== 请在这里修改你的配置 ==========
#     INPUT_FILE = "4.xls"          # 输入文件路径
#     ROWS_PER_SHEET = 10           # 每个sheet的行数
#     MAX_SHEETS_PER_FILE = 26      # 每个文件的sheet数量（固定26）
    
#     # 执行拆分
#     split_excel_by_rows(
#         input_file=INPUT_FILE,
#         output_prefix="",
#         rows_per_sheet=ROWS_PER_SHEET,
#         max_sheets_per_file=MAX_SHEETS_PER_FILE
#     )

import pandas as pd
import math

def split_excel_by_rows(input_file, rows_per_sheet=45, max_sheets_per_file=26):
    """
    完整功能：
    1. 按指定行数拆分Excel
    2. 每个Sheet序号从1开始
    3. 每26个Sheet保存为一个文件：第1个月、第2个月...
    4. ✅ 在【第四列前面插入新列】，新列内容 = 复制第一列序号
    """
    try:
        # 读取文件
        df = pd.read_excel(input_file, header=0)
        total_rows = len(df)
        total_cols = len(df.columns)
        print(f"原始数据总行数: {total_rows}, 总列数: {total_cols}")

        total_sheets = math.ceil(total_rows / rows_per_sheet)
        print(f"总Sheet数: {total_sheets}，每Sheet {rows_per_sheet} 行")
        print(f"每 {max_sheets_per_file} 个Sheet保存为一个文件\n")

        # 第一列 = 序号列
        serial_col = df.columns[0]
        print(f"序号列: {serial_col}\n")

        current_sheet_count = 0
        file_number = 1
        writer = None

        for sheet_num in range(total_sheets):
            # 每满26个Sheet → 新建文件
            if current_sheet_count % max_sheets_per_file == 0:
                if writer is not None:
                    writer.close()
                    print(f"\n✅ 已保存：第{file_number}个月.xlsx")
                    file_number += 1

                output_filename = f"第{file_number}个月.xlsx"
                writer = pd.ExcelWriter(output_filename, engine='openpyxl')

            # 截取当前Sheet数据
            start_row = sheet_num * rows_per_sheet
            end_row = min((sheet_num + 1) * rows_per_sheet, total_rows)
            sheet_data = df.iloc[start_row:end_row].copy()

            # ====================== 核心操作 ======================
            # 1. 重置序号从1开始
            sheet_data[serial_col] = range(1, len(sheet_data) + 1)

            # 2. ✅ 在第四列前插入新列，内容 = 复制第一列
            # 插入位置：第3列索引（前面是0,1,2 → 插入后变成第4列）
            sheet_data.insert(loc=3, column=f'{serial_col}_复制', value=sheet_data[serial_col])
            # ======================================================

            # Sheet写入
            sheet_name = f"Unit_{current_sheet_count + 1:03d}"
            sheet_data.to_excel(writer, sheet_name=sheet_name, index=False)

            print(f"已创建 {sheet_name} | 行：{start_row + 1}~{end_row} | 文件：第{file_number}个月")
            current_sheet_count += 1

        # 保存最后一个文件
        if writer is not None:
            writer.close()
            print(f"\n✅ 最后一个文件保存完成：第{file_number}个月.xlsx")

        print(f"\n🎉 全部拆分完成！总共生成 {file_number} 个文件")
        return True

    except FileNotFoundError:
        print(f"错误：找不到文件 {input_file}")
        return False
    except Exception as e:
        print(f"错误：{str(e)}")
        return False

# ====================== 运行配置 ======================
if __name__ == "__main__":
    INPUT_FILE = "4.xls"         # 你的输入文件
    ROWS_PER_SHEET = 10          # 每个Sheet行数
    MAX_SHEETS_PER_FILE = 26     # 每个文件多少个Sheet（固定26）

    split_excel_by_rows(INPUT_FILE, ROWS_PER_SHEET, MAX_SHEETS_PER_FILE)
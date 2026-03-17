import pandas as pd
import math

def split_excel_by_rows(input_file, output_file, rows_per_sheet=45):
    """
    将Excel文件按指定行数拆分到多个sheet中，每个sheet的序号列重新从1开始计数
    
    参数:
    input_file: 输入的Excel文件路径
    output_file: 输出的Excel文件路径
    rows_per_sheet: 每个sheet的最大行数，默认15行
    """
    try:
        # 读取原始Excel文件（支持.xls和.xlsx格式）
        # header=0 表示第一行是列名
        df = pd.read_excel(input_file, header=0)
        
        # 获取总行数和列数
        total_rows = len(df)
        total_cols = len(df.columns)
        
        print(f"原始数据总行数: {total_rows}, 总列数: {total_cols}")
        
        # 计算需要创建的sheet数量
        total_sheets = math.ceil(total_rows / rows_per_sheet)
        print(f"将拆分为 {total_sheets} 个sheet，每个sheet最多 {rows_per_sheet} 行")
        
        # 获取序号列的列名（第一列）
        serial_number_col = df.columns[0]
        print(f"检测到序号列名为: {serial_number_col}")
        
        # 创建Excel写入器
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # 遍历每个sheet
            for sheet_num in range(total_sheets):
                # 计算当前sheet的起始行和结束行
                start_row = sheet_num * rows_per_sheet
                end_row = min((sheet_num + 1) * rows_per_sheet, total_rows)
                
                # 切片获取当前sheet的数据
                sheet_data = df.iloc[start_row:end_row].copy()  # 加copy避免SettingWithCopyWarning
                
                # 重新生成序号：从1开始，到当前sheet的行数结束
                sheet_data[serial_number_col] = range(1, len(sheet_data) + 1)
                
                # 定义sheet名称（格式：Unit_001, Unit_002...）
                sheet_name = f"Unit_{sheet_num + 1:03d}"
                
                # 将数据写入当前sheet
                sheet_data.to_excel(writer, sheet_name=sheet_name, index=False)
                
                print(f"已创建 {sheet_name}，包含行 {start_row + 1} 至 {end_row}，序号已重置为1-{len(sheet_data)}")
        
        print(f"\n拆分完成！新文件已保存至: {output_file}")
        return True
        
    except FileNotFoundError:
        print(f"错误：找不到文件 '{input_file}'")
        return False
    except Exception as e:
        print(f"错误：{str(e)}")
        return False

# 主程序执行
if __name__ == "__main__":
    # 请修改以下路径为你的实际文件路径
    INPUT_FILE = "4.xls"    # 输入文件路径
    OUTPUT_FILE = "4_10.xlsx"  # 输出文件路径
    ROWS_PER_SHEET = 10                  # 每个sheet的行数
    
    # 执行拆分
    split_excel_by_rows(INPUT_FILE, OUTPUT_FILE, ROWS_PER_SHEET)
import pandas as pd

# 1. 定义文件名前缀（只需要修改这里的 name 变量即可）
name_xls = "4"  # 你可以改成任意名称，比如 "vocab"、"english_words" 等
name = "4_10"  # 这是生成的文件

# 2. 拼接完整的文件路径
input_file = f'{name_xls}.xls'   # 输入的 Excel 文件
output_file = f'{name}.txt'  # 输出的 TXT 文件

# 3. 读取 Excel 文件
try:
    df = pd.read_excel(input_file)
except FileNotFoundError:
    print(f"错误：找不到文件 {input_file}，请检查文件名是否正确！")
    exit()

# 4. 提取单词列和释义列
# 这里假设列名就是 '单词' 和 '释义'，如果你的列名不同，请修改为实际名称
try:
    words = df['单词'].tolist()
    definitions = df['释义'].tolist()
except KeyError as e:
    print(f"错误：Excel 文件中找不到列 {e}，请检查列名是否正确！")
    exit()

# 5. 按每 45 个单词为一个 unit 分组
unit_size = 10
total_units = (len(words) + unit_size - 1) // unit_size  # 向上取整

# 6. 生成输出文本
output_content = []
for unit_num in range(total_units):
    start_idx = unit_num * unit_size
    end_idx = start_idx + unit_size
    unit_words = words[start_idx:end_idx]
    unit_defs = definitions[start_idx:end_idx]
    
    output_content.append(f'#unit {unit_num + 1}')
    for w, d in zip(unit_words, unit_defs):
        output_content.append(f'{w},{d}')
    #output_content.append('')  # 每个 unit 后空一行

# 7. 写入到 txt 文件
with open(output_file, 'w', encoding='utf-8') as f:
    f.write('\n'.join(output_content))

print(f"文件已生成：{output_file}")
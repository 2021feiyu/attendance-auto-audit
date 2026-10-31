#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
考勤表字段索引总表生成脚本
============================
功能：自动读取同文件夹下的Excel文件，生成字段索引映射表并保存为txt文件

工作流程：
    1. 自动查找脚本所在目录下的唯一Excel文件（.xlsx或.xls）
    2. 读取Excel的第一个工作表，以第4行作为表头行
    3. 遍历所有列，生成 F01、F02、F03... 格式的字段索引
    4. 自动识别每列的数据类型（文本/数字/日期）
    5. 将映射表保存为「考勤表字段索引总表.txt」

依赖库：
    - pandas：用于读取和处理Excel数据
    - openpyxl：p读取.xlsx文件的底层引擎

作者：Cline
版本：1.1
更新日期：2026-04-12
"""

import os       # 用于文件和路径操作
import glob     # 用于通配符匹配文件名
import pandas as pd  # 用于读取和分析Excel数据


# ============================================================
# 函数：find_excel_file
# 功能：自动查找脚本所在目录下的唯一Excel文件
# 参数：script_dir - 脚本所在的目录路径（字符串）
# 返回：Excel文件的完整路径（字符串）
# 异常：FileNotFoundError - 未找到Excel文件
#       ValueError - 找到多个Excel文件
# ============================================================
def find_excel_file(script_dir):
    """
    自动查找脚本所在目录下的唯一Excel文件
    
    查找逻辑：
        1. 使用glob模块匹配 *.xlsx 和 *.xls 文件
        2. 确保目录下只有一个Excel文件（避免歧义）
        3. 返回找到的Excel文件路径
    
    Args:
        script_dir (str): 脚本所在的目录路径
        
    Returns:
        str: Excel文件的完整路径
        
    Raises:
        FileNotFoundError: 当目录下没有Excel文件时抛出
        ValueError: 当目录下有多个Excel文件时抛出
    """
    # 定义Excel文件扩展名匹配模式
    # *.xlsx - Excel 2007+ 格式
    # *.xls  - Excel 97-2003 格式
    excel_patterns = ['*.xlsx', '*.xls']
    excel_files = []
    
    # 遍历所有匹配模式，收集所有Excel文件
    for pattern in excel_patterns:
        # glob.glob 返回匹配指定模式的所有文件路径
        excel_files.extend(glob.glob(os.path.join(script_dir, pattern)))
    
    # 根据找到的文件数量判断下一步操作
    if len(excel_files) == 0:
        # 没有找到任何Excel文件，抛出文件未找到异常
        raise FileNotFoundError("错误：未找到Excel文件（.xlsx或.xls）")
    elif len(excel_files) > 1:
        # 找到多个Excel文件，无法确定使用哪个，抛出值异常
        raise ValueError(f"错误：找到多个Excel文件，请确保文件夹内只有一个Excel文件\n找到的文件：{excel_files}")
    
    # 返回唯一找到的Excel文件路径
    return excel_files[0]


# ============================================================
# 函数：detect_column_type
# 功能：自动识别列的字段类型（文本/数字/日期）
# 参数：series - pandas Series对象，代表Excel中的一列数据
# 返回：字段类型字符串（"文本" / "数字" / "日期"）
# ============================================================
def detect_column_type(series):
    """
    自动识别列的字段类型
    
    识别逻辑：
        1. 遍历列中所有非空值
        2. 对每个值尝试判断其类型：
           - pd.Timestamp → 日期类型
           - int/float → 数字类型
           - 字符串 → 尝试解析为日期格式，失败则尝试转为数字，再失败则为文本
        3. 统计各类型占比，超过50%的类型即为该列的类型
    
    支持的日期格式：
        - %Y-%m-%d      (如：2024-01-15)
        - %Y/%m/%d      (如：2024/01/15)
        - %Y-%m-%d %H:%M:%S (如：2024-01-15 09:30:00)
        - %Y/%m/%d %H:%M:%S (如：2024/01/15 09:30:00)
        - %m/%d/%Y      (如：01/15/2024)
        - %d/%m/%Y      (如：15/01/2024)
    
    Args:
        series (pd.Series): pandas Series对象，代表一列数据
        
    Returns:
        str: 字段类型，可选值为 "文本"、"数字"、"日期"
    """
    # 去除空值（NaN/None），只分析有效数据
    non_null_values = series.dropna()
    
    # 如果整列都是空值，默认归类为文本类型
    if len(non_null_values) == 0:
        return "文本"
    
    # 初始化各类型计数器
    date_count = 0    # 日期类型计数
    number_count = 0  # 数字类型计数
    text_count = 0    # 文本类型计数
    
    # 遍历该列所有非空值，逐一判断类型
    for val in non_null_values:
        # 情况1：检查是否为pandas的日期时间类型（Timestamp）
        # pandas读取Excel时，日期格式的数据通常会被自动转为Timestamp对象
        if isinstance(val, pd.Timestamp):
            date_count += 1
            
        # 情况2：检查是否为原生数字类型（整数或浮点数）
        elif isinstance(val, (int, float)):
            number_count += 1
            
        # 情况3：其他类型（通常是字符串），需要进一步判断
        else:
            # 转为字符串并去除首尾空格
            str_val = str(val).strip()
            
            # 步骤A：尝试匹配常见日期格式
            is_date = False
            # 遍历支持的日期格式列表
            for fmt in ['%Y-%m-%d', '%Y/%m/%d', '%Y-%m-%d %H:%M:%S', '%Y/%m/%d %H:%M:%S', '%m/%d/%Y', '%d/%m/%Y']:
                try:
                    # 尝试按当前格式解析日期
                    pd.to_datetime(str_val, format=fmt)
                    is_date = True  # 解析成功，标记为日期
                    break           # 找到匹配格式，跳出循环
                except (ValueError, TypeError):
                    # 解析失败，继续尝试下一个格式
                    continue
            
            if is_date:
                # 成功解析为日期
                date_count += 1
            else:
                # 步骤B：尝试将字符串转换为数字
                try:
                    float(str_val)  # 能成功转换说明是数字
                    number_count += 1
                except (ValueError, TypeError):
                    # 转换失败，归类为文本
                    text_count += 1
    
    # 计算有效值的总数
    total = date_count + number_count + text_count
    
    # 根据占比判断该列的整体类型
    # 判断阈值：超过50%的值属于同一类型，则判定为该类型
    if date_count / total > 0.5:
        return "日期"
    elif number_count / total > 0.5:
        return "数字"
    else:
        # 默认归类为文本（包括混合类型或文本占多数的情况）
        return "文本"


# ============================================================
# 函数：generate_field_index_mapping
# 功能：读取Excel并生成字段索引映射表
# 参数：excel_path - Excel文件的完整路径（字符串）
# 返回：映射表的文本内容（字符串）
# ============================================================
def generate_field_index_mapping(excel_path):
    """
    生成字段索引映射表
    
    处理流程：
        1. 使用pandas读取Excel文件，指定第4行为表头
        2. 获取所有列名
        3. 遍历每列，生成 F01、F02... 格式的索引
        4. 调用 detect_column_type 识别每列的数据类型
        5. 组装映射表文本
    
    Args:
        excel_path (str): Excel文件的完整路径
        
    Returns:
        str: 完整的映射表文本内容，可直接写入txt文件
    """
    # 使用pandas读取Excel文件
    # header=3 表示将第4行（0-based索引为3）作为表头行
    # pandas会自动将表头行之前的行作为元数据跳过
    # engine='openpyxl' 指定使用openpyxl引擎读取.xlsx文件
    df = pd.read_excel(excel_path, header=3, engine='openpyxl')
    
    # 获取DataFrame的所有列名，转为Python列表
    # 这些列名就是Excel表头行（第4行）中的原始内容
    columns = df.columns.tolist()
    
    # 初始化映射表内容列表
    mapping_lines = []
    # 添加标题行
    mapping_lines.append("【考勤表字段索引总表】")
    # 添加空行，使输出更美观
    mapping_lines.append("")
    
    # 遍历所有列，enumerate从1开始计数（idx=1,2,3...）
    for idx, col_name in enumerate(columns, start=1):
        # 生成两位数的字段索引格式：F01, F02, F03, ..., F10, F11...
        # :02d 表示格式化为至少2位数字，不足补0
        field_index = f"F{idx:02d}"
        
        # 调用类型识别函数，获取该列的数据类型
        # df.iloc[:, idx - 1] 获取第idx列的所有数据（iloc使用0-based索引）
        col_type = detect_column_type(df.iloc[:, idx - 1])
        
        # 处理列名为空值（NaN）的情况
        # Excel中可能存在没有列名的空列
        if pd.isna(col_name):
            col_name_str = "未命名列"
        else:
            # 将列名转为字符串（可能是数字或其他类型）
            col_name_str = str(col_name)
        
        # 组装该行的映射信息，格式：Fxx: 列名 - 类型
        mapping_lines.append(f"{field_index}: {col_name_str} - {col_type}")
    
    # 将所有行用换行符连接，返回完整的映射表文本
    return "\n".join(mapping_lines)


# ============================================================
# 函数：main
# 功能：主函数，协调整个脚本的执行流程
# ============================================================
def main():
    """
    主函数 - 程序入口点
    
    执行流程：
        1. 获取脚本所在目录
        2. 查找Excel文件
        3. 生成字段索引映射
        4. 保存为txt文件
    """
    # 打印分隔线，美化输出
    print("=" * 50)
    print("考勤表字段索引总表生成工具")
    print("=" * 50)
    
    # 获取当前脚本文件所在的目录路径
    # __file__ 是Python内置变量，表示当前脚本的文件路径
    # os.path.abspath() 将路径转为绝对路径
    # os.path.dirname() 获取目录部分（去掉文件名）
    script_dir = os.path.dirname(os.path.abspath(__file__))
    print(f"\n脚本所在目录：{script_dir}")
    
    # ========== 步骤1：自动查找Excel文件 ==========
    print("\n正在查找Excel文件...")
    try:
        # 调用查找函数，获取Excel文件路径
        excel_path = find_excel_file(script_dir)
        # os.path.basename() 从完整路径中提取文件名
        print(f"找到Excel文件：{os.path.basename(excel_path)}")
    except (FileNotFoundError, ValueError) as e:
        # 捕获查找过程中的异常（未找到或多个文件）
        print(str(e))
        return  # 终止程序
    
    # ========== 步骤2：生成字段索引映射表 ==========
    print("\n正在生成字段索引映射表...")
    try:
        # 调用映射生成函数
        mapping_content = generate_field_index_mapping(excel_path)
    except Exception as e:
        # 捕获读取Excel或处理数据时的异常
        print(f"生成映射表失败：{str(e)}")
        return  # 终止程序
    
    # ========== 步骤3：保存映射表为txt文件 ==========
    # 构建输出文件的完整路径
    output_path = os.path.join(script_dir, "考勤表字段索引总表.txt")
    try:
        # 以写入模式（'w'）打开文件，使用UTF-8编码
        with open(output_path, 'w', encoding='utf-8') as f:
            # 将映射表内容写入文件
            f.write(mapping_content)
        
        # 打印成功信息和文件路径
        print(f"\n字段索引映射表已保存至：{output_path}")
        # 打印映射表内容预览
        print("\n映射表内容预览：")
        print("-" * 40)
        print(mapping_content)
        print("-" * 40)
    except Exception as e:
        # 捕获文件写入时的异常
        print(f"保存文件失败：{str(e)}")
        return  # 终止程序
    
    # 所有步骤完成
    print("\n处理完成！")


# ============================================================
# 程序入口
# ============================================================
# __name__ 是Python内置变量
# 当直接运行此脚本时，__name__ 的值为 "__main__"
# 当被其他脚本import时，__name__ 的值为模块名
# 这样可以确保只有直接运行时才执行main函数
if __name__ == "__main__":
    main()
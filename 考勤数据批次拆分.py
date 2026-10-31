#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
考勤数据批次拆分脚本 V2
===================
功能概述：
    1. 自动读取同目录下唯一Excel考勤数据文件
    2. 读取字段索引总表TXT文件，建立Fxx到列名的映射关系
    3. 为每行数据生成行索引（R0001格式）和列索引（Fxx_列名格式）
    4. 删除BQ列及之后的列
    5. 按20行/批次自动拆分数据
    6. 输出标准化TXT文件至「考勤批次数据」文件夹

依赖库：pandas, openpyxl
安装命令：pip install pandas openpyxl

运行方式：python 考勤数据批次拆分V2.py
"""

import pandas as pd
import os
import glob
import math
import re
import shutil


# ============================================================
# 配置区域 - 以下参数可根据需要修改
# ============================================================

# ========== 此处可手动修改索引表文件名 ==========
# 说明：该文件需与脚本放在同一目录下
# 格式要求：每行一个字段，格式为 F01:列名 或 F01 列名 或 F01=列名
# 示例：
#   F01: 序号 - 文本
#   F02: 姓名 - 文本
#   F03: 工作室 - 文本
INDEX_FILE_NAME = "考勤表字段索引总表.txt"
# ========== 此处可手动修改索引表文件名 ==========

# 每批次员工数据行数（固定20行，不可修改）
BATCH_SIZE = 20

# 输出文件夹名称（自动创建）
OUTPUT_FOLDER = "考勤批次数据"


# ============================================================
# 核心函数
# ============================================================

def find_excel_file():
    """
    自动查找当前目录下的唯一Excel文件
    
    返回:
        str: Excel文件的完整路径
        
    异常:
        FileNotFoundError: 未找到Excel文件时抛出
        ValueError: 找到多个Excel文件时抛出
    """
    script_dir = os.path.dirname(os.path.abspath(__file__))
    excel_patterns = ['*.xlsx', '*.xls']
    excel_files = []
    
    # 查找所有.xlsx和.xls文件
    for pattern in excel_patterns:
        excel_files.extend(glob.glob(os.path.join(script_dir, pattern)))
    
    if len(excel_files) == 0:
        raise FileNotFoundError(
            "错误：未找到任何Excel文件（.xlsx或.xls）\n"
            "请将考勤数据Excel文件放在与脚本同一目录下。"
        )
    elif len(excel_files) > 1:
        raise ValueError(
            f"错误：找到多个Excel文件 ({len(excel_files)}个)\n"
            f"请确保目录下只有一个考勤数据Excel文件。\n"
            f"找到的文件: {excel_files}"
        )
    
    return excel_files[0]


def read_index_file(index_file_path):
    """
    读取字段索引总表，返回Fxx到列名的映射字典
    
    参数:
        index_file_path: 索引表文件的完整路径
        
    返回:
        dict: {F01: '序号', F02: '姓名', F03: '工作室', ...}
        
    支持的索引表格式:
        - F01: 序号 - 文本    （推荐格式）
        - F01	序号 - 文本    （制表符分隔）
        - F01=序号 - 文本    （等号分隔）
        - F01|序号 - 文本    （竖线分隔）
    """
    index_mapping = {}
    
    with open(index_file_path, 'r', encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            
            # 方式1：正则表达式匹配 "F01: 列名 - 类型" 格式
            # 匹配规则：F开头+数字 + 分隔符(:=|) + 列名 + 可选的(- 类型)
            match = re.match(r'^(F\d+)\s*[:=\t|]\s*(.+?)(?:\s*-\s*.+)?$', line)
            if match:
                f_code = match.group(1)      # 如: F01
                col_name = match.group(2).strip()  # 如: 序号
                index_mapping[f_code] = col_name
                continue
            
            # 方式2：备用解析 - 简单分隔符分割
            parts = None
            for sep in ['\t', '=', '|']:
                if sep in line:
                    parts = line.split(sep, 1)
                    break
            
            if parts:
                f_code = parts[0].strip()
                col_name = parts[1].strip() if len(parts) > 1 else ''
                if f_code.startswith('F') and f_code[1:].isdigit():
                    index_mapping[f_code] = col_name
    
    return index_mapping


def read_excel_data(excel_path):
    """
    读取Excel文件数据
    
    参数:
        excel_path: Excel文件的完整路径
        
    返回:
        pandas.DataFrame: 包含考勤数据的DataFrame
        
    读取规则:
        - sheet_name=0: 读取第一个工作表
        - header=0: 跳过skiprows指定行后，取第0行作为表头
        - skiprows=[0,1,2]: 跳过前3行（通常是月份信息、统计行、星期行）
        - 最终表头为原始Excel的第4行
    """
    df = pd.read_excel(
        excel_path,
        sheet_name=0,           # 读取第一个工作表
        header=0,               # 跳过前3行后的第0行作为表头（即原始第4行）
        skiprows=[0, 1, 2]      # 跳过前3行（月份、统计、星期行）
    )
    return df


def delete_columns_from_bq(df):
    """
    删除DataFrame中BQ列（第69列，0索引为68）及之后的所有列
    
    参数:
        df: pandas DataFrame对象
        
    返回:
        pandas.DataFrame: 删除BQ列及之后列后的DataFrame
    
    Excel列号映射:
        A=0, ..., Z=25, AA=26, ..., BP=67, BQ=68
    """
    # BQ列在Excel中是第69列，对应0索引为68
    bq_col_index = 68
    
    if len(df.columns) > bq_col_index:
        cols_to_drop = df.columns[bq_col_index:]
        df = df.drop(columns=cols_to_drop)
    
    return df


def find_column_value(df, col_name, f_code, row_idx):
    """
    智能查找列对应的单元格值（支持多种列名类型匹配）
    
    参数:
        df: pandas DataFrame对象
        col_name: 索引表中的列名（字符串）
        f_code: Fxx编码（如F01）
        row_idx: 行索引
        
    返回:
        单元格值，如果未找到返回None
        
    匹配策略（按优先级）:
        1. 直接字符串匹配列名
        2. 整数转换匹配（日期列 1,2,3...31）
        3. 浮点数转换匹配（如 1.1, 1.2 等）
        4. Fxx编码直接匹配
        5. 字符串化列名匹配
    """
    # 策略1：直接字符串匹配列名
    if col_name in df.columns:
        return df.iloc[row_idx][col_name]
    
    # 策略2：整数转换匹配（处理日期列 1,2,3...31）
    try:
        int_col = int(col_name)
        if int_col in df.columns:
            return df.iloc[row_idx][int_col]
    except (ValueError, TypeError):
        pass
    
    # 策略3：浮点数转换匹配（处理 1.1, 1.2 等）
    try:
        float_col = float(col_name)
        if float_col in df.columns:
            return df.iloc[row_idx][float_col]
    except (ValueError, TypeError):
        pass
    
    # 策略4：Fxx编码直接匹配
    if f_code in df.columns:
        return df.iloc[row_idx][f_code]
    
    # 策略5：将Excel列名转为字符串后匹配
    for col in df.columns:
        if str(col) == col_name:
            return df.iloc[row_idx][col]
    
    return None


def generate_batch_content(df, index_mapping, batch_num, start_row, end_row):
    """
    生成单个批次的完整内容文本
    
    参数:
        df: pandas DataFrame对象
        index_mapping: Fxx到列名的映射字典
        batch_num: 批次编号（从1开始）
        start_row: 起始行索引（从0开始）
        end_row: 结束行索引（不包含）
        
    返回:
        str: 批次文件的完整内容
        
    输出格式示例:
        【待审核批次1】
        
        === 字段索引映射表 ===
        F01: 序号
        F02: 姓名
        ...
        =====================
        
        === 员工考勤数据 ===
        行R0001: F01_序号:HT306, F02_姓名:陈*, F03_工作室:深圳分公司, ...
        行R0002: F01_序号:FK012, F02_姓名:陈*, ...
        =====================
    """
    lines = []
    
    # 批次标题
    lines.append(f"【待审核批次{batch_num}】")
    lines.append("")
    
    # 字段索引映射表（按F01、F02...顺序排列）
    lines.append("=== 字段索引映射表 ===")
    sorted_f_codes = sorted(index_mapping.keys(), key=lambda x: int(x[1:]))
    for f_code in sorted_f_codes:
        col_name = index_mapping[f_code]
        lines.append(f"{f_code}: {col_name}")
    lines.append("=====================")
    lines.append("")
    
    # 员工考勤数据
    lines.append("=== 员工考勤数据 ===")
    
    for idx in range(start_row, end_row):
        if idx >= len(df):
            break
        
        # 生成行索引 R0001, R0002, ...（4位数字，与原始行顺序永久绑定）
        row_index = f"R{idx + 1:04d}"
        
        # 构建该行的字段数据
        field_values = []
        for f_code in sorted_f_codes:
            col_name = index_mapping[f_code]
            
            # 智能查找列值
            value = find_column_value(df, col_name, f_code, idx)
            
            # 空值/NaN原版保留，不做任何替换、修改、填充
            if pd.isna(value):
                field_values.append(f"{f_code}_{col_name}:")
            else:
                # 其他数据类型原样输出（日期、文本、数字等）
                field_values.append(f"{f_code}_{col_name}:{value}")
        
        # 单行格式: 行RXXXX: F01_列名:值, F02_列名:值, ...
        line = f"行{row_index}: " + ", ".join(field_values)
        lines.append(line)
    
    lines.append("=====================")
    
    return "\n".join(lines)


# ============================================================
# 主函数
# ============================================================

def main():
    """
    主函数：执行完整的考勤数据批次拆分流程
    
    执行步骤:
        1. 查找Excel文件
        2. 读取字段索引总表
        3. 读取Excel考勤数据
        4. 删除BQ列及之后的列
        5. 按20行/批次拆分
        6. 清理旧的输出文件夹
        7. 生成并保存批次文件
    """
    script_dir = os.path.dirname(os.path.abspath(__file__))
    
    print("=" * 60)
    print("考勤数据批次拆分脚本 V2")
    print("=" * 60)
    
    # 步骤1：查找Excel文件
    print("\n[1/7] 正在查找Excel文件...")
    excel_path = find_excel_file()
    print(f"  找到Excel文件: {os.path.basename(excel_path)}")
    
    # 步骤2：读取字段索引总表
    print("\n[2/7] 正在读取字段索引总表...")
    index_file_path = os.path.join(script_dir, INDEX_FILE_NAME)
    if not os.path.exists(index_file_path):
        raise FileNotFoundError(
            f"错误：未找到字段索引文件 '{INDEX_FILE_NAME}'\n"
            f"请确保该文件与脚本在同一目录下。\n"
            f"如需修改文件名，请编辑脚本中的 INDEX_FILE_NAME 变量。"
        )
    
    index_mapping = read_index_file(index_file_path)
    if not index_mapping:
        raise ValueError("错误：字段索引文件为空或格式不正确。")
    print(f"  成功读取 {len(index_mapping)} 个字段索引")
    
    # 步骤3：读取Excel考勤数据
    print("\n[3/7] 正在读取Excel考勤数据...")
    df = read_excel_data(excel_path)
    total_rows = len(df)
    print(f"  成功读取 {total_rows} 行员工数据")
    print(f"  原始列数: {len(df.columns)}")
    
    # 步骤4：删除BQ列及之后的列
    print("\n[4/7] 正在删除BQ列及之后的列...")
    original_col_count = len(df.columns)
    df = delete_columns_from_bq(df)
    new_col_count = len(df.columns)
    if original_col_count > new_col_count:
        print(f"  已删除 {original_col_count - new_col_count} 列 (BQ列及之后)，保留 {new_col_count} 列")
    else:
        print(f"  列数不超过BQ位置，无需删除 (当前共 {new_col_count} 列)")
    
    # 步骤5：批次拆分计算
    print("\n[5/7] 正在进行批次拆分...")
    num_batches = math.ceil(total_rows / BATCH_SIZE)
    print(f"  每批次 {BATCH_SIZE} 行，共需 {num_batches} 个批次")
    
    # 步骤6：清理旧的输出文件夹及数据
    print("\n[6/7] 正在清理旧的批次数据...")
    output_dir = os.path.join(script_dir, OUTPUT_FOLDER)
    if os.path.exists(output_dir):
        shutil.rmtree(output_dir)
        print(f"  已删除旧的输出文件夹: {OUTPUT_FOLDER}/")
    else:
        print(f"  未发现旧的输出文件夹，无需清理")
    
    # 创建新的输出文件夹
    os.makedirs(output_dir)
    print(f"  已创建新的输出文件夹: {OUTPUT_FOLDER}/")
    
    # 步骤7：生成并保存批次文件
    print("\n[7/7] 正在生成批次文件...")
    for batch_num in range(1, num_batches + 1):
        # 计算当前批次的起止行
        start_row = (batch_num - 1) * BATCH_SIZE
        end_row = min(start_row + BATCH_SIZE, total_rows)
        
        # 生成批次内容
        content = generate_batch_content(df, index_mapping, batch_num, start_row, end_row)
        
        # 保存文件
        output_filename = f"考勤批次_{batch_num}.txt"
        output_path = os.path.join(output_dir, output_filename)
        
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(content)
        
        print(f"  批次 {batch_num}: 行 {start_row + 1}-{end_row} -> {output_filename}")
    
    # 完成提示
    print("\n" + "=" * 60)
    print(f"处理完成！共生成 {num_batches} 个批次文件")
    print(f"输出位置: {output_dir}")
    print("=" * 60)


if __name__ == "__main__":
    main()
#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
故障码解析测试脚本
用于诊断故障分析问题

使用方法:
python test_fault_parse.py <Excel文件路径>
"""

import sys
import os

# 确保能找到 fault_code_parser 模块
script_dir = os.path.dirname(os.path.abspath(__file__))
if script_dir not in sys.path:
    sys.path.insert(0, script_dir)

print("=" * 60)
print("故障码解析诊断工具")
print("=" * 60)

# 1. 测试模块导入
print("\n[步骤1] 测试模块导入...")
try:
    from fault_code_parser import (
        parse_fault_value, 
        parse_bit_list,
        load_fault_codes_from_excel,
        get_fault_category,
        MAIN_FAULT_CODES,
        ALL_FAULT_CODES,
        USER_FAULT_CODES
    )
    print("  ✓ 模块导入成功")
except ImportError as e:
    print(f"  ✗ 模块导入失败: {e}")
    sys.exit(1)

# 2. 测试故障码定义加载
print("\n[步骤2] 测试故障码定义加载...")
load_fault_codes_from_excel()
print(f"  用户展示故障码寄存器数: {len(USER_FAULT_CODES)}")
print(f"  主故障码寄存器数: {len(MAIN_FAULT_CODES)}")
print(f"  总故障码寄存器数: {len(ALL_FAULT_CODES)}")

# 3. 测试解析函数
print("\n[步骤3] 测试解析函数...")
test_cases = [
    (20715, '[2,4]', '电池故障'),
    (20715, '[6]', '电池过压'),
    (20650, '[6]', '主故障码 bit6'),
    (20650, '[6,15]', '主故障码 bit6,15'),
]

for register, value, desc in test_cases:
    faults = parse_fault_value(register, value)
    if faults:
        fault_str = ', '.join([f'{f.code}:{f.name_cn}' for f in faults])
        print(f"  ✓ {desc}: {value} -> {fault_str}")
    else:
        print(f"  ✗ {desc}: {value} -> 无解析结果")

# 4. 测试Excel文件
if len(sys.argv) > 1:
    excel_file = sys.argv[1]
    print(f"\n[步骤4] 测试Excel文件: {excel_file}")
    
    if not os.path.exists(excel_file):
        print(f"  ✗ 文件不存在: {excel_file}")
        sys.exit(1)
    
    try:
        import pandas as pd
        df = pd.read_excel(excel_file)
        print(f"  ✓ 文件读取成功，共 {len(df)} 行")
        
        # 查找故障列
        fault_cols = []
        for col in df.columns:
            if '20650' in str(col) or '20651' in str(col) or '20712' in str(col) or '20715' in str(col):
                fault_cols.append(col)
        
        print(f"  找到故障列: {len(fault_cols)} 个")
        
        # 分析每个故障列
        for col in fault_cols[:3]:  # 只分析前3列
            print(f"\n  列: {col}")
            
            # 提取寄存器编号
            import re
            match = re.search(r'\((\d+)\)', col)
            if match:
                register = int(match.group(1))
                non_null = df[df[col].notna()]
                print(f"    寄存器: {register}")
                print(f"    非空值: {len(non_null)}")
                
                if len(non_null) > 0:
                    # 显示前3个非空值
                    sample_values = non_null[col].head(3).tolist()
                    print(f"    示例值: {sample_values}")
                    
                    # 解析
                    for val in sample_values[:1]:
                        faults = parse_fault_value(register, val)
                        if faults:
                            fault_str = ', '.join([f'{f.code}:{f.name_cn}' for f in faults])
                            print(f"    解析结果: {fault_str}")
                        else:
                            print(f"    解析结果: 无")
        
    except Exception as e:
        print(f"  ✗ 文件解析失败: {e}")
        import traceback
        traceback.print_exc()
else:
    print("\n[步骤4] 未提供Excel文件，跳过文件测试")
    print("  使用方法: python test_fault_parse.py <Excel文件路径>")

print("\n" + "=" * 60)
print("诊断完成")
print("=" * 60)

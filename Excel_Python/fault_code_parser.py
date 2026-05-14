#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
故障码解析模块
用于解析户储光伏平台产品的故障数据

功能：
1. 从Excel文件加载故障码定义（户储光伏平台产品客户故障码.xlsx）
2. 支持多种寄存器范围：
   - 主故障码：20650-20689（偏移量0-39）
   - 用户展示故障码：20712-20717
3. 解析bit位转换为故障码和故障名称

作者：OpenClaw
版本：v2.1.0
"""

import os
import re
import ast
from typing import Dict, List, Tuple, Optional
from dataclasses import dataclass


@dataclass
class FaultInfo:
    """故障信息数据类"""
    code: str           # 故障码，如 E402
    name_cn: str        # 中文名称
    name_en: str        # 英文名称
    fault_type: str     # 类型：故障/告警/提示
    register: int       # 寄存器地址
    bit: int            # bit位
    register_name: str  # 寄存器名称


# ===== 用户展示故障码定义 (寄存器20712-20717) =====
# 这是从旧版本保留的简化版故障码定义，用于用户界面展示

USER_FAULT_CODES = {
    # 电网故障 (寄存器20712)
    20712: {
        0: ('E100', '电网未接', 'No AC Connection', '告警'),
        1: ('E101', '电网电压异常', 'Grid voltage abnormal', '告警'),
        2: ('E102', '电网频率异常', 'Grid frequency abnormal', '告警'),
        3: ('E103', '电网相序接反', 'Grid phase sequence is error', '告警'),
        4: ('E104', '输出DCI超范围', 'Output DCI over-range', '告警'),
        5: ('E105', '防逆流超限故障', 'Anti-backflow over-limit fault', '告警'),
        6: ('E106', '电表或CT异常', 'Meter or CT abnormal', '告警'),
        7: ('E107', '电表或CT反接', 'Meter or CT reverse connection', '告警'),
        8: ('E108', '电网未满足并网范围', 'Grid Connection requirements not met', '告警'),
    },
    # 离网侧故障 (寄存器20713)
    20713: {
        0: ('E200', '离网输出电压过高', 'The off-grid output voltage is too high', '告警'),
        1: ('E201', '离网输出电压过低', 'The off-grid output voltage is too low', '告警'),
        2: ('E202', '交流过载', 'Ac overload', '告警'),
    },
    # PV侧故障 (寄存器20714)
    20714: {
        0: ('E300', 'PV输入过压', 'PV input overvoltage', '告警'),
        1: ('E301', 'PV组串反接', 'PV string reverse connection', '告警'),
        2: ('E302', 'PV过载', 'PV overload', '告警'),
        3: ('E303', '拉弧故障', 'AFCI fault', '告警'),
        4: ('E304', 'GFCI异常', 'Residual current High', '告警'),
        5: ('E305', '绝缘阻抗低', 'PV Isolation Low', '告警'),
        6: ('E306', 'PV短路', 'PV Short Circuit', '告警'),
    },
    # 电池故障 (寄存器20715)
    20715: {
        0: ('E400', '电池BMS故障', 'Battery BMS fault', '故障'),
        1: ('E401', '电池BMS运行异常', 'Battery BMS running abnormal', '告警'),
        2: ('E402', '电池BMS通信异常', 'Battery BMS communication abnormal', '故障'),
        3: ('E403', '电池DCDC故障', 'Battery DCDC fault', '故障'),
        4: ('E404', '电池DCDC运行异常', 'Battery DCDC run abnormal', '告警'),
        5: ('E405', '电池超额输出', 'Battery Overdischarge', '告警'),
        6: ('E406', '电池过压', 'Battery overvoltage', '告警'),
        7: ('E407', '电池欠压', 'Battery undervoltage', '告警'),
        8: ('E408', '电池电量低', 'Battery SOC is low', '提示'),
        9: ('E409', '电池过温故障', 'Battery Over-temperature protection', '告警'),
        10: ('E410', '电池低温故障', 'Battery Low-temperature protection', '告警'),
        11: ('E411', '电池端口过压故障', 'Battery Port overvoltage', '告警'),
        12: ('E412', '电池加热异常', 'Battery PTC Abnormal', '告警'),
    },
    # 油机故障 (寄存器20716)
    20716: {
        0: ('E500', '油机输出过载', 'Oil machine overload', '告警'),
        1: ('E501', '油机故障', 'Oil machine fault', '故障'),
    },
    # 系统故障 (寄存器20717)
    20717: {
        0: ('E600', '内部通信故障', 'Internal communication fault', '故障'),
        1: ('E601', 'PCS运行异常', 'PCS is running abnormally', '告警'),
        2: ('E602', 'PCS故障', 'PCS fault', '故障'),
        3: ('E603', 'OTA升级失败', 'OTA upgrade failure', '故障'),
        4: ('E604', '过温保护', 'Over-temperature protection', '告警'),
        5: ('E605', '风扇异常', 'Fan abnormal', '故障'),
        6: ('E606', '并机异常', 'Parallel machine abnormal', '故障'),
        7: ('E607', '系统版本不一致', 'System Version Incompatible', '告警'),
        8: ('E608', '软件版本错误', 'Software version wrong', '故障'),
        9: ('E609', '软件参数有误', 'Software parameter incorrect', '告警'),
    },
}

# 寄存器名称映射
REGISTER_NAMES = {
    20712: '电网故障（用户展示）',
    20713: '离网侧故障（用户展示）',
    20714: 'PV侧故障（用户展示）',
    20715: '电池故障（用户展示）',
    20716: '油机故障（用户展示）',
    20717: '系统故障（用户展示）',
}

# 故障分类名称
FAULT_CATEGORIES = {
    'E1': '电网故障',
    'E2': '离网侧故障',
    'E3': 'PV侧故障',
    'E4': '电池故障',
    'E5': '油机故障',
    'E6': '系统故障',
}

# 全局故障码字典（从Excel加载后填充）
MAIN_FAULT_CODES = {}  # 主故障码：20650-20689
ALL_FAULT_CODES = {}   # 所有故障码（合并用户展示+主故障码）


def load_fault_codes_from_excel(excel_path: str = None) -> bool:
    """
    从Excel文件加载故障码定义
    
    注意：此函数用于解析"户储光伏平台产品客户故障码.xlsx"格式。
    如需解析"平台逆变器错误代码对照表.xlsx"（点表格式），请使用 load_fault_codes_from_point_table()
    
    Args:
        excel_path: Excel文件路径
        
    Returns:
        是否加载成功
    """
    global MAIN_FAULT_CODES, ALL_FAULT_CODES
    
    if excel_path is None:
        # 尝试多个可能的路径（优先使用户储光伏平台产品客户故障码.xlsx）
        possible_paths = [
            '户储光伏平台产品客户故障码.xlsx',
            'assets/户储光伏平台产品客户故障码.xlsx',
            'assets/excel_workspace/户储光伏平台产品客户故障码.xlsx',
            os.path.join(os.path.dirname(__file__), '..', '户储光伏平台产品客户故障码.xlsx'),
            os.path.join(os.path.dirname(__file__), '户储光伏平台产品客户故障码.xlsx'),
            os.path.join(os.path.dirname(__file__), '..', 'assets', '户储光伏平台产品客户故障码.xlsx'),
            '/workspace/projects/assets/户储光伏平台产品客户故障码.xlsx',  # 添加绝对路径
            '/workspace/projects/assets/excel_workspace/户储光伏平台产品客户故障码.xlsx',  # 添加绝对路径
        ]
        for p in possible_paths:
            if os.path.exists(p):
                excel_path = p
                print(f"[INFO] 找到故障码定义文件: {excel_path}")
                break
    
    if excel_path is None or not os.path.exists(excel_path):
        print(f"[WARN] 未找到故障码定义Excel文件，使用内置定义")
        # 即使没有Excel，也要初始化ALL_FAULT_CODES
        ALL_FAULT_CODES = {}
        ALL_FAULT_CODES.update(USER_FAULT_CODES)
        return False
    
    try:
        import pandas as pd
        
        # 尝试不同的sheet名称
        xl = pd.ExcelFile(excel_path)
        sheet_name = None
        for name in xl.sheet_names:
            if '客户故障码' in name or '故障码' in name:
                sheet_name = name
                break
        
        if sheet_name is None:
            sheet_name = xl.sheet_names[0] if xl.sheet_names else 0
        
        df = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)
        
        MAIN_FAULT_CODES.clear()
        
        # 解析新的Excel格式
        # 列结构：序号(0)、分类(1)、主故障码(2)、主故障码中文描述(3)、主故障码英文描述(4)、类型(5)、显示对象(6)、
        #         逆变器动作(7)、处理方式(8)、触发条件(9)、恢复条件(10)、新项目映射关系(11)
        
        for i in range(len(df)):
            row = df.iloc[i]
            
            # 获取故障码
            fault_code = str(row[2]).strip() if pd.notna(row[2]) else ''
            if not fault_code or not fault_code.startswith('E'):
                continue
            
            # 获取故障描述
            name_cn = str(row[3]).strip() if pd.notna(row[3]) else ''
            name_en = str(row[4]).strip() if pd.notna(row[4]) else ''
            fault_type = str(row[5]).strip() if pd.notna(row[5]) else '告警'
            
            # 获取映射关系（列11）
            mapping = str(row[11]).strip() if pd.notna(row[11]) else ''
            
            # 解析映射关系，支持多行格式：
            # 1、20650：
            # bit5（电网安规过频保护）= 1
            # bit6（电网安规欠频保护）= 1
            # 也支持带括号注释的格式：
            # 1、20666（20666为告警，非故障）：
            # bit15（电池低电量）= 1
            if mapping and mapping != 'nan':
                # 第一步：找到所有寄存器，支持带括号注释的格式
                # 匹配: 20650： 或 20666（注释）： 或 20666: 
                register_matches = list(re.finditer(r'(\d{5})(?:[（(][^）)]*[）)])?\s*[：:]', mapping))
                
                for idx, match in enumerate(register_matches):
                    register = int(match.group(1))
                    
                    # 确定这个寄存器的映射文本范围
                    start_pos = match.end()
                    if idx + 1 < len(register_matches):
                        end_pos = register_matches[idx + 1].start()
                    else:
                        end_pos = len(mapping)
                    
                    # 提取这个寄存器的映射文本
                    reg_text = mapping[start_pos:end_pos]
                    
                    # 从这个文本中提取所有 bit
                    # 支持：bit5（描述）= 1 或 bit5 = 1 或 bit5
                    bit_matches = re.findall(r'bit(\d+)', reg_text)
                    
                    for bit_str in bit_matches:
                        bit = int(bit_str)
                        
                        # 确保寄存器在主故障码范围内
                        if 20650 <= register <= 20699:
                            if register not in MAIN_FAULT_CODES:
                                MAIN_FAULT_CODES[register] = {}
                            
                            MAIN_FAULT_CODES[register][bit] = (
                                fault_code,
                                name_cn or fault_code,
                                name_en,
                                fault_type
                            )
        
        # 为没有映射关系的寄存器设置默认名称
        for addr in range(20650, 20690):
            if addr not in MAIN_FAULT_CODES:
                MAIN_FAULT_CODES[addr] = {}
        
        # 合并所有故障码
        ALL_FAULT_CODES = {}
        ALL_FAULT_CODES.update(USER_FAULT_CODES)
        ALL_FAULT_CODES.update(MAIN_FAULT_CODES)
        
        # 统计加载的故障码数量
        total_bits = sum(len(bits) for bits in MAIN_FAULT_CODES.values())
        print(f"[INFO] 从故障码定义表加载了 {len(MAIN_FAULT_CODES)} 个寄存器，{total_bits} 个故障码定义")
        return True
        
    except Exception as e:
        print(f"[ERROR] 加载故障码定义失败: {e}")
        import traceback
        traceback.print_exc()
        # 即使加载失败，也要初始化ALL_FAULT_CODES
        ALL_FAULT_CODES = {}
        ALL_FAULT_CODES.update(USER_FAULT_CODES)
        return False


def load_fault_codes_from_point_table(excel_path: str = None) -> bool:
    """
    从点表格式Excel文件加载故障码定义
    
    点表格式（平台逆变器错误代码对照表.xlsx）：
    - 列1: 寄存器偏移量（只有每个寄存器第一行有值）
    - 列2: 寄存器名称
    - 列7: 故障ID（bit位）如 "bit0"
    - 列8: 十进制值
    - 列9: 故障中文名称
    - 列10: 故障英文名称
    
    寄存器地址 = 20650 + 偏移量
    
    Args:
        excel_path: Excel文件路径
        
    Returns:
        是否加载成功
    """
    global MAIN_FAULT_CODES, ALL_FAULT_CODES
    
    # 获取当前脚本所在目录，用于构建相对路径
    script_dir = os.path.dirname(os.path.abspath(__file__))
    project_root = os.path.dirname(script_dir)  # 项目根目录
    
    if excel_path is None:
        # 尝试查找点表文件（按优先级排序）
        possible_paths = [
            # 优先使用项目根目录下的assets目录
            os.path.join(project_root, 'assets', '平台逆变器错误代码对照表.xlsx'),
            os.path.join(project_root, '平台逆变器错误代码对照表.xlsx'),
            # 当前脚本目录
            os.path.join(script_dir, '平台逆变器错误代码对照表.xlsx'),
            # 当前工作目录
            '平台逆变器错误代码对照表.xlsx',
            'assets/平台逆变器错误代码对照表.xlsx',
            # 其他可能的位置
            os.path.join(project_root, 'assets', 'MARS项目故障底层逻辑声明.xlsx'),
            os.path.join(project_root, 'assets', '故障点表.xlsx'),
            # 绝对路径（作为后备）
            '/workspace/projects/assets/平台逆变器错误代码对照表.xlsx',
            '/workspace/projects/assets/excel_workspace/../平台逆变器错误代码对照表.xlsx',
        ]
        for p in possible_paths:
            if os.path.exists(p):
                excel_path = p
                print(f"[INFO] 找到点表文件: {excel_path}")
                break
    
    if excel_path is None or not os.path.exists(excel_path):
        print(f"[WARN] 未找到点表文件")
        print(f"[WARN] 项目根目录: {project_root}")
        print(f"[WARN] 脚本目录: {script_dir}")
        return False
    
    try:
        import pandas as pd
        
        xl = pd.ExcelFile(excel_path)
        
        # 优先使用"软件平台故障码定义"sheet
        sheet_name = None
        for name in xl.sheet_names:
            if '软件平台故障码定义' in name or '故障码定义' in name:
                sheet_name = name
                break
        
        if sheet_name is None:
            for name in xl.sheet_names:
                if '故障' in name:
                    sheet_name = name
                    break
        
        if sheet_name is None:
            sheet_name = xl.sheet_names[0] if xl.sheet_names else 0
        
        df = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)
        
        # 查找标题行（包含"寄存器偏移量"的行）
        header_row = None
        for i in range(min(20, len(df))):
            row = df.iloc[i]
            for cell in row:
                if pd.notna(cell) and '寄存器偏移量' in str(cell):
                    header_row = i
                    break
            if header_row is not None:
                break
        
        # 确定数据起始行
        if header_row is not None:
            start_row = header_row + 1
        else:
            start_row = 0
        
        print(f"[DEBUG] 点表文件: {excel_path}, Sheet: {sheet_name}")
        print(f"[DEBUG] 数据行数: {len(df)}, 标题行: {header_row}, 数据起始行: {start_row}")
        
        # 固定列索引（基于实际文件格式）
        OFFSET_COL = 1    # 寄存器偏移量
        REG_NAME_COL = 2  # 寄存器名称
        BIT_COL = 7       # 故障ID（bit位）
        NAME_CN_COL = 9   # 故障中文名称
        NAME_EN_COL = 10  # 故障英文名称
        
        # 解析每一行
        loaded_count = 0
        current_register = None
        current_register_name = None
        
        for i in range(start_row, len(df)):
            row = df.iloc[i]
            
            # 获取寄存器偏移量（只有每个寄存器的第一行有值）
            offset_value = row[OFFSET_COL] if OFFSET_COL < len(row) else None
            if pd.notna(offset_value):
                try:
                    offset = int(float(offset_value))
                    if 0 <= offset <= 100:  # 合理的偏移量范围
                        current_register = 20650 + offset
                        
                        # 获取寄存器名称
                        if REG_NAME_COL < len(row) and pd.notna(row[REG_NAME_COL]):
                            current_register_name = str(row[REG_NAME_COL]).strip()
                        
                        if current_register not in MAIN_FAULT_CODES:
                            MAIN_FAULT_CODES[current_register] = {}
                except (ValueError, TypeError):
                    pass
            
            # 获取bit位
            bit_value = row[BIT_COL] if BIT_COL < len(row) else None
            if pd.isna(bit_value):
                continue
            
            bit_str = str(bit_value).strip().lower()
            if not bit_str.startswith('bit'):
                continue
            
            # 提取bit数字
            bit_match = re.search(r'bit(\d+)', bit_str)
            if not bit_match:
                continue
            
            bit = int(bit_match.group(1))
            
            # 如果还没确定寄存器，跳过
            if current_register is None:
                continue
            
            # 获取故障名称
            name_cn = str(row[NAME_CN_COL]).strip() if NAME_CN_COL < len(row) and pd.notna(row[NAME_CN_COL]) else ''
            name_en = str(row[NAME_EN_COL]).strip() if NAME_EN_COL < len(row) and pd.notna(row[NAME_EN_COL]) else ''
            
            # 跳过空名称或保留位
            if not name_cn or name_cn.lower() == 'rsvd' or name_cn == 'nan':
                continue
            
            # 生成故障码
            fault_code = f'E{current_register}_{bit}'
            
            # 确定故障类型
            fault_type = '告警'
            if '故障' in name_cn or 'fault' in name_en.lower():
                fault_type = '故障'
            elif '提示' in name_cn or 'info' in name_en.lower():
                fault_type = '提示'
            
            # 添加到故障码字典（首次定义优先，不覆盖）
            if current_register not in MAIN_FAULT_CODES:
                MAIN_FAULT_CODES[current_register] = {}
            
            if bit not in MAIN_FAULT_CODES[current_register]:
                MAIN_FAULT_CODES[current_register][bit] = (
                    fault_code,
                    name_cn,
                    name_en,
                    fault_type
                )
                loaded_count += 1
        
        # 合并所有故障码
        ALL_FAULT_CODES = {}
        ALL_FAULT_CODES.update(USER_FAULT_CODES)
        ALL_FAULT_CODES.update(MAIN_FAULT_CODES)
        
        print(f"[INFO] 从点表加载了 {loaded_count} 个故障码定义，{len(MAIN_FAULT_CODES)} 个寄存器")
        return loaded_count > 0
        
    except Exception as e:
        print(f"[ERROR] 加载点表失败: {e}")
        import traceback
        traceback.print_exc()
        return False


def parse_bit_list(value) -> List[int]:
    """
    解析bit位列表
    
    Args:
        value: 原始值，可能是字符串 '[1,2,3]' 或列表 [1,2,3] 或数字
        
    Returns:
        bit位列表
    """
    if value is None:
        return []
    
    if isinstance(value, (list, tuple)):
        return list(value)
    
    if isinstance(value, str):
        try:
            if value.startswith('['):
                return ast.literal_eval(value)
            return [int(value)]
        except:
            return []
    
    if isinstance(value, (int, float)):
        return [int(value)]
    
    return []


def parse_register_bits(value: int) -> List[int]:
    """
    解析寄存器值，返回所有置位的bit位
    
    Args:
        value: 寄存器值（整数）
        
    Returns:
        bit位列表
    """
    if value is None or value == 0:
        return []
    
    bits = []
    for i in range(16):  # 最多支持16位
        if value & (1 << i):
            bits.append(i)
    
    return bits


def get_fault_info(register: int, bit: int, use_excel: bool = True) -> Optional[FaultInfo]:
    """
    根据寄存器地址和bit位获取故障信息
    
    Args:
        register: 寄存器地址
        bit: bit位
        use_excel: 是否使用Excel加载的故障码定义
        
    Returns:
        FaultInfo对象，如果未找到返回None
    """
    # 优先使用Excel加载的定义
    fault_dict = ALL_FAULT_CODES if use_excel and ALL_FAULT_CODES else USER_FAULT_CODES
    
    if register not in fault_dict:
        return None
    
    if bit not in fault_dict[register]:
        return None
    
    code, name_cn, name_en, fault_type = fault_dict[register][bit]
    register_name = REGISTER_NAMES.get(register, f'主故障码({register})')
    
    return FaultInfo(
        code=code,
        name_cn=name_cn,
        name_en=name_en,
        fault_type=fault_type,
        register=register,
        bit=bit,
        register_name=register_name
    )


def parse_fault_value(register: int, value, use_excel: bool = True) -> List[FaultInfo]:
    """
    解析故障值，返回所有故障信息
    
    Args:
        register: 寄存器地址
        value: 故障值（可能是字符串列表、列表或数字）
        use_excel: 是否使用Excel加载的故障码定义
        
    Returns:
        故障信息列表
    """
    # 解析bit位
    if isinstance(value, (int, float)):
        bits = parse_register_bits(int(value))
    else:
        bits = parse_bit_list(value)
    
    faults = []
    
    for bit in bits:
        fault_info = get_fault_info(register, bit, use_excel)
        if fault_info:
            faults.append(fault_info)
    
    return faults


def parse_fault_value_simple(register: int, value) -> str:
    """
    简单解析故障值，返回故障描述字符串
    
    Args:
        register: 寄存器地址
        value: 故障值
        
    Returns:
        故障描述字符串，如 'E402:电池BMS通信异常, E404:电池DCDC运行异常'
    """
    faults = parse_fault_value(register, value)
    
    if not faults:
        return ''
    
    return ', '.join([f'{f.code}:{f.name_cn}' for f in faults])


def get_register_name(register: int) -> str:
    """
    获取寄存器名称
    
    Args:
        register: 寄存器地址
        
    Returns:
        寄存器名称
    """
    if register in REGISTER_NAMES:
        return REGISTER_NAMES[register]
    
    # 检查是否是主故障码范围
    if 20650 <= register <= 20689:
        offset = register - 20650
        return f'PCS主故障码({register})'
    
    return f'寄存器{register}'


def get_fault_category(code: str) -> str:
    """
    根据故障码获取故障分类
    
    Args:
        code: 故障码，如 E402
        
    Returns:
        故障分类名称
    """
    if len(code) >= 2:
        prefix = code[:2]
        return FAULT_CATEGORIES.get(prefix, '未知分类')
    return '未知分类'


def get_supported_registers() -> Dict[int, str]:
    """
    获取所有支持的寄存器列表
    
    Returns:
        寄存器地址到名称的映射
    """
    registers = {}
    
    # 添加用户展示故障码
    for addr, name in REGISTER_NAMES.items():
        registers[addr] = name
    
    # 添加主故障码
    for addr in range(20650, 20690):
        if addr not in registers:
            registers[addr] = get_register_name(addr)
    
    return registers


# 初始化时尝试加载Excel（支持多种格式）
# 优先使用"平台逆变器错误代码对照表.xlsx"（点表格式），因为它包含完整的bit定义
try:
    # 先尝试点表格式（平台逆变器错误代码对照表.xlsx）
    print("[INFO] 尝试加载点表格式（平台逆变器错误代码对照表.xlsx）...")
    loaded = load_fault_codes_from_point_table()
    
    # 如果失败，再尝试故障码定义表格式（户储光伏平台产品客户故障码.xlsx）
    total_bits = sum(len(bits) for bits in MAIN_FAULT_CODES.values())
    if not loaded or total_bits < 50:
        print("[INFO] 点表格式加载失败或不完整，尝试加载故障码定义表格式...")
        load_fault_codes_from_excel()
    
    # 最终统计
    total_bits = sum(len(bits) for bits in MAIN_FAULT_CODES.values())
    print(f"[INFO] 故障码定义加载完成: {len(MAIN_FAULT_CODES)} 个寄存器, {total_bits} 个故障码")
    
except Exception as e:
    print(f"[WARN] 初始化加载故障码定义失败: {e}")
    # 确保ALL_FAULT_CODES被初始化
    ALL_FAULT_CODES = {}
    ALL_FAULT_CODES.update(USER_FAULT_CODES)


# ===== 测试代码 =====
if __name__ == '__main__':
    print('=== 故障码解析测试 ===')
    
    # 显示支持的寄存器
    print('\n支持的寄存器：')
    registers = get_supported_registers()
    for addr in sorted(registers.keys()):
        print(f'  {addr}: {registers[addr]}')
    
    # 显示已加载的故障码
    print(f'\n已加载的故障码定义：')
    print(f'  用户展示故障码寄存器: {len(USER_FAULT_CODES)}')
    print(f'  主故障码寄存器: {len(MAIN_FAULT_CODES)}')
    print(f'  总故障码寄存器: {len(ALL_FAULT_CODES)}')
    
    # 测试解析
    print('\n=== 测试解析 ===')
    
    # 测试1: 解析用户展示故障码
    faults = parse_fault_value(20715, 6)  # bit6 = 电池过压
    print(f'\n用户展示故障码 20715 bit6:')
    for f in faults:
        print(f'  {f.code}: {f.name_cn} ({f.fault_type})')
    
    # 测试2: 解析整数寄存器值
    faults = parse_fault_value(20712, 0b000000011)  # bit0和bit1
    print(f'\n用户展示故障码 20712 值=0b000000011:')
    for f in faults:
        print(f'  {f.code}: {f.name_cn} ({f.fault_type})')
    
    # 测试3: 解析主故障码（如果Excel已加载）
    if MAIN_FAULT_CODES and any(MAIN_FAULT_CODES.values()):
        print(f'\n主故障码（已加载）:')
        # 找一个有故障码定义的寄存器
        for reg, bits in sorted(MAIN_FAULT_CODES.items()):
            if bits:
                print(f'\n寄存器 {reg}:')
                for bit, info in sorted(bits.items()):
                    print(f'  bit{bit}: {info[0]} - {info[1]}')
                break
    else:
        print('\n[WARN] 主故障码定义未加载，请检查Excel文件')

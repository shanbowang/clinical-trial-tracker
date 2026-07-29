"""
clinical_engine_v3.py - 通用临床试验 Tracker 分析引擎 V3
支持自动字段匹配，适用于不同结构的Tracker数据

特点：
1. 智能字段匹配 - 支持多别名映射
2. 模块化分析 - 每个分析项独立函数，自动跳过缺失字段
3. 生物标志物自动识别 - CCNE1/Kras等
4. 完整性报告 - 显示成功/跳过的分析项
"""

import os
import json
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime
from typing import Dict, Optional, List, Tuple
import warnings
warnings.filterwarnings('ignore')

import platform
from io import BytesIO
import zipfile
from xml.etree import ElementTree as ET


def _read_xlsx_raw(file_path, sheet_name=0):
    """
    终极兜底：直接从 xlsx（ZIP）中提取纯数据，完全绕过 openpyxl 样式解析。
    支持 .xlsx 格式。
    """
    SS_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    
    with zipfile.ZipFile(file_path, 'r') as z:
        # 获取sheet列表
        wb_xml = z.read('xl/workbook.xml')
        wb_root = ET.fromstring(wb_xml)
        sheets = []
        for sheet_elem in wb_root.findall(f'.//{{{SS_NS}}}sheet'):
            sheets.append({
                'name': sheet_elem.get('name'),
                'id': sheet_elem.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
            })
        
        # 解析共享字符串
        try:
            sst_xml = z.read('xl/sharedStrings.xml')
            sst_root = ET.fromstring(sst_xml)
            shared_strings = [si.find(f'{{{SS_NS}}}t') for si in sst_root.findall(f'.//{{{SS_NS}}}si')]
            shared_strings = [s.text if s is not None and s.text else '' for s in shared_strings]
        except:
            shared_strings = []
        
        # 选择目标sheet
        if isinstance(sheet_name, int):
            if sheet_name >= len(sheets):
                return pd.DataFrame()
            target_sheet = sheets[sheet_name]
        else:
            target_sheet = next((s for s in sheets if s['name'] == sheet_name), None)
            if target_sheet is None:
                return pd.DataFrame()
        
        # 读取sheet数据
        sheet_xml = z.read(f'xl/worksheets/sheet{sheets.index(target_sheet) + 1}.xml')
        sheet_root = ET.fromstring(sheet_xml)
        
        rows = []
        for row_elem in sheet_root.findall(f'.//{{{SS_NS}}}row'):
            row_data = {}
            for cell_elem in row_elem.findall(f'{{{SS_NS}}}c'):
                ref = cell_elem.get('r')  # e.g., 'A1', 'B3'
                col_letter = ''.join(c for c in ref if c.isalpha())
                col_idx = sum((ord(c) - ord('A') + 1) * (26 ** i) for i, c in enumerate(reversed(col_letter))) - 1
                value_elem = cell_elem.find(f'{{{SS_NS}}}v')
                
                if value_elem is not None and value_elem.text:
                    cell_type = cell_elem.get('t', '')
                    if cell_type == 's':  # shared string
                        idx = int(value_elem.text)
                        val = shared_strings[idx] if idx < len(shared_strings) else ''
                    else:
                        val = value_elem.text
                    row_data[col_idx] = val
            
            rows.append(row_data)
        
        if not rows:
            return pd.DataFrame()
        
        # 转换为DataFrame
        max_col = max(max(r.keys()) for r in rows if r) if any(rows) else 0
        data = []
        for row in rows:
            data.append([row.get(i, None) for i in range(max_col + 1)])
        
        # 第一行作为列名
        columns = [str(c) if c is not None else f'Col{i}' for i, c in enumerate(data[0])]
        df = pd.DataFrame(data[1:], columns=columns)
        return df


def setup_chinese_font():
    system = platform.system()
    if system == "Windows":
        plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'KaiTi', 'FangSong']
    elif system == "Darwin":
        plt.rcParams['font.sans-serif'] = ['PingFang SC', 'STHeiti', 'Heiti SC']
    else:
        try:
            import glob as _glob
            # 清除 matplotlib 字体缓存，确保新安装的字体被识别
            for cache_file in _glob.glob(os.path.join(matplotlib.get_cachedir(), 'fontlist*')):
                try:
                    os.remove(cache_file)
                except:
                    pass
        except:
            pass
        plt.rcParams['font.sans-serif'] = ['WenQuanYi Micro Hei', 'WenQuanYi Zen Hei', 'Noto Sans CJK SC', 'DejaVu Sans']
    plt.rcParams['axes.unicode_minus'] = False
    plt.rcParams['figure.dpi'] = 150

setup_chinese_font()
sns.set_style("whitegrid")

COLOR_PALETTES = {
    'primary': ['#48CAE4', '#90E0EF', '#00B4D8', '#0077B6', '#023E8A'],
    'health': ['#95D5B2', '#74C69D', '#52B788', '#40916C', '#2D6A4F'],
    'accent': ['#FFB703', '#FB8500', '#F77F00', '#DC2F02', '#6A040F'],
    'purple': ['#E0AAFF', '#C77DFF', '#9D4EDD', '#7B2CBF', '#5A189A'],
    'warm': ['#FFCDB2', '#FFB4A2', '#FF8BA7', '#FF6D60', '#F72585'],
}

# ==================== 字段别名映射表 ====================
FIELD_ALIASES = {
    'Site': ['Site No.', 'Site No', 'Site', '中心', 'site no.'],
    'Subject': ['Subject ID', 'Subject_ID', '受试者', 'subject id'],
    'Phase': ['Phase', '阶段', 'phase'],
    'Cohort': ['Cohort(mg)', 'Cohort', '剂量组', 'cohort(mg)'],
    'Tumor': ['Tumor', '肿瘤', '肿瘤类型', 'tumor'],
    'Date_ICF': ['Date ICF', 'Date_ICF', '知情日期', 'date icf'],
    'C1D1': ['C1D1', '首次给药', 'c1d1'],
    'Status': ['Status', '状态', 'status'],
    'Best_Response': ['Best Response', 'Best_Response', '疗效', 'best response'],
    'EOT_Date': ['EOT date', 'EOT_Date', '终止日期', 'eot date'],
    'EOT_Reason': ['EOT reason', 'EOT_Reason', '终止原因', 'eot reason'],
    'EOT_Reason_Type': ['EOT reason type', 'EOT reson type', 'EOT_Reason_Type', 'eot reason type'],
    'Latest_Date': ['Latest Date', 'Latest_Date', '最新日期', 'latest date'],
    'Latest_Cycle': ['Latest Cycle', 'Latest_Cycle', '最新周期', 'latest cycle'],
    'Protocol_Version': ['protocol version', 'Protocol', '方案版本', 'protocol version'],
    'Source': ['Source of Patients', 'from', '来源', '病例来源', 'source of patients'],
    'Screen_Failure_Type': ['Screen Failure Type', 'Screen Failure type', '筛选失败类型'],
    'Screen_Failure_Reason': ['Screen Failure Reason', '筛选失败原因'],
    # 生物标志物 - CCNE1
    'CCNE1_Result': ['CCNE1 result', 'CCNE1_result', 'ccne1 result'],
    'CCNE1_Send_Date': ['Tumor slide send out date', '肿瘤切片送出日期'],
    'CCNE1_Report_Date': ['CCNE1 result report date', 'CCNE1_report_date'],
    # 生物标志物 - Kras
    'Kras_Result': ['Kras result', 'Kras_result', 'kras result'],
    'Kras_Report_Date': ['Kras result report date', 'Kras_report_date'],
    # 生物标志物 - 通用
    'Biomarker_Result': ['CCNE1 result', 'Kras result', '检测结果'],
    'Biomarker_Send_Date': ['Tumor slide send out date'],
    'Biomarker_Report_Date': ['CCNE1 result report date', 'Kras result report date'],
}

# ==================== 工具函数 ====================

def find_column(df_columns, *aliases) -> Optional[str]:
    """查找第一个匹配的列名"""
    for col in df_columns:
        col_clean = str(col).strip().lower().replace(' ', '').replace('_', '').replace('.', '')
        for alias in aliases:
            alias_clean = str(alias).strip().lower().replace(' ', '').replace('_', '').replace('.', '')
            if alias_clean == col_clean or alias_clean in col_clean or col_clean in alias_clean:
                return col
    return None

def find_all_columns(df_columns, *aliases) -> list:
    """查找所有匹配的列名"""
    matches = []
    for col in df_columns:
        col_clean = str(col).strip().lower().replace(' ', '').replace('_', '').replace('.', '')
        for alias in aliases:
            alias_clean = str(alias).strip().lower().replace(' ', '').replace('_', '').replace('.', '')
            if alias_clean == col_clean or alias_clean in col_clean or col_clean in alias_clean:
                matches.append(col)
                break
    return matches

def find_best_column(df, key: str) -> Optional[str]:
    """查找最佳匹配列：优先精确匹配，其次多列时选非空最多的"""
    if key not in FIELD_ALIASES:
        return None
    matches = find_all_columns(df.columns, *FIELD_ALIASES[key])
    if not matches:
        return None
    if len(matches) == 1:
        return matches[0]
    
    def _match_quality(col, aliases):
        """匹配质量：精确匹配=2，包含匹配=1"""
        col_clean = str(col).strip().lower().replace(' ', '').replace('_', '').replace('.', '')
        for alias in aliases:
            alias_clean = str(alias).strip().lower().replace(' ', '').replace('_', '').replace('.', '')
            if alias_clean == col_clean:
                return 2
            if alias_clean in col_clean or col_clean in alias_clean:
                return 1
        return 0
    
    # 按匹配质量排序，同质量按非空数排序
    aliases = FIELD_ALIASES[key]
    scored = [(m, _match_quality(m, aliases), df[m].notna().sum()) for m in matches]
    scored.sort(key=lambda x: (-x[1], -x[2]))  # 质量优先，然后非空数
    return scored[0][0]

def find_column_by_key(df_columns, key: str) -> Optional[str]:
    """通过预定义的key查找列名（简单版，不区分多匹配）"""
    if key in FIELD_ALIASES:
        return find_column(df_columns, *FIELD_ALIASES[key])
    return None

def standardize_tumor_name(tumor):
    """标准化肿瘤名称"""
    if pd.isna(tumor):
        return tumor
    tumor_lower = str(tumor).lower().strip()
    if 'breast' in tumor_lower or 'brest' in tumor_lower:
        return 'Breast cancer'
    elif 'ovarian' in tumor_lower:
        return 'Ovarian cancer'
    elif 'endometrial' in tumor_lower:
        return 'Endometrial cancer'
    elif 'prostate' in tumor_lower or '前列腺' in tumor_lower:
        return 'Prostate cancer'
    elif 'pancreatic' in tumor_lower or '胰腺' in tumor_lower:
        return 'Pancreatic cancer'
    elif 'colon' in tumor_lower or '结肠' in tumor_lower:
        return 'Colon cancer'
    elif 'lung' in tumor_lower or '肺' in tumor_lower:
        return 'Lung cancer'
    return tumor

def save_chart(fig, output_dir, filename):
    path = os.path.join(output_dir, "charts", filename)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    fig.savefig(path, dpi=150, bbox_inches='tight')
    plt.close(fig)
    return path

def save_table(df, output_dir, filename):
    path = os.path.join(output_dir, "tables", filename)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    df.to_csv(path, index=True)
    return path

# ==================== 字段检测类 ====================

class FieldDetector:
    """字段检测器"""
    
    def __init__(self, df: pd.DataFrame):
        self.df = df
        self.columns = df.columns.tolist()
        self.found_fields = {}
        self.missing_fields = []
        self._detect_all()
    
    def _detect_all(self):
        """检测所有预定义字段（多列匹配时选非空最多的）"""
        for key in FIELD_ALIASES:
            col = find_best_column(self.df, key)
            if col:
                self.found_fields[key] = col
            else:
                self.missing_fields.append(key)
    
    def get(self, key: str) -> Optional[str]:
        """获取字段名"""
        return self.found_fields.get(key)
    
    def has(self, *keys) -> bool:
        """检查字段是否存在"""
        return all(key in self.found_fields for key in keys)
    
    def has_any(self, *keys) -> bool:
        """检查任意字段是否存在"""
        return any(key in self.found_fields for key in keys)
    
    def get_biomarker_type(self) -> Optional[str]:
        """检测生物标志物类型"""
        if 'CCNE1_Result' in self.found_fields:
            return 'CCNE1'
        elif 'Kras_Result' in self.found_fields:
            return 'Kras'
        return None
    
    def report(self) -> Dict:
        """生成字段检测报告"""
        return {
            'found': self.found_fields,
            'missing': self.missing_fields,
            'biomarker_type': self.get_biomarker_type()
        }
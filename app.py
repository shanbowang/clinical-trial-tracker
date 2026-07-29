"""
app.py - 临床试验 Tracker 分析 Web 应用（V3）
"""

import streamlit as st
import sys

st.set_page_config(
    page_title="临床试验 Tracker 分析系统",
    page_icon="📊",
    layout="wide"
)

st.title("📊 临床试验 Tracker 分析系统")
st.markdown(f"**Python**: {sys.version} | **Platform**: {sys.platform}")

# 逐步检测依赖
st.subheader("环境检测")

import_results = {}

# 1. pandas
try:
    import pandas as pd
    import_results['pandas'] = f"✅ {pd.__version__}"
except Exception as e:
    import_results['pandas'] = f"❌ {e}"

# 2. numpy
try:
    import numpy as np
    import_results['numpy'] = f"✅ {np.__version__}"
except Exception as e:
    import_results['numpy'] = f"❌ {e}"

# 3. matplotlib
try:
    import matplotlib
    matplotlib.use('Agg')
    import matplotlib.pyplot as plt
    import_results['matplotlib'] = f"✅ {matplotlib.__version__}"
except Exception as e:
    import_results['matplotlib'] = f"❌ {e}"

# 4. openpyxl
try:
    import openpyxl
    import_results['openpyxl'] = f"✅ {openpyxl.__version__}"
except Exception as e:
    import_results['openpyxl'] = f"❌ {e}"

# 5. seaborn
try:
    import seaborn as sns
    import_results['seaborn'] = f"✅ {sns.__version__}"
except Exception as e:
    import_results['seaborn'] = f"⚠️ 不可用: {e}"

# 6. clinical_engine_v3
try:
    from clinical_engine_v3 import run_tracker_analysis, setup_chinese_font, load_tracker_data, validate_tracker_data
    import_results['clinical_engine_v3'] = "✅ 导入成功"
except Exception as e:
    import_results['clinical_engine_v3'] = f"❌ {e}"

for name, status in import_results.items():
    st.text(f"{status}")

# 字体检测
st.subheader("字体检测")
try:
    setup_chinese_font()
    import matplotlib.font_manager as fm
    fonts = [f.name for f in fm.fontManager.ttflist]
    cjk_fonts = [f for f in fonts if any(k in f for k in ['WenQuanYi', 'Noto', 'CJK', 'SimHei', 'Micro', 'Hei'])]
    if cjk_fonts:
        st.success(f"发现中文字体: {', '.join(cjk_fonts[:10])}")
    else:
        st.warning("未发现中文字体，将使用默认字体")
except Exception as e:
    st.warning(f"字体检测失败: {e}")

st.success("✅ 基础环境检测完成")
st.info("如果以上检测全部正常，请联系我恢复完整功能")

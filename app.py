"""
app.py - 临床试验 Tracker 分析 Web 应用（V3）
基于 clinical_engine_v3.py 的通用分析引擎

运行: streamlit run app.py
"""

import streamlit as st
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import os
from datetime import datetime
import warnings
import zipfile
import io
warnings.filterwarnings('ignore')

try:
    from clinical_engine_v3 import run_tracker_analysis, setup_chinese_font, load_tracker_data, validate_tracker_data
except Exception as _import_err:
    st.set_page_config(page_title="临床试验 Tracker 分析系统", page_icon="📊", layout="wide")
    st.error(f"❌ 导入分析引擎失败：{_import_err}")
    import traceback
    st.code(traceback.format_exc())
    st.stop()

def _fmt_rows(rows, max_show=5):
    """格式化行号列表"""
    if not rows:
        return ""
    if len(rows) <= max_show:
        return f"第{', '.join(str(r+2) for r in rows)}行"
    return f"第{', '.join(str(r+2) for r in rows[:max_show])}行等共{len(rows)}行"

st.set_page_config(
    page_title="临床试验 Tracker 分析系统",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

setup_chinese_font()

st.sidebar.header("📊 定义说明")
st.sidebar.markdown("""
**核心指标定义：**
- **筛选数**: 有 Date ICF（知情同意日期）
- **入组数**: 有 C1D1（首次给药日期）
- **筛选失败**: Status = 'SF' 或 'Screen Failure'
- **筛选成功率**: 入组数 / 筛选数 × 100%

**分析维度（14项）：**
1. Phase & 剂量组筛选入组表
2. 中心筛选入组表
3. Phase & 肿瘤类型筛选入组表
4. Phase & 状态分布表
5. 月度筛选入组趋势图
6. 筛选失败原因分布
7. 生物标志物检测分析（自动识别CCNE1/Kras）
8. 在组时间游泳图
9. EOT 原因分布
10. 疗效数据表
11. 方案版本筛选入组
12. 不同来源分析
13. 各中心筛选时长
14. 风险预警看板
""")

st.sidebar.header("📁 支持的格式")
st.sidebar.markdown("- Excel: .xlsx, .xls\n- CSV: .csv")
st.sidebar.markdown("- 自动检测多Sheet中的有效数据表")

st.title("📊 临床试验 Tracker 分析系统（V3）")
st.markdown("""
上传临床试验 Tracker 数据文件，系统将自动进行 **14 项分析**，智能匹配字段，生成统计表格和可视化图表。

> V3 特性：智能字段匹配 | 多Sheet自动检测 | 生物标志物自动识别 | 模块化分析
""")

st.header("1️⃣ 上传数据文件")

uploaded_file = st.file_uploader(
    "选择要分析的文件",
    type=['csv', 'xlsx', 'xls'],
    help="支持 CSV 或 Excel 格式的临床试验 Tracker 数据，自动检测多Sheet"
)

if uploaded_file is None:
    st.info("👆 请上传文件开始分析")
    st.stop()

st.header("2️⃣ 数据预览")

try:
    df = load_tracker_data(uploaded_file)
    
    st.success(f"✅ 成功读取文件：{uploaded_file.name}（{len(df)} 行 × {len(df.columns)} 列）")
    
    col1, col2 = st.columns(2)
    with col1:
        st.metric("总记录数", len(df))
    with col2:
        st.metric("字段数", len(df.columns))
    
    with st.expander("📋 查看原始数据（前 20 行）"):
        st.dataframe(df.head(20), use_container_width=True)
    
    with st.expander("📋 查看所有字段"):
        st.write(list(df.columns))

except Exception as e:
    st.error(f"❌ 读取文件失败：{e}")
    st.stop()

st.header("3️⃣ 数据校验")

# 运行数据校验
validation = validate_tracker_data(df)

errs = validation.get('errors', [])
warns = validation.get('warnings', [])
vsum = validation.get('summary', {})

if len(errs) == 0 and len(warns) == 0:
    st.success("✅ 数据校验通过，未发现任何问题。")
else:
    # 错误（阻断级）
    if errs:
        st.error(f"### 🚫 发现 {len(errs)} 项严重问题（建议修复后再分析）")
        for e in errs:
            cat = e.get('category', '')
            rows = e.get('rows', [])
            row_text = _fmt_rows(rows)
            st.markdown(f"- **[{cat}]** {e['message']} ({row_text})")
    
    # 警告
    if warns:
        st.warning(f"### ⚠️ 发现 {len(warns)} 项注意问题")
        for w in warns:
            cat = w.get('category', '')
            rows = w.get('rows', [])
            row_text = _fmt_rows(rows)
            st.markdown(f"- **[{cat}]** {w['message']}" + (f" ({row_text})" if row_text else ""))

# 确认框
has_errors = vsum.get('has_blocking', False)
if has_errors:
    acknowledge = st.checkbox(
        "我已确认上述数据问题，仍然继续分析（建议先修复数据后再分析）",
        value=False
    )
else:
    acknowledge = True

st.header("4️⃣ 运行分析")

col_btn, col_spacer = st.columns([1, 3])
with col_btn:
    analyze_disabled = not acknowledge
    analyze_clicked = st.button(
        "🚀 开始分析",
        type="primary",
        use_container_width=True,
        disabled=analyze_disabled
    )

if analyze_clicked:
    with st.spinner("正在分析数据..."):
        try:
            output_dir = os.path.join(os.path.dirname(__file__), "output")
            os.makedirs(output_dir, exist_ok=True)
            
            result = run_tracker_analysis(df, output_dir)
            
            st.success(f"✅ 分析完成！（成功 {len(result['executed'])} 项，跳过 {len(result['skipped'])} 项）")
            
            st.markdown("---")
            st.subheader("分析结果")
            
            # 汇总指标
            summary = result['summary']
            col1, col2, col3, col4, col5 = st.columns(5)
            with col1:
                st.metric("总筛选数", summary.get('total_screened', 0))
            with col2:
                st.metric("总入组数", summary.get('total_enrolled', 0))
            with col3:
                st.metric("筛选失败数", summary.get('total_screen_failure', 0))
            with col4:
                st.metric("筛选成功率", f"{summary.get('success_rate', 0)}%")
            with col5:
                biomarker = summary.get('biomarker_type', '无')
                st.metric("生物标志物", biomarker if biomarker else '无')
            
            tables = result.get('tables', {})
            charts = result.get('charts', {})
            
            # 分析项序号到显示名称的映射
            analysis_names = [
                (1, "Phase & 剂量组筛选入组表"),
                (2, "中心筛选入组表"),
                (3, "Phase & 肿瘤类型筛选入组表"),
                (4, "Phase & 状态分布表"),
                (5, "月度筛选入组趋势"),
                (6, "筛选失败原因分布"),
                (7, "生物标志物检测分析"),
                (8, "在组时间游泳图"),
                (9, "EOT 原因分布"),
                (10, "疗效数据表"),
                (11, "方案版本筛选入组"),
                (12, "不同来源分析"),
                (13, "各中心筛选时长"),
                (14, "风险预警看板"),
            ]
            
            name_to_num = {name: num for num, name in analysis_names}
            
            # 按编号排序显示成功执行的分析
            executed = result.get('executed', [])
            
            # 为排序创建映射
            def get_sort_key(name):
                for num, ana_name in analysis_names:
                    if ana_name == name:
                        return num
                return 99
            
            sorted_executed = sorted(executed, key=get_sort_key)
            
            for name in sorted_executed:
                num = name_to_num.get(name, 0)
                st.subheader(f"{num}. {name}")
                
                # 显示对应的表格
                matched_tables = {k: v for k, v in tables.items() if name in k or name.split(' ')[0] in k}
                for tkey, tval in matched_tables.items():
                    if isinstance(tval, pd.DataFrame):
                        st.dataframe(tval, use_container_width=True)
                
                # 显示对应的图表
                if name in charts:
                    st.image(charts[name], use_container_width=True)
            
            # 生物标志物详情（图表已在上面分析项中展示，这里仅展示指标和分布表）
            if 'biomarker' in summary:
                biomarker = summary['biomarker']
                btype = biomarker.get('type', '')
                if btype:
                    st.markdown(f"**检测类型：{btype}**")
                
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("样本数", biomarker.get('samples', 'N/A'))
                with col2:
                    st.metric("平均天数", biomarker.get('mean_days', 'N/A'))
                with col3:
                    st.metric("中位天数", biomarker.get('median_days', 'N/A'))
                
                if 'result_table' in biomarker and isinstance(biomarker['result_table'], pd.DataFrame):
                    st.dataframe(biomarker['result_table'], use_container_width=True)
            
            # 跳过项
            skipped = result.get('skipped', [])
            if skipped:
                with st.expander(f"⚠️ 跳过的分析项（{len(skipped)} 项）"):
                    for item in skipped:
                        st.markdown(f"- **{item['name']}**：{item['reason']}")
            
            # 字段检测报告
            if 'fields_detected' in summary:
                with st.expander(f"🔍 字段检测报告（{len(summary['fields_detected'])} 个已识别字段）"):
                    st.write(summary['fields_detected'])
            
            st.header("5️⃣ 下载报告")
            
            charts_dir = os.path.join(output_dir, "charts")
            tables_dir = os.path.join(output_dir, "tables")
            
            all_files = []
            if os.path.exists(charts_dir):
                all_files.extend([(os.path.join(charts_dir, f), f"charts/{f}") 
                                  for f in os.listdir(charts_dir) if f.endswith('.png')])
            if os.path.exists(tables_dir):
                all_files.extend([(os.path.join(tables_dir, f), f"tables/{f}") 
                                  for f in os.listdir(tables_dir) if f.endswith('.csv')])
            
            if all_files:
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
                    for file_path, arc_name in all_files:
                        zf.write(file_path, arc_name)
                zip_buffer.seek(0)
                
                st.download_button(
                    label="📥 下载所有报告（ZIP压缩包）",
                    data=zip_buffer,
                    file_name=f"tracker_analysis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                    mime="application/zip",
                    use_container_width=True
                )
                
                st.caption(f"包含 {len([f for f in all_files if f[1].startswith('charts/')])} 个图表文件、"
                          f"{len([f for f in all_files if f[1].startswith('tables/')])} 个表格文件")
            else:
                st.warning("没有生成报告文件")
            
        except Exception as e:
            st.error(f"❌ 分析过程出错：{e}")
            import traceback
            st.code(traceback.format_exc())

st.markdown("---")
st.markdown("""
<div style="text-align: center; color: #666;">
    <p>临床试验 Tracker 分析系统 V3 | 基于 clinical_engine_v3 通用分析引擎</p>
    <p>智能字段匹配 | 多Sheet自动检测 | 生物标志物自动识别</p>
</div>
""", unsafe_allow_html=True)

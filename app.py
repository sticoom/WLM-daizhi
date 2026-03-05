# app.py
import streamlit as st
from step1_framework import step1_add_new_rows
from step2_fill import step2_fill_and_calculate

st.set_page_config(page_title="沃尔玛库存 - 解耦自动化流水线", layout="wide")
st.title("🧩 沃尔玛库存更新流水线 (模块化解耦版)")
st.markdown("""
本系统实现了完全解耦的 Agent Skill 架构设计，由三个文件无缝驱动：
1. **模块一 (`step1_framework.py`)**：负责解析历史边界、对比表单、创建新框架并滤除死数据。
2. **模块二 (`step2_fill.py`)**：接收步骤一的成果，独立负责全量数据注入、SKU智能解析及各项复杂公式计算。
3. **主程序 (`app.py`)**：一键式调度管道。
""")

col1, col2 = st.columns(2)
with col1:
    f_inv = st.file_uploader("1. 上传库存明细表 (必选)", type=['xlsx'], key="inv")
with col2:
    f_prod = st.file_uploader("2. 上传产品资料表 (可选,用于提取SKU)", type=['xlsx'], key="prod")

if f_inv and st.button("🚀 启动自动化流水线"):
    with st.status("正在执行自动化工作流...", expanded=True) as status:
        try:
            st.write("⏳ [模块一] 启动：解析框架并追加新记录...")
            intermediate_file, new_count = step1_add_new_rows(f_inv)
            st.write(f"✅ [模块一] 完成：成功过滤死数据，实际新增了 {new_count} 条记录结构。")
            
            st.write("⏳ [模块二] 启动：提取 SKU、注入全量数据与周转计算...")
            final_file = step2_fill_and_calculate(intermediate_file, f_prod)
            st.write("✅ [模块二] 完成：全量数据注入并计算完毕！")
            
            status.update(label="🎉 流水线全部执行成功！", state="complete", expanded=False)
            
            st.download_button(
                label="📥 下载最终计算完成的库存表",
                data=final_file,
                file_name=f"Final_{f_inv.name}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
            st.balloons()
            
        except Exception as e:
            status.update(label="❌ 流水线执行中断", state="error")
            st.error(f"处理发生异常: {e}")

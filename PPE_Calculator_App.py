import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="PPE 申购自动计算器 V2", layout="wide")

st.title("🛡️ 劳防用品 (PPE) 月度申购自动计算器 V2")
st.markdown("根据**最新版《劳防用品配置标准（2026）》**自动计算各部门采购量。您只需修改**人数**和**工作天数**即可。")

# 侧边栏输入参数
st.sidebar.header("参数设置")
days = st.sidebar.number_input("本月工作天数", min_value=1, max_value=31, value=21, step=1)

st.sidebar.subheader("各部门人数")
p_count = st.sidebar.number_input("生产部人数", min_value=0, value=16)
t_count = st.sidebar.number_input("技术部人数", min_value=0, value=12)
q_count = st.sidebar.number_input("品管部人数", min_value=0, value=2)
w_count = st.sidebar.number_input("仓库人数", min_value=0, value=2)

# 最新配置标准逻辑 (基于 2026 PPE -  测试.xlsx)
# 生产部: 口罩 3个/天, 护目镜 1副/月, 耳塞 3副/月, 化学品防护服 2件/月, 棉线手套 12副/月, 乳胶手套 30副/月 (60只), 防化学品手套 2副/月
# 技术部: 口罩 1.5个/天, 护目镜 1副/年 (月度为0), 耳塞 3副/月, 乳胶手套 4.5盒/月 (4.5*100=450只)
# 品管部: 口罩 1个/天, 护目镜 1副/年 (月度为0), 耳塞 1副/月, 乳胶手套 4.5盒/月 (4.5*100=450只)
# 仓库: 口罩 1个/天, 棉线手套 10副/月

data = [
    {
        "部门": "生产部",
        "人数": p_count,
        "工作日": days,
        "头戴式口罩(个)": p_count * days * 3,
        "耳挂式口罩(个)": 0,
        "护目镜(副)": p_count * 1,
        "耳塞(副)": p_count * 3,
        "化学防护服(件)": p_count * 2,
        "棉线手套(副)": p_count * 12,
        "乳胶手套(只)": p_count * 60,
        "防化手套(副)": p_count * 2
    },
    {
        "部门": "技术部",
        "人数": t_count,
        "工作日": days,
        "头戴式口罩(个)": 0,
        "耳挂式口罩(个)": int(t_count * days * 1.5),
        "护目镜(副)": 0,
        "耳塞(副)": t_count * 3,
        "化学防护服(件)": 0,
        "棉线手套(副)": 0,
        "乳胶手套(只)": int(t_count * 4.5 * 100),
        "防化手套(副)": 0
    },
    {
        "部门": "品管部",
        "人数": q_count,
        "工作日": days,
        "头戴式口罩(个)": q_count * days * 1,
        "耳挂式口罩(个)": 0,
        "护目镜(副)": 0,
        "耳塞(副)": q_count * 1,
        "化学防护服(件)": 0,
        "棉线手套(副)": 0,
        "乳胶手套(只)": int(q_count * 4.5 * 100),
        "防化手套(副)": 0
    },
    {
        "部门": "仓库",
        "人数": w_count,
        "工作日": days,
        "头戴式口罩(个)": 0,
        "耳挂式口罩(个)": w_count * days * 1,
        "护目镜(副)": 0,
        "耳塞(副)": 0,
        "化学防护服(件)": 0,
        "棉线手套(副)": w_count * 10,
        "乳胶手套(只)": 0,
        "防化手套(副)": 0
    }
]

df = pd.DataFrame(data)

# 增加合计行
totals = df.select_dtypes(include=['int', 'float']).sum()
totals["部门"] = "合计"
df = pd.concat([df, pd.DataFrame([totals])], ignore_index=True)

st.dataframe(df, use_container_width=True)

# 导出功能
def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='PPE最新核算结果')
    processed_data = output.getvalue()
    return processed_data

excel_data = to_excel(df)

st.download_button(
    label="📥 导出最新计算结果至 Excel",
    data=excel_data,
    file_name="2026_PPE_月度自动核算单.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

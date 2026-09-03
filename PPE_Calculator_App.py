import streamlit as st
import pandas as pd
import io
import datetime

st.set_page_config(page_title="慧瑞 - PPE自动计算器", layout="wide")

# ================= 新增：隐藏右上角菜单和底部水印 =================
hide_streamlit_style = """
            <style>
            #MainMenu {visibility: hidden;} /* 隐藏右上角汉堡菜单 */
            header {visibility: hidden;}    /* 隐藏顶部的 header (含 Deploy 按钮) */
            footer {visibility: hidden;}    /* 隐藏底部的 Streamlit 水印 */
            </style>
            """
st.markdown(hide_streamlit_style, unsafe_allow_html=True)
# ===============================================================

# ================= 新增：Logo 显示区 =================
# 将公司的 logo 图片命名为 logo.png，并与此 app.py 放在同一个文件夹或 GitHub 仓库中
try:
    st.sidebar.image("logo.png", use_container_width=True)
except Exception:
    st.sidebar.caption("（提示：在同目录下放入 logo.png 即可在此处显示公司标识）")
st.sidebar.markdown("---")
# ===================================================

# 1. 修改了网页主标题
st.title("🛡️ 慧瑞 - 劳防用品 (PPE) 月度申购自动计算器")
st.markdown("根据**最新版《劳防用品配置标准（2026）》**自动计算各部门采购量。")

# 侧边栏输入参数
st.sidebar.header("参数设置")
days = st.sidebar.number_input("本月常规工作天数", min_value=1, max_value=31, value=21, step=1)

st.sidebar.subheader("各部门人数")
p_count = st.sidebar.number_input("生产部人数", min_value=0, value=16)
t_count = st.sidebar.number_input("技术部人数", min_value=0, value=12)
q_count = st.sidebar.number_input("品管部人数", min_value=0, value=2)
w_count = st.sidebar.number_input("仓库人数", min_value=0, value=2)
z_count = st.sidebar.number_input("驻外人数", min_value=0, value=9)

# 最新配置标准逻辑核算
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
    },
    # 驻外逻辑修改：工作日固定25天，最终数量自动乘以2 (2个月的量)
    {
        "部门": "驻外 (按2个月计)",
        "人数": z_count,
        "工作日": 25, 
        "头戴式口罩(个)": 0,
        "耳挂式口罩(个)": z_count * 25 * 1 * 2,    
        "护目镜(副)": 0,
        "耳塞(副)": z_count * 3 * 2,               
        "化学防护服(件)": 0,
        "棉线手套(副)": 0,
        "乳胶手套(只)": z_count * 1 * 100 * 2,     
        "防化手套(副)": 0
    }
]

df = pd.DataFrame(data)

# 增加合计行
totals = df.select_dtypes(include=['int', 'float']).sum()
totals["部门"] = "合计"
df = pd.concat([df, pd.DataFrame([totals])], ignore_index=True)

st.dataframe(df, use_container_width=True)

st.markdown("<br>", unsafe_allow_html=True)

# ================= 导出功能区 =================
col1, col2 = st.columns(2)

# 1. Excel 导出功能
def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='PPE最新核算结果')
    return output.getvalue()

excel_data = to_excel(df)

with col1:
    st.download_button(
        label="📥 导出至 Excel",
        data=excel_data,
        file_name="慧瑞_PPE_月度自动核算单.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

# 2. 打印与 PDF 导出功能
def to_printable_html(df, days):
    current_date = datetime.datetime.now().strftime("%Y-%m-%d")
    html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <title>慧瑞 - 劳防用品月度申购单</title>
        <style>
            body {{ font-family: 'Microsoft YaHei', sans-serif; padding: 20px; color: #333; }}
            h2 {{ text-align: center; border-bottom: 2px solid #333; padding-bottom: 10px; }}
            .info {{ display: flex; justify-content: space-between; margin-bottom: 15px; font-weight: bold; }}
            table {{ border-collapse: collapse; width: 100%; font-size: 14px; margin-bottom: 30px; }}
            th, td {{ border: 1px solid #666; padding: 8px; text-align: center; }}
            th {{ background-color: #f2f2f2; }}
            .footer {{ text-align: center; font-size: 12px; color: gray; margin-top: 50px; }}
            @media print {{
                .no-print {{ display: none; }}
                body {{ padding: 0; }}
            }}
        </style>
    </head>
    <body onload="window.print()">
        <div class="no-print" style="text-align: center; margin-bottom: 20px;">
            <button onclick="window.print()" style="padding: 10px 20px; font-size: 16px; cursor: pointer; background-color: #4CAF50; color: white; border: none; border-radius: 5px;">
                🖨️ 点击打印或另存为 PDF
            </button>
            <p style="font-size: 12px; color: #666;">（提示：在打印设置的“目标打印机”中选择“另存为 PDF”即可导出为 PDF 文件）</p>
        </div>
        
        <h2>慧瑞 - 劳防用品 (PPE) 月度申购单</h2>
        <div class="info">
            <span>核算日期：{current_date}</span>
            <span>本月常规工作日：{days} 天</span>
        </div>
        
        {df.to_html(index=False)}
        
        <div class="info" style="margin-top: 40px;">
            <span>编制人签字：____________________</span>
            <span>审核人签字：____________________</span>
        </div>
        <div class="footer">仅供学习，严禁商业用途，版权所有：Yefei</div>
    </body>
    </html>
    """
    return html.encode('utf-8')

pdf_print_data = to_printable_html(df, days)

with col2:
    st.download_button(
        label="🖨️ 生成打印版 / 导出 PDF",
        data=pdf_print_data,
        file_name="慧瑞_PPE_打印单.html",
        mime="text/html",
        use_container_width=True
    )
# ===================================================

# 3. 增加底部版权声明
st.markdown("---")
st.markdown("<p style='text-align: center; color: gray; font-size: 14px;'>仅供学习，严禁商业用途，版权所有：Yefei</p>", unsafe_allow_html=True)

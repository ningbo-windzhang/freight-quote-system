import os
os.system("pip install pandas==2.2.3")

import streamlit as st
import pandas as pd
from datetime import datetime, timedelta

# 设置网页标题
st.set_page_config(page_title="货代报价系统", layout="centered")

st.markdown("<h1 style='text-align: center;'>货代报价系统</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; color: gray;'>读取 rate_table.xlsx 自动算价 · 改价格只需改 Excel</p>", unsafe_allow_html=True)

# ---------------- 左侧边栏表单 ----------------
with st.sidebar:
    st.header("报价条件")
    origin = st.selectbox("起运港", ["NINGBO", "SHANGHAI"])
    dest = st.selectbox("目的港", ["HAMBURG", "ROTTERDAM", "LOS ANGELES"])
    box_type = st.selectbox("箱型", ["20GP", "40GP", "40HQ"])
    box_qty = st.number_input("箱量", min_value=1, value=1, step=1)
    customer_name = st.text_input("客户名称 (可选)")

# ---------------- 读取 Excel 数据 ----------------
@st.cache_data
def load_data():
    df = pd.read_excel("rate_table.xlsx", sheet_name="海运费")
    return df

df = load_data()

# ---------------- 处理点击事件 ----------------
if st.button("生成报价单", type="primary"):
    
    try:
        # 根据起运港和目的港筛选
        result = df[
            (df['起运港'] == origin) & 
            (df['目的港'] == dest)
        ]
        
        if not result.empty:
            # 获取对应箱型列的价格（例如选20GP，就取20GP那一列的值）
            unit_price = float(result.iloc[0][box_type])
            total_price = unit_price * box_qty
            
            st.success(" 报价计算成功！")
            
            # ---- 显示报价单样式 ----
            st.divider()
            st.subheader("报 价 单")
            
            col1, col2 = st.columns(2)
            with col1:
                st.write(f"**客户：** {customer_name or '未填写'}")
                st.write(f"**航线：** {origin} -> {dest}")
            with col2:
                st.write(f"**报价日期：** {datetime.now().strftime('%Y-%m-%d')}")
                st.write(f"**有效期至：** {(datetime.now() + timedelta(days=7)).strftime('%Y-%m-%d')}")
            
            # ---- 显示明细表格 ----
            st.write("**费用明细**")
            # 组装表格数据
            table_data = [
                {"项目": "海运费", "说明": f"{box_type} x {box_qty}", "金额(USD)": total_price}
            ]
            st.table(table_data)
            
            # ---- 显示合计 ----
            st.markdown(f"### 合计：**${total_price:,.2f} USD**")
            st.caption("目的港费用到付/另议；以上价格仅供参考，以最终确认为准。")
            
        else:
            st.warning(" 未找到该航线报价，请检查 Excel 表格。")
            
    except Exception as e:
        st.error(f"发生错误：{e}")

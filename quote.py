# -*- coding: utf-8 -*-
"""
货代报价系统 + 东南亚运价查询
================================
功能：
  1. 报价单生成（读取 rate_table.xlsx，生成 PDF 报价单）
  2. 东南亚运价查询（读取 southeast_asia_rates.xlsx，英文展示给海外客户）

运行：
  streamlit run quote.py

文件说明（都要放在同一目录 / GitHub 仓库）：
  - quote.py                     本文件
  - rate_table.xlsx              原报价表（海运费 + 杂费）
  - southeast_asia_rates.xlsx    东南亚运价表
  - logo.png                     平头哥 Logo（可选）
"""

import os
import io
import pandas as pd
import streamlit as st
from datetime import datetime

# ============================================================
# 0. 页面基础配置
# ============================================================
st.set_page_config(
    page_title="Freight & Sourcing 报价系统",
    page_icon="🚢",
    layout="wide",
    initial_sidebar_state="expanded",
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

def file_path(name: str) -> str:
    """兼容本地运行与 Streamlit 云端：优先当前目录，其次脚本目录"""
    if os.path.exists(name):
        return name
    p = os.path.join(BASE_DIR, name)
    return p

# ============================================================
# 1. 语言字典（中英切换）
# ============================================================
LANG = {
    "中文": {
        "title": "货代报价系统",
        "nav_quote": "📄 生成报价单",
        "nav_sea": "🌏 东南亚运价查询",
        "pol": "起运港", "pod": "目的港", "type": "箱型", "qty": "箱量",
        "customer": "客户名称（可选）", "gen": "生成报价单",
        "result": "报价结果", "freight": "海运费", "surcharge": "杂费",
        "total": "合计（USD）", "port": "目的港", "line": "航线",
        "vessel": "船名航次", "transit": "直达/中转", "time": "时效",
        "gp20": "20GP", "gp40": "40GP", "hq40": "40HQ",
        "carrier": "船公司", "remark": "备注", "search": "搜索港口/船公司",
        "search_hint": "输入港口或船公司英文名过滤…",
        "download_csv": "下载 CSV", "no_data": "未找到匹配运价",
        "sea_note": "价格单位：USD / 集装箱（20GP·40GP·40HQ）。以下为东南亚航线参考运价，实际以订舱确认为准。",
        "sea_filter": "筛选",
    },
    "English": {
        "title": "Freight Quotation System",
        "nav_quote": "📄 Generate Quotation",
        "nav_sea": "🌏 Southeast Asia Rates",
        "pol": "Port of Loading", "pod": "Port of Discharge", "type": "Container Type",
        "qty": "Quantity", "customer": "Customer (optional)", "gen": "Generate Quotation",
        "result": "Quotation", "freight": "Ocean Freight", "surcharge": "Surcharges",
        "total": "Total (USD)", "port": "Destination", "line": "Route",
        "vessel": "Vessel/Voyage", "transit": "Direct/Transit", "time": "Transit Time",
        "gp20": "20GP", "gp40": "40GP", "hq40": "40HQ",
        "carrier": "Carrier", "remark": "Remark", "search": "Search port/carrier",
        "search_hint": "Type port or carrier name to filter…",
        "download_csv": "Download CSV", "no_data": "No matching rates found",
        "sea_note": "Rates in USD per container (20GP/40GP/40HQ). Reference rates for Southeast Asia; final quote subject to booking confirmation.",
        "sea_filter": "Filter",
    },
}

# ============================================================
# 2. 数据加载（带缓存，加快运行速度）
# ============================================================
@st.cache_data(ttl=600, show_spinner=False)
def load_rates(path: str):
    """读取原报价表 rate_table.xlsx"""
    p = file_path(path)
    if not os.path.exists(p):
        return None, None
    freight = pd.read_excel(p, sheet_name="海运费")
    surcharge = pd.read_excel(p, sheet_name="杂费")
    return freight, surcharge

@st.cache_data(ttl=600, show_spinner=False)
def load_sea(path: str):
    """读取东南亚运价表 southeast_asia_rates.xlsx"""
    p = file_path(path)
    if not os.path.exists(p):
        return None
    return pd.read_excel(p, sheet_name="运价总表")

# ============================================================
# 3. 备注翻译（中文 → 英文，给海外客户）
# ============================================================
def translate_remark(text: str) -> str:
    """把备注列里的中文规则翻译成英文（英文原样保留）"""
    if not isinstance(text, str) or not text.strip():
        return ""
    # 常见术语映射
    terms = {
        "目免": "Free time", "超重费": "overweight surcharge",
        "亏舱费": "cancellation/no-show fee", "亏舱": "cancellation fee",
        "小柜": "20GP", "高柜": "40HQ", "高箱": "40HQ",
        "限重": "weight limit", "不限重": "no weight limit",
        "轻货优先": "light cargo priority", "重量搭配": "weight balancing",
        "直达": "direct", "中转": "transshipment", "转": "transit",
        "排队": "tight/queuing", "紧张": "tight", "爆仓": "oversold",
        "未出": "TBA", "现仓": "space available", "现舱": "space available",
        "多柜可批": "multi-container negotiable", "大票可批": "large shipment negotiable",
        "可到付": "collect payable", "默认到付": "default collect",
        "可申请预付": "prepaid on request", "可以申请预付": "prepaid on request",
        "预付": "prepaid", "到付": "collect", "免电放费": "no telex-release fee",
        "保舱保箱不甩": "guaranteed space, no rollover", "保舱不甩": "guaranteed space",
        "保舱": "guaranteed space", "不甩": "no rollover",
        "截港前": "before C/O cutoff", "退关": "customs cancellation",
        "含箱重": "incl. container tare", "加班船": "extra-loader",
        "待更新": "to be updated", "等运价更新": "awaiting rate update",
        "目的港": "destination", "起运港": "POL",
        "燃油附加费": "bunker surcharge", "海运费已含": "included in freight",
        "纺织品与非纺织品同价": "same rate textile & non-textile",
        "大榭码头": "Daxie terminal", "三期码头": "Terminal 3", "码头": "terminal",
        "小重柜": "heavy 20GP", "重柜": "heavy container",
        "小重不接": "heavy 20GP not accepted", "停收": "not accepting",
        "不接": "not accepted", "敞开要货": "open for cargo",
        "若干清单": "some allocations", "可刷": "bookable",
        "捡漏": "last-minute space", "整船已刷爆": "vessel fully booked",
        "二程仓位需确认": "on-carriage space to confirm",
        "印尼": "Indonesia", "东马": "East Malaysia",
        "外传外": "outer transshipment", "乍浦": "Zhapu", "温州": "Wenzhou",
        "超高": "overheight", "限收": "limited acceptance",
        "超高": "overheight", "场内外": "on/off-terminal",
        "单边不少于": "min per side", "自由组合": "free combination",
        "混合使用免箱期": "mixed free time", "混用": "mixed free time",
        "买21": "buy 21 days", "可免费申请": "free on request",
        "预提": "pre-pull", "小仓": "limited space",
        "仓库": "warehouse", "靠泊": "berth", "还空": "empty return",
        "强制退运": "forced return", "越南盾结算": "settle in VND",
        "曼谷": "Bangkok", "胡志明": "HoChiMinh", "海防": "Haiphong",
        "金边": "Phnom Penh", "新加坡": "Singapore", "林查班": "Laem Chabang",
        "巴生": "Port Klang", "雅加达": "Jakarta", "泗水": "Surabaya",
        "三宝垄": "Semarang", "槟城": "Penang", "马尼拉": "Manila",
        "拉塔班": "Lat Krabang", "莱卡帮": "Lat Krabang", "西哈": "Sihanoukville",
    }
    # 先处理容易拆分的，再逐词替换
    out = text
    for zh, en in terms.items():
        out = out.replace(zh, en)
    return out.strip()

# ============================================================
# 4. 生成报价单 PDF
# ============================================================
def build_pdf_payload(
    pol, pod, ctype, qty, freight_price, surcharges, customer, lang
):
    """构造 PDF 数据（HTML 模板，用 pdfkit/weasyprint 或前端打印）"""
    T = LANG[lang]
    lines = []
    lines.append(f"<h2>{T['title']}</h2>")
    if customer:
        lines.append(f"<p><b>Customer:</b> {customer}</p>")
    lines.append(f"<p><b>{T['pol']}:</b> {pol} &nbsp;&nbsp; <b>{T['pod']}:</b> {pod} "
                 f"&nbsp;&nbsp; <b>{T['type']}:</b> {ctype} &nbsp;&nbsp; <b>{T['qty']}:</b> {qty}</p>")
    lines.append("<table border='1' cellspacing='0' cellpadding='6' style='border-collapse:collapse'>")
    lines.append(f"<tr><th>{T['freight']}</th><th>{T['surcharge']}</th><th>{T['total']}</th></tr>")
    lines.append(f"<tr><td>{freight_price:.2f}</td><td>{surcharges:.2f}</td><td><b>{freight_price + surcharges:.2f}</b></td></tr>")
    lines.append("</table>")
    lines.append("<p style='font-size:12px'>Quotation Date: {}</p>".format(datetime.now().strftime("%Y-%m-%d")))
    return "\n".join(lines)

# ============================================================
# 5. 主界面
# ============================================================
def main():
    lang = st.sidebar.selectbox("语言 / Language", list(LANG.keys()), index=1)
    T = LANG[lang]

    # 顶部标题 + Logo
    col_logo, col_title = st.columns([1, 6])
    logo = file_path("logo.png")
    if os.path.exists(logo):
        col_logo.image(logo, width=90)
    col_title.title(T["title"])

    # 导航
    page = st.sidebar.radio("Menu / 菜单", [T["nav_quote"], T["nav_sea"]])

    # ============ 报价单生成 ============
    if page == T["nav_quote"]:
        freight, surcharge = load_rates("rate_table.xlsx")
        if freight is None or freight.empty:
            st.warning("未找到 rate_table.xlsx，请确认文件已上传到仓库。")
            return

        st.subheader(T["result"])
        with st.form("quote_form"):
            c1, c2, c3, c4 = st.columns(4)
            pol = c1.selectbox(T["pol"], sorted(freight["起运港"].dropna().unique()))
            pod = c2.selectbox(T["pod"], sorted(freight["目的港"].dropna().unique()))
            ctype = c3.selectbox(T["type"], ["20GP", "40GP", "40HQ"])
            qty = c4.number_input(T["qty"], min_value=1, value=1, step=1)
            customer = st.text_input(T["customer"])
            submitted = st.form_submit_button(T["gen"], type="primary")

        if submitted:
            row = freight[(freight["起运港"] == pol) & (freight["目的港"] == pod)]
            if row.empty:
                st.warning(f"No rate for {pol} → {pod}")
            else:
                price = float(row.iloc[0].get(ctype, 0) or 0)
                freight_total = price * qty
                sur_total = 0.0
                if surcharge is not None and not surcharge.empty:
                    sur_total = float(surcharge.iloc[0].get("金额", 0) or 0) * qty
                st.success(f"{T['freight']}: ${freight_total:,.2f} | "
                           f"{T['surcharge']}: ${sur_total:,.2f} | "
                           f"{T['total']}: ${freight_total + sur_total:,.2f}")
                html = build_pdf_payload(pol, pod, ctype, qty, freight_total, sur_total, customer, lang)
                st.markdown(html, unsafe_allow_html=True)
                st.download_button(
                    "📄 下载报价单 PDF", html, file_name=f"quote_{pod}_{datetime.now():%Y%m%d}.html",
                    mime="text/html",
                )

    # ============ 东南亚运价查询 ============
    else:
        sea = load_sea("southeast_asia_rates.xlsx")
        if sea is None or sea.empty:
            st.warning("未找到 southeast_asia_rates.xlsx，请确认文件已上传到仓库。")
            return

        st.subheader("Southeast Asia Ocean Freight Rates")
        st.caption(T["sea_note"])

        # 筛选
        col_f1, col_f2, col_f3 = st.columns(3)
        with col_f1:
            pod_options = sorted(sea["目的港"].dropna().unique())
            pod_sel = st.multiselect("Destination / 目的港", pod_options)
        with col_f2:
            car_options = sorted(sea["船公司"].dropna().unique())
            car_sel = st.multiselect("Carrier / 船公司", car_options)
        with col_f3:
            kw = st.text_input(T["search"], placeholder=T["search_hint"])

        # 应用筛选
        df = sea.copy()
        if pod_sel:
            df = df[df["目的港"].isin(pod_sel)]
        if car_sel:
            df = df[df["船公司"].isin(car_sel)]
        if kw:
            mask = (
                df["目的港"].astype(str).str.contains(kw, case=False, na=False)
                | df["船公司"].astype(str).str.contains(kw, case=False, na=False)
            )
            df = df[mask]

        if df.empty:
            st.info(T["no_data"])
        else:
            # 英文显示：翻译备注
            view = pd.DataFrame({
                "Destination": df["目的港"],
                "Carrier": df["船公司"],
                "Vessel/Voyage": df["船名航次"].fillna(""),
                "Direct/Transit": df["直达/中转"].fillna(""),
                "Transit Time": df["时效"].fillna(""),
                "20GP (USD)": df["20GP(USD)"],
                "40GP (USD)": df["40GP(USD)"],
                "40HQ (USD)": df["40HQ(USD)"],
                "Sail Date": df["船期"].fillna(""),
                "Remark": df["备注"].apply(translate_remark),
            })

            st.dataframe(
                view,
                use_container_width=True,
                hide_index=True,
                height=560,
                column_config={
                    "20GP (USD)": st.column_config.NumberColumn(format="$%d"),
                    "40GP (USD)": st.column_config.NumberColumn(format="$%d"),
                    "40HQ (USD)": st.column_config.NumberColumn(format="$%d"),
                },
            )

            st.caption(f"Total records: {len(view)}")
            csv = view.to_csv(index=False).encode("utf-8-sig")
            st.download_button(T["download_csv"], csv,
                               file_name="se_asia_rates.csv", mime="text/csv")

if __name__ == "__main__":
    main()

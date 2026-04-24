import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px

st.set_page_config(page_title="周度平台销售分析", layout="wide")

st.title("📊 周度平台销售分析工具")
st.caption("上传 Artemis 销售表、PB 销售表、ASIN 三级分类匹配表，自动聚合并生成周度分析")

# =========================
# 工具函数
# =========================
def normalize_columns(df):
    df = df.copy()
    df.columns = df.columns.astype(str).str.strip()
    return df


def read_file(uploaded_file):
    if uploaded_file is None:
        return None
    file_name = uploaded_file.name.lower()
    try:
        if file_name.endswith(".csv"):
            return pd.read_csv(uploaded_file)
        elif file_name.endswith(".xlsx") or file_name.endswith(".xls"):
            return pd.read_excel(uploaded_file)
        else:
            st.error(f"不支持的文件格式：{uploaded_file.name}")
            st.stop()
    except Exception as e:
        st.error(f"读取文件失败：{uploaded_file.name}，错误信息：{e}")
        st.stop()


def find_column(df, candidates, field_name, platform_name):
    for c in candidates:
        if c in df.columns:
            return c
    st.error(f"{platform_name} 缺少必要字段：{field_name}")
    st.write("当前检测到的列名：", list(df.columns))
    st.stop()


def clean_sales_value(series):
    return pd.to_numeric(
        series.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("$", "", regex=False)
        .str.replace("¥", "", regex=False)
        .str.strip(),
        errors="coerce"
    ).fillna(0)


def build_sales_df(df_raw, platform_name):
    """
    标准化平台销售表字段为：
    Date / ASIN / Sales / Conversions / Platform
    """
    df = normalize_columns(df_raw)

    date_col = find_column(df, ["Date", "日期", "date"], "Date/日期", platform_name)
    asin_col = find_column(df, ["ASIN", "asin"], "ASIN", platform_name)
    sales_col = find_column(df, ["Sales", "销额", "sales", "Ordered Product Sales"], "Sales/销额", platform_name)
    conv_col = find_column(df, ["Conversions", "销量", "conversion", "Ordered Units", "Units Ordered"], "Conversions/销量", platform_name)

    df = df[[date_col, asin_col, sales_col, conv_col]].copy()
    df.columns = ["Date", "ASIN", "Sales", "Conversions"]

    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df["ASIN"] = df["ASIN"].astype(str).str.strip()
    df["Sales"] = clean_sales_value(df["Sales"])
    df["Conversions"] = pd.to_numeric(df["Conversions"], errors="coerce").fillna(0)

    df = df.dropna(subset=["Date"])
    df = df[
        (df["ASIN"] != "") &
        (df["ASIN"].str.lower() != "nan") &
        (df["ASIN"] != "无")
    ].copy()

    df["Platform"] = platform_name
    return df


def build_mapping_df(df_raw):
    """
    匹配表要求字段：
    ASIN / 三级分类
    """
    df = normalize_columns(df_raw)

    asin_col = find_column(df, ["ASIN", "asin"], "ASIN", "匹配表")
    cat3_col = find_column(df, ["三级分类"], "三级分类", "匹配表")

    df = df[[asin_col, cat3_col]].copy()
    df.columns = ["ASIN", "三级分类"]

    df["ASIN"] = df["ASIN"].astype(str).str.strip()
    df["三级分类"] = df["三级分类"].astype(str).str.strip()

    df = df[
        (df["ASIN"] != "") &
        (df["ASIN"].str.lower() != "nan") &
        (df["ASIN"] != "无")
    ].copy()

    df["三级分类"] = df["三级分类"].replace("", np.nan).fillna("未匹配")
    df = df.drop_duplicates(subset=["ASIN"], keep="first")
    return df


def safe_divide(a, b):
    return np.where(b == 0, 0, a / b)


def format_week_label(date_series):
    """
    生成类似：
    2026.4.13-2026.4.19
    """
    week_start = date_series.dt.to_period("W-SUN").apply(lambda r: r.start_time)
    week_end = date_series.dt.to_period("W-SUN").apply(lambda r: r.end_time)

    label = (
        week_start.dt.year.astype(str) + "." +
        week_start.dt.month.astype(str) + "." +
        week_start.dt.day.astype(str) + "-" +
        week_end.dt.year.astype(str) + "." +
        week_end.dt.month.astype(str) + "." +
        week_end.dt.day.astype(str)
    )
    return week_start, week_end, label


def add_week_fields(df):
    df = df.copy()
    iso = df["Date"].dt.isocalendar()
    df["WeekNum"] = iso.week.astype(int)
    df["Year"] = iso.year.astype(int)

    week_start, week_end, week_range = format_week_label(df["Date"])
    df["WeekStart"] = week_start
    df["WeekEnd"] = week_end
    df["WeekRange"] = week_range
    df["WeekLabel"] = df["WeekRange"] + " W" + df["WeekNum"].astype(str)
    return df


def to_excel_bytes(detail_df, weekly_df, platform_weekly_df, category_weekly_df):
    from io import BytesIO
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        detail_df.to_excel(writer, index=False, sheet_name="聚合明细")
        weekly_df.to_excel(writer, index=False, sheet_name="周度汇总")
        platform_weekly_df.to_excel(writer, index=False, sheet_name="平台周度")
        category_weekly_df.to_excel(writer, index=False, sheet_name="三级分类周度")
    return output.getvalue()


# =========================
# 上传区
# =========================
st.subheader("一、上传数据")

col1, col2, col3 = st.columns(3)

with col1:
    artemis_file = st.file_uploader("上传 Artemis 销售表", type=["xlsx", "xls", "csv"])

with col2:
    pb_file = st.file_uploader("上传 PB 销售表", type=["xlsx", "xls", "csv"])

with col3:
    mapping_file = st.file_uploader("上传 ASIN-三级分类匹配表", type=["xlsx", "xls", "csv"])

if artemis_file and pb_file and mapping_file:
    # =========================
    # 读取数据
    # =========================
    artemis_raw = read_file(artemis_file)
    pb_raw = read_file(pb_file)
    mapping_raw = read_file(mapping_file)

    artemis_df = build_sales_df(artemis_raw, "Artemis")
    pb_df = build_sales_df(pb_raw, "PB")
    mapping_df = build_mapping_df(mapping_raw)

    # 合并平台数据
    sales_df = pd.concat([artemis_df, pb_df], ignore_index=True)

    # 按 ASIN 匹配三级分类
    merged_df = sales_df.merge(mapping_df, on="ASIN", how="left")
    merged_df["三级分类"] = merged_df["三级分类"].replace("", np.nan).fillna("未匹配")

    # =========================
    # 按方案B聚合
    # Date + Platform + 三级分类 + ASIN
    # =========================
    detail_df = (
        merged_df.groupby(["Date", "Platform", "三级分类", "ASIN"], as_index=False)
        .agg(
            Sales=("Sales", "sum"),
            Conversions=("Conversions", "sum")
        )
    )

    detail_df["客单价"] = safe_divide(detail_df["Sales"], detail_df["Conversions"])
    detail_df = add_week_fields(detail_df)

    st.success("数据处理完成：已完成平台合并、三级分类匹配、按维度聚合、客单价计算。")

    # =========================
    # 筛选区
    # =========================
    st.subheader("二、筛选条件")

    f1, f2, f3 = st.columns(3)

    with f1:
        platform_options = ["全部"] + sorted(detail_df["Platform"].dropna().unique().tolist())
        selected_platform = st.selectbox("选择平台", platform_options)

    with f2:
        category_options = ["全部"] + sorted(detail_df["三级分类"].dropna().unique().tolist())
        selected_category = st.selectbox("选择三级分类", category_options)

    with f3:
        week_options = ["全部"] + (
            detail_df[["WeekStart", "WeekLabel"]]
            .drop_duplicates()
            .sort_values("WeekStart")["WeekLabel"]
            .tolist()
        )
        selected_week = st.selectbox("选择周", week_options)

    filtered_df = detail_df.copy()

    if selected_platform != "全部":
        filtered_df = filtered_df[filtered_df["Platform"] == selected_platform]

    if selected_category != "全部":
        filtered_df = filtered_df[filtered_df["三级分类"] == selected_category]

    if selected_week != "全部":
        filtered_df = filtered_df[filtered_df["WeekLabel"] == selected_week]

    # =========================
    # 总览指标
    # =========================
    st.subheader("三、总览指标")

    total_sales = filtered_df["Sales"].sum()
    total_conv = filtered_df["Conversions"].sum()
    avg_price = 0 if total_conv == 0 else total_sales / total_conv
    asin_count = filtered_df["ASIN"].nunique()
    category_count = filtered_df["三级分类"].nunique()

    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("总销额", f"{total_sales:,.2f}")
    m2.metric("总销量", f"{total_conv:,.0f}")
    m3.metric("平均客单价", f"{avg_price:,.2f}")
    m4.metric("ASIN数", f"{asin_count:,}")
    m5.metric("三级分类数", f"{category_count:,}")

    # =========================
    # 周度汇总
    # =========================
    st.subheader("四、周度汇总分析")

    weekly_df = (
        filtered_df.groupby(["Year", "WeekNum", "WeekStart", "WeekEnd", "WeekRange", "WeekLabel"], as_index=False)
        .agg(
            Sales=("Sales", "sum"),
            Conversions=("Conversions", "sum")
        )
        .sort_values("WeekStart")
    )
    weekly_df["客单价"] = safe_divide(weekly_df["Sales"], weekly_df["Conversions"])

    st.dataframe(
        weekly_df[["WeekLabel", "Sales", "Conversions", "客单价"]],
        use_container_width=True
    )

    c1, c2 = st.columns(2)

    with c1:
        fig_week_sales = px.line(
            weekly_df,
            x="WeekLabel",
            y="Sales",
            markers=True,
            title="周度销额趋势"
        )
        st.plotly_chart(fig_week_sales, use_container_width=True)

    with c2:
        fig_week_conv = px.line(
            weekly_df,
            x="WeekLabel",
            y="Conversions",
            markers=True,
            title="周度销量趋势"
        )
        st.plotly_chart(fig_week_conv, use_container_width=True)

    fig_week_price = px.line(
        weekly_df,
        x="WeekLabel",
        y="客单价",
        markers=True,
        title="周度客单价趋势"
    )
    st.plotly_chart(fig_week_price, use_container_width=True)

    # =========================
    # 平台周度分析
    # =========================
    st.subheader("五、平台周度分析")

    platform_weekly_df = (
        filtered_df.groupby(["WeekStart", "WeekLabel", "Platform"], as_index=False)
        .agg(
            Sales=("Sales", "sum"),
            Conversions=("Conversions", "sum")
        )
        .sort_values(["WeekStart", "Platform"])
    )
    platform_weekly_df["客单价"] = safe_divide(platform_weekly_df["Sales"], platform_weekly_df["Conversions"])

    st.dataframe(platform_weekly_df, use_container_width=True)

    p1, p2 = st.columns(2)

    with p1:
        fig_platform_sales = px.bar(
            platform_weekly_df,
            x="WeekLabel",
            y="Sales",
            color="Platform",
            barmode="group",
            title="平台周度销额对比"
        )
        st.plotly_chart(fig_platform_sales, use_container_width=True)

    with p2:
        fig_platform_conv = px.bar(
            platform_weekly_df,
            x="WeekLabel",
            y="Conversions",
            color="Platform",
            barmode="group",
            title="平台周度销量对比"
        )
        st.plotly_chart(fig_platform_conv, use_container_width=True)

    fig_platform_price = px.line(
        platform_weekly_df,
        x="WeekLabel",
        y="客单价",
        color="Platform",
        markers=True,
        title="平台周度客单价对比"
    )
    st.plotly_chart(fig_platform_price, use_container_width=True)

    # =========================
    # 三级分类分析
    # =========================
    st.subheader("六、三级分类分析")

    category_summary_df = (
        filtered_df.groupby("三级分类", as_index=False)
        .agg(
            Sales=("Sales", "sum"),
            Conversions=("Conversions", "sum"),
            ASIN数=("ASIN", "nunique")
        )
        .sort_values("Sales", ascending=False)
    )
    category_summary_df["客单价"] = safe_divide(category_summary_df["Sales"], category_summary_df["Conversions"])

    st.dataframe(category_summary_df, use_container_width=True)

    top_n = st.slider("展示 Top N 三级分类", min_value=5, max_value=30, value=10)

    cat1, cat2 = st.columns(2)

    with cat1:
        fig_cat_sales = px.bar(
            category_summary_df.head(top_n),
            x="三级分类",
            y="Sales",
            title=f"Top {top_n} 三级分类销额"
        )
        st.plotly_chart(fig_cat_sales, use_container_width=True)

    with cat2:
        fig_cat_price = px.bar(
            category_summary_df.head(top_n),
            x="三级分类",
            y="客单价",
            title=f"Top {top_n} 三级分类客单价"
        )
        st.plotly_chart(fig_cat_price, use_container_width=True)

    # =========================
    # 三级分类周度分析
    # =========================
    st.subheader("七、三级分类周度分析")

    category_weekly_df = (
        filtered_df.groupby(["WeekStart", "WeekLabel", "三级分类"], as_index=False)
        .agg(
            Sales=("Sales", "sum"),
            Conversions=("Conversions", "sum")
        )
        .sort_values(["WeekStart", "Sales"], ascending=[True, False])
    )
    category_weekly_df["客单价"] = safe_divide(category_weekly_df["Sales"], category_weekly_df["Conversions"])

    st.dataframe(category_weekly_df, use_container_width=True)

    top_categories = (
        category_summary_df.head(min(top_n, len(category_summary_df)))["三级分类"].tolist()
    )
    category_weekly_top_df = category_weekly_df[category_weekly_df["三级分类"].isin(top_categories)]

    fig_cat_week_sales = px.line(
        category_weekly_top_df,
        x="WeekLabel",
        y="Sales",
        color="三级分类",
        markers=True,
        title="重点三级分类周度销额趋势"
    )
    st.plotly_chart(fig_cat_week_sales, use_container_width=True)

    # =========================
    # 聚合明细表
    # =========================
    st.subheader("八、聚合明细数据")

    show_cols = [
        "Date", "WeekLabel", "Platform", "三级分类", "ASIN",
        "Sales", "Conversions", "客单价"
    ]
    st.dataframe(
        filtered_df[show_cols].sort_values(["Date", "Platform", "三级分类", "ASIN"]),
        use_container_width=True,
        height=500
    )

    # =========================
    # 导出
    # =========================
    st.subheader("九、导出结果")

    excel_data = to_excel_bytes(
        detail_df=filtered_df,
        weekly_df=weekly_df,
        platform_weekly_df=platform_weekly_df,
        category_weekly_df=category_weekly_df
    )

    st.download_button(
        label="📥 下载分析结果 Excel",
        data=excel_data,
        file_name="周度平台销售分析结果.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("请先上传 Artemis 销售表、PB 销售表 和 ASIN-三级分类匹配表。")

# -*- coding: utf-8 -*-

import pandas as pd
import streamlit as st
import plotly.express as px

# =========================================================
# إعداد الصفحة
# =========================================================
st.set_page_config(
    page_title="لوحة تحليل الاجتماعات",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# =========================================================
# CSS — RTL + شكل احترافي بسيط
# =========================================================
st.markdown(
    """
    <style>
    html, body, [class*="css"] {
        direction: rtl !important;
        text-align: right !important;
        font-family: "Tahoma", "Segoe UI", Arial, sans-serif;
    }

    .stApp {
        direction: rtl !important;
    }

    .block-container {
        max-width: 1500px;
        padding-top: 1.2rem;
        padding-bottom: 1.5rem;
    }

    h1, h2, h3, h4, p, label {
        direction: rtl !important;
        text-align: right !important;
    }

    /* العنوان */
    .dashboard-title {
        font-size: 32px;
        font-weight: 800;
        color: #173f73;
        margin-bottom: 24px;
    }

    /* بطاقات الأرقام */
    [data-testid="stMetric"] {
        direction: rtl !important;
        text-align: right !important;
        background: #ffffff;
        border: 1px solid #e1e7ef;
        border-radius: 12px;
        padding: 18px 20px;
        min-height: 145px;
        box-shadow: 0 2px 8px rgba(20, 50, 90, 0.04);
    }

    [data-testid="stMetricLabel"] {
        direction: rtl !important;
        text-align: right !important;
        font-size: 15px !important;
        color: #31567e !important;
        margin-bottom: 8px !important;
    }

    [data-testid="stMetricValue"] {
        direction: ltr !important;
        text-align: right !important;
        font-size: 34px !important;
        color: #123f73 !important;
        white-space: normal !important;
        overflow-wrap: anywhere !important;
    }

    [data-testid="stMetricDelta"] {
        direction: ltr !important;
        text-align: right !important;
        font-size: 14px !important;
    }

    /* زيادة مساحة كارت أكثر مندوب نشاطاً */
    div[data-testid="column"]:nth-child(3) [data-testid="stMetric"] {
        min-height: 175px;
    }

    div[data-testid="column"]:nth-child(3) [data-testid="stMetricValue"] {
        font-size: 27px !important;
        line-height: 1.25 !important;
        white-space: normal !important;
        word-break: break-word !important;
    }

    /* صناديق الرسوم */
    div[data-testid="stPlotlyChart"] {
        background: #ffffff;
        border: 1px solid #e1e7ef;
        border-radius: 12px;
        padding: 8px;
        box-shadow: 0 2px 8px rgba(20, 50, 90, 0.035);
    }

    /* الجدول */
    div[data-testid="stDataFrame"] {
        border: 1px solid #e1e7ef;
        border-radius: 12px;
        overflow: hidden;
    }

    hr {
        margin-top: 20px !important;
        margin-bottom: 24px !important;
        border-color: #e3e8ef !important;
    }

    /* إخفاء العناصر غير الضرورية */
    #MainMenu {
        visibility: hidden;
    }

    footer {
        visibility: hidden;
    }

    </style>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# تحميل البيانات
# =========================================================
@st.cache_data
def load_data(path):

    df = pd.read_excel(path)

    # حذف صف العناوين المكرر
    if "Source" in df.columns:
        df = df[
            df["Source"].astype(str).str.strip() != "Source"
        ].copy()

    # التواريخ
    df["Date"] = pd.to_datetime(
        df["Date"],
        errors="coerce"
    )

    df["Creation Date"] = pd.to_datetime(
        df["Creation Date"],
        errors="coerce"
    )

    df = df.dropna(
        subset=["Date"]
    ).copy()

    # =====================================================
    # المصدر — الاحتفاظ بالقيمة الإنجليزية الأصلية
    # =====================================================
    df["Source"] = (
        df["Source"]
        .astype(str)
        .str.strip()
    )

    df.loc[
        df["Source"].isin(["nan", "None", ""]),
        "Source"
    ] = "Unknown"

    # اسم عربي للعرض فقط عند الحاجة، بدون تغيير القيمة الأصلية
    df["المصدر"] = df["Source"]

    # =====================================================
    # مندوب المبيعات
    # =====================================================
    df["Sales Rep"] = (
        df["Sales Rep"]
        .astype(str)
        .str.strip()
    )

    df.loc[
        df["Sales Rep"].isin(["nan", "None", ""]),
        "Sales Rep"
    ] = "Unknown"

    # =====================================================
    # الوقت
    # =====================================================
    df["الشهر"] = (
        df["Date"]
        .dt.to_period("M")
        .astype(str)
    )

    df["الأسبوع"] = (
        df["Date"]
        .dt.to_period("W")
        .apply(lambda x: x.start_time.date())
    )

    day_names_ar = {
        0: "الإثنين",
        1: "الثلاثاء",
        2: "الأربعاء",
        3: "الخميس",
        4: "الجمعة",
        5: "السبت",
        6: "الأحد",
    }

    df["اليوم"] = (
        df["Date"]
        .dt.dayofweek
        .map(day_names_ar)
    )

    return df


# =========================================================
# ملف البيانات
# =========================================================
DATA_PATH = "meetings_data.xlsx"

try:
    df = load_data(DATA_PATH)

except FileNotFoundError:

    uploaded = st.file_uploader(
        "اختر ملف Excel",
        type=["xlsx"]
    )

    if uploaded is None:
        st.stop()

    df = load_data(uploaded)


# =========================================================
# عنوان لوحة التحكم
# =========================================================
st.markdown(
    """
    <div style="
        width: 100%;
        direction: rtl;
        text-align: right;
        font-size: 32px;
        font-weight: 800;
        color: #173f73;
        margin-bottom: 24px;
        padding-right: 5px;
    ">
        📊 لوحة تحليل الاجتماعات
    </div>
    """,
    unsafe_allow_html=True
)


# =========================================================
# الأرقام المهمة
# =========================================================
total_meetings = len(df)

by_month = df.groupby("الشهر").size()

aug = int(
    by_month.get("2026-08", 0)
)

sep = int(
    by_month.get("2026-09", 0)
)

change = (
    ((sep - aug) / aug * 100)
    if aug > 0
    else 0
)

rep_counts = (
    df["Sales Rep"]
    .value_counts()
)

top_rep = (
    rep_counts.index[0]
    if not rep_counts.empty
    else "—"
)

top_rep_count = (
    int(rep_counts.iloc[0])
    if not rep_counts.empty
    else 0
)

source_counts = (
    df["Source"]
    .value_counts()
)

top_source = (
    source_counts.index[0]
    if not source_counts.empty
    else "—"
)

top_source_pct = (
    source_counts.iloc[0] / total_meetings * 100
    if total_meetings > 0 and not source_counts.empty
    else 0
)

# ترتيب البطاقات بصرياً من اليمين لليسار
# =========================================================
# بطاقات الأرقام المهمة
# =========================================================

c1, c2, c3, c4 = st.columns(
    [1, 2.2, 1, 1],
    gap="medium"
)

with c1:
    st.metric(
        "إجمالي الاجتماعات",
        total_meetings
    )

with c2:
    st.metric(
        "أكثر مندوب نشاطًا",
        top_rep,
        f"{top_rep_count} اجتماع"
    )

with c3:
    st.metric(
        "أغسطس / سبتمبر",
        f"{aug} / {sep}",
        f"{change:+.0f}%"
    )

with c4:
    st.metric(
        "أكبر مصدر",
        top_source,
        f"{top_source_pct:.0f}%"
    )

# =========================================================
# تجهيز الرسوم
# =========================================================

# الاجتماعات أسبوعياً
weekly = (
    df.groupby("الأسبوع")
    .size()
    .reset_index(name="العدد")
)

fig_trend = px.bar(
    weekly,
    x="الأسبوع",
    y="العدد",
    text="العدد",
)

fig_trend.update_traces(
    textposition="outside",
    cliponaxis=False,
)

fig_trend.update_layout(
    title={
        "text": "عدد الاجتماعات أسبوعيًا",
        "x": 0.98,
        "xanchor": "right",
    },
    xaxis_title="الأسبوع",
    yaxis_title="عدد الاجتماعات",
    margin=dict(l=50, r=25, t=60, b=45),
    height=350,
    showlegend=False,
)


# =========================================================
# الاجتماعات حسب اليوم
# =========================================================
day_order = [
    "الإثنين",
    "الثلاثاء",
    "الأربعاء",
    "الخميس",
    "الجمعة",
    "السبت",
    "الأحد",
]

by_day = (
    df.groupby("اليوم")
    .size()
    .reindex(day_order)
    .fillna(0)
    .reset_index(name="العدد")
)

fig_day = px.bar(
    by_day,
    x="اليوم",
    y="العدد",
    text="العدد",
)

fig_day.update_traces(
    textposition="outside",
    cliponaxis=False,
)

fig_day.update_layout(
    title={
        "text": "الاجتماعات حسب يوم الأسبوع",
        "x": 0.98,
        "xanchor": "right",
    },
    xaxis_title="اليوم",
    yaxis_title="عدد الاجتماعات",
    margin=dict(l=50, r=25, t=60, b=45),
    height=350,
    showlegend=False,
)


# =========================================================
# الاجتماعات لكل مندوب
# =========================================================
rep_chart = (
    df["Sales Rep"]
    .value_counts()
    .head(15)
    .sort_values()
    .reset_index()
)

rep_chart.columns = [
    "المندوب",
    "العدد"
]

fig_reps = px.bar(
    rep_chart,
    x="العدد",
    y="المندوب",
    orientation="h",
    text="العدد",
)

fig_reps.update_traces(
    textposition="outside",
    cliponaxis=False,
)

fig_reps.update_layout(
    title={
        "text": "عدد اجتماعات كل مندوب",
        "x": 0.98,
        "xanchor": "right",
    },
    xaxis_title="عدد الاجتماعات",
    yaxis_title="",
    margin=dict(l=55, r=25, t=60, b=45),
    height=390,
    showlegend=False,
)


# =========================================================
# مصادر العملاء
# =========================================================
src_counts = (
    df["Source"]
    .value_counts()
    .reset_index()
)

src_counts.columns = [
    "المصدر",
    "العدد"
]


# =========================================================
# الصف الأول — أسبوعياً + حسب اليوم
# =========================================================
left, right = st.columns(
    2,
    gap="medium"
)

# RTL: اليمين يظهر أولاً للمستخدم
with right:
    st.plotly_chart(
        fig_day,
        use_container_width=True,
        config={
            "displayModeBar": False
        }
    )

with left:
    st.plotly_chart(
        fig_trend,
        use_container_width=True,
        config={
            "displayModeBar": False
        }
    )


# =========================================================
# الصف الثاني — المندوبون + المصادر
# =========================================================
left2, right2 = st.columns(
    [1.25, 1],
    gap="medium"
)

with right2:
    st.plotly_chart(
        fig_reps,
        use_container_width=True,
        config={
            "displayModeBar": False
        }
    )

with left2:

    st.markdown(
        '<div style="font-size:22px;font-weight:700;color:#173f73;'
        'margin-bottom:12px;text-align:right;">أكثر مصادر العملاء</div>',
        unsafe_allow_html=True
    )

    st.dataframe(
        src_counts,
        use_container_width=True,
        hide_index=True,
        height=390,
        column_config={
            "المصدر": st.column_config.TextColumn(
                "المصدر",
                width="large"
            ),
            "العدد": st.column_config.NumberColumn(
                "العدد",
                width="small"
            ),
        }
    )

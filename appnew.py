
# ==================================================
# STACKED BAR CHART
# 1 EXCEL FILE = 1 BAR
# ==================================================

st.divider()
st.header("📊 CRS Gain - t / f Comparison")

# เตรียมข้อมูล
chart_data = []

for r in all_results:

    chart_data.append({
        "Excel File": r["filename"],
        "t": r["t"],
        "f": r["f"]
    })

chart_df = pd.DataFrame(chart_data)

# แปลงข้อมูลเป็น Long Format
long_df = chart_df.melt(
    id_vars=["Excel File"],
    value_vars=["t", "f"],
    var_name="Result",
    value_name="Count"
)

# กำหนดลำดับไฟล์
file_order = chart_df["Excel File"].tolist()

# กราฟแท่งซ้อน: 1 ไฟล์ = 1 แท่ง
chart_spec = {
    "width": "container",
    "height": 550,

    "mark": {
        "type": "bar",
        "tooltip": True
    },

    "encoding": {

        "x": {
            "field": "Excel File",
            "type": "nominal",
            "sort": file_order,
            "scale": {
                "paddingInner": 0.02,
                "paddingOuter": 0.01
            },
            "axis": {
                "title": "Excel File",
                "labelAngle": -60,
                "labelFontSize": 10,
                "labelLimit": 180
            }
        },

        "y": {
            "aggregate": "sum",
            "field": "Count",
            "type": "quantitative",
            "stack": "zero",
            "axis": {
                "title": "Total Records"
            }
        },

        "color": {
            "field": "Result",
            "type": "nominal",
            "scale": {
                "domain": ["t", "f"],
                "range": ["#16B98A", "#F5C542"]
            },
            "legend": {
                "title": "Result",
                "orient": "right"
            }
        },

        "order": {
            "field": "Result",
            "sort": ["t", "f"]
        }
    },

    "config": {
        "view": {
            "stroke": None
        }
    }
}

st.vega_lite_chart(
    long_df,
    chart_spec,
    use_container_width=True
)

# ตารางข้อมูลประกอบ
st.subheader("📋 รายละเอียดจำนวน t / f")

st.dataframe(
    chart_df,
    use_container_width=True,
    hide_index=True
)

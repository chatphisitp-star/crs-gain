
# ==================================================
# 6. COMPARISON CHART - STACKED COLUMN
# ==================================================

st.divider()
st.header("📊 กราฟเปรียบเทียบ CRS Gain ทุกไฟล์")

comparison_data = []

for r in all_results:

    comparison_data.append({
        "Excel File": r["filename"],
        "t": r["t"],
        "f": r["f"]
    })

comparison_df = pd.DataFrame(comparison_data)

# แปลงข้อมูลเป็นรูปแบบสำหรับกราฟ
chart_data = []

for _, row in comparison_df.iterrows():

    chart_data.append({
        "Excel File": row["Excel File"],
        "Result": "t - Correct",
        "Count": row["t"]
    })

    chart_data.append({
        "Excel File": row["Excel File"],
        "Result": "f - Incorrect",
        "Count": row["f"]
    })

chart_df = pd.DataFrame(chart_data)

# กำหนดกราฟแบบแท่งซ้อน
chart_spec = {
    "width": "container",
    "height": 500,

    "mark": {
        "type": "bar",
        "tooltip": True,
        "width": {"band": 0.97}
    },

    "encoding": {

        "x": {
            "field": "Excel File",
            "type": "nominal",
            "sort": [r["filename"] for r in all_results],
            "scale": {
                "paddingInner": 0.03,
                "paddingOuter": 0.01
            },
            "axis": {
                "title": "Excel File",
                "labelAngle": -45,
                "labelFontSize": 10,
                "labelLimit": 150
            }
        },

        "y": {
            "field": "Count",
            "type": "quantitative",
            "stack": "zero",
            "axis": {
                "title": "Number of Records"
            }
        },

        "color": {
            "field": "Result",
            "type": "nominal",
            "scale": {
                "domain": [
                    "t - Correct",
                    "f - Incorrect"
                ],
                "range": [
                    "#00B894",
                    "#F4C542"
                ]
            },
            "legend": {
                "title": "Validation Result",
                "orient": "right"
            }
        },

        "order": {
            "field": "Result",
            "sort": [
                "t - Correct",
                "f - Incorrect"
            ]
        }
    }
}

st.vega_lite_chart(
    chart_df,
    chart_spec,
    use_container_width=True
)

# ตารางสรุป
st.subheader("📋 ตารางสรุปผลแต่ละไฟล์")

st.dataframe(
    comparison_df,
    use_container_width=True,
    hide_index=True
)

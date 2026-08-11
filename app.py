from query_engine import (
    # ... your existing imports ...,
    generate_docx_report,
    generate_pdf_report,
)
import uuid
import plotly.express as px
import streamlit as st
from query_engine import run_predictive_regression
import streamlit as st
import plotly.express as px
import streamlit as st
from query_engine import client
import pandas as pd
import plotly.express as px
from profiler import load_data, generate_schema_catalog, detect_iqr_outliers
from query_engine import (
    nl_query_pipeline,
    format_result_with_llm,
    analyze_anomalies_iqr,
    ConversationMemory,
    generate_preset_insight,
)

# Page setup
st.set_page_config(
    page_title="AI Data Analytics Platform", layout="wide", page_icon="📊"
)

st.title("📊 AI-Powered Data Analytics Platform")
st.caption("Capstone Project | Headless LLM + Pandas Query Engine")

# Session State Initialization
if "memory" not in st.session_state:
    st.session_state.memory = ConversationMemory()
if "chat_history" not in st.session_state:
    st.session_state.chat_history = []

# --- 1. DATA INGESTION & SIDEBAR CONTROLS ---
with st.sidebar:
    st.header("1. Data Ingestion")
    uploaded_file = st.file_uploader("Upload CSV Dataset", type=["csv"])

    st.divider()
    st.header("Controls")
    if st.button("🧹 Clear Chat History"):
        st.session_state.memory.reset()
        st.session_state.chat_history = []
        st.rerun()

# Default or Uploaded Data Loading
if uploaded_file is not None:
    df = load_data(uploaded_file)
else:
    st.info("💡 Upload a CSV file in the sidebar to start analyzing your own data.")
    st.stop()

# --- 2. DYNAMIC FILTERS (Applied AFTER df is loaded) ---
st.sidebar.header("2. Dynamic Filters")
categorical_cols = df.select_dtypes(include=["object", "category"]).columns.tolist()

filtered_df = df.copy()
if categorical_cols:
    selected_col = st.sidebar.selectbox(
        "Filter by Dimension", ["None"] + categorical_cols
    )
    if selected_col != "None":
        unique_vals = df[selected_col].dropna().unique().tolist()
        selected_vals = st.sidebar.multiselect(
            f"Select {selected_col}", unique_vals, default=unique_vals
        )
        if selected_vals:
            filtered_df = df[df[selected_col].isin(selected_vals)]

# Build Schema Catalog based on active data
schema_str = generate_schema_catalog(filtered_df)


# --- 3. SMART MULTI-CHART GENERATOR ---
def render_smart_charts(res_data: pd.DataFrame, key_prefix: str = None):
    if not isinstance(res_data, pd.DataFrame) or res_data.empty:
        return

    # 1. Convert all column headers to strings to avoid c.lower() exceptions
    res_data.columns = [str(c) for c in res_data.columns]

    # 2. Suppress charts for single-row KPI overviews or 1-column tables
    if len(res_data) <= 1 or len(res_data.columns) <= 1:
        return

    if key_prefix is None:
        key_prefix = str(uuid.uuid4())[:8]

    # Exclude non-metric/key columns
    excluded_keywords = [
        "id",
        "key",
        "code",
        "index",
        "level_0",
        "number",
        "customer_key",
    ]

    num_cols = [
        c
        for c in res_data.select_dtypes(include=["number"]).columns
        if not any(kw in c.lower() for kw in excluded_keywords)
    ]

    if not num_cols:
        num_cols = res_data.select_dtypes(include=["number"]).columns.tolist()
        if not num_cols:
            return

    cat_cols = [
        c
        for c in res_data.columns
        if not any(kw in str(c).lower() for kw in excluded_keywords)
        and c not in num_cols
    ]
    if not cat_cols:
        cat_cols = res_data.columns.tolist()

    st.markdown("### 📊 Interactive Data Visualization")

    col_type, col_x, col_y = st.columns([1, 1.5, 1.5])

    with col_type:
        chart_type = st.selectbox(
            "Chart Type",
            [
                "Bar Chart",
                "Line Chart",
                "Pie / Donut",
                "Scatter Plot",
                "Histogram",
                "Box Plot",
                "Heatmap",
            ],
            key=f"chart_type_{key_prefix}",
        )

    with col_x:
        x_axis = st.selectbox(
            "X-Axis (Dimension)", cat_cols, index=0, key=f"x_axis_{key_prefix}"
        )

    with col_y:
        y_axis = st.selectbox(
            "Y-Axis (Metric)", num_cols, index=0, key=f"y_axis_{key_prefix}"
        )

    plot_df = res_data.copy()
    if plot_df[x_axis].dtype == "object":
        plot_df[x_axis] = plot_df[x_axis].astype(str)

    try:
        if chart_type == "Bar Chart":
            fig = px.bar(
                plot_df,
                x=x_axis,
                y=y_axis,
                color=x_axis if len(plot_df) <= 10 else None,
                title=f"{y_axis} by {x_axis}",
                text_auto=".2s",
            )
            fig.update_layout(showlegend=False)

        elif chart_type == "Line Chart":
            fig = px.line(
                plot_df,
                x=x_axis,
                y=y_axis,
                markers=True,
                title=f"{y_axis} Trend over {x_axis}",
            )

        elif chart_type == "Pie / Donut":
            fig = px.pie(
                plot_df,
                names=x_axis,
                values=y_axis,
                hole=0.4,
                title=f"{y_axis} Share by {x_axis}",
            )

        elif chart_type == "Scatter Plot":
            color_col = num_cols[1] if len(num_cols) > 1 else None
            fig = px.scatter(
                plot_df,
                x=x_axis,
                y=y_axis,
                color=color_col,
                size=y_axis,
                title=f"{y_axis} vs {x_axis}",
            )

        elif chart_type == "Histogram":
            fig = px.histogram(
                plot_df,
                x=y_axis,
                title=f"Distribution of {y_axis}",
            )

            fig.update_traces(
                marker_line_color="white",
                marker_line_width=1,
                marker_color=None,  # Let Plotly assign colors
            )
            fig.update_traces(marker_line_color="white", marker_line_width=1)

        elif chart_type == "Box Plot":
            fig = px.box(
                plot_df,
                x=x_axis,
                y=y_axis,
                color=x_axis if len(plot_df[x_axis].unique()) <= 10 else None,
                title=f"{y_axis} Distribution by {x_axis}",
            )

        elif chart_type == "Heatmap":
            if len(num_cols) < 2:
                st.warning("Heatmap requires at least two numeric columns.")
                return

            corr = plot_df[num_cols].corr(numeric_only=True)

            fig = px.imshow(
                corr,
                text_auto=".2f",
                aspect="auto",
                color_continuous_scale="RdBu_r",
                title="Correlation Heatmap",
            )

        st.plotly_chart(fig, use_container_width=True, key=f"plotly_chart_{key_prefix}")

    except Exception as e:
        st.warning(f"Could not render {chart_type}: {e}")


# --- 4. MAIN LAYOUT TABS ---
tab1, tab2, tab3 = st.tabs(
    ["💬 NL Query Engine", "📈 Data Quality & Profiling", "🔍 Raw Data Explorer"]
)

# --- TAB 1: Natural Language Query Engine ---
with tab1:
    st.subheader("Ask Questions in Plain English")

    # Preset Prompt Shortcuts
    st.write("**Quick Insights:**")
    col1, col2, col3 = st.columns(3)
    preset_query = None

    with col1:
        if st.button("📊 Executive Overview"):
            preset_query = (
                "Provide an executive summary showing total sales, total orders,"
                " average order value, and average customer age."
            )
    with col2:
        if st.button("🏆 Category Performance"):
            preset_query = (
                "Group by categorical columns (like customer_segment or age_segment)"
                " and show total sales for each group sorted descending."
            )
    with col3:
        if st.button("⚠️ Outlier Summary"):
            preset_query = "Show me rows with unusual or extreme profit values."

    # User Input Field
    user_input = st.chat_input("e.g., Which region has the highest total sales?")
    query_to_run = user_input or preset_query

    # Display Previous Chat History
    for idx, chat in enumerate(st.session_state.chat_history):
        with st.chat_message(chat["role"]):
            st.markdown(chat["content"])
            if "code" in chat and chat["code"]:
                with st.expander("Show Generated Python Code"):
                    st.code(chat["code"], language="python")
            if "data" in chat and chat["data"] is not None:
                st.dataframe(chat["data"], use_container_width=True)
                render_smart_charts(chat["data"], key_prefix=f"history_{idx}")

                # Task C4: Export Buttons for past responses
                st.markdown("---")
                exp_col1, exp_col2 = st.columns(2)
                active_narrative = chat.get("content", "No narrative available.")
                res_df = chat.get("data", None)
                filter_summary = f"Active Rows: {len(filtered_df):,}"

                with exp_col1:
                    pdf_data = generate_pdf_report(
                        dataset_name="Global E-Commerce Sales",
                        total_rows=len(filtered_df),
                        total_cols=len(filtered_df.columns),
                        filters_applied=filter_summary,
                        narrative_text=active_narrative,
                        summary_df=res_df,
                    )
                    st.download_button(
                        label="📄 Export PDF Report",
                        data=pdf_data,
                        file_name=f"Analytics_Report_{idx}.pdf",
                        mime="application/pdf",
                        key=f"pdf_btn_{idx}",
                        use_container_width=True,
                    )

                with exp_col2:
                    docx_data = generate_docx_report(
                        dataset_name="Global E-Commerce Sales",
                        total_rows=len(filtered_df),
                        total_cols=len(filtered_df.columns),
                        filters_applied=filter_summary,
                        narrative_text=active_narrative,
                        summary_df=res_df,
                    )
                    st.download_button(
                        label="📝 Export Word (.docx)",
                        data=docx_data,
                        file_name=f"Analytics_Report_{idx}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key=f"docx_btn_{idx}",
                        use_container_width=True,
                    )

    # Process New Query
    if query_to_run:
        with st.chat_message("user"):
            st.markdown(query_to_run)

        with st.spinner("🤖 Writing code and querying dataset..."):
            context_str = st.session_state.memory.get_context_string()
            active_df = filtered_df.copy()
            active_schema = generate_schema_catalog(active_df)
            # SECURITY SANBOX RATIONALE:
            # Executing LLM-generated strings directly poses an arbitrary code execution risk.

            exec_res = nl_query_pipeline(
                query_to_run, active_df, active_schema, memory_context=context_str
            )

        if exec_res["success"]:
            narrative = format_result_with_llm(query_to_run, exec_res["result"])
            st.session_state.memory.add_interaction(query_to_run, narrative)

            # Format result data properly
            res_data = exec_res["result"]
            if isinstance(res_data, pd.Series):
                res_data = res_data.reset_index()

            if isinstance(res_data, pd.DataFrame):
                res_data = res_data.copy().reset_index(drop=True)
                cat_cols = res_data.select_dtypes(
                    include=["object", "category"]
                ).columns.tolist()
                num_cols = res_data.select_dtypes(include=["number"]).columns.tolist()
                if cat_cols and num_cols:
                    res_data = res_data.groupby(cat_cols, as_index=False)[
                        num_cols
                    ].sum()

            with st.chat_message("assistant"):
                st.markdown(narrative)

                with st.expander("Show Generated Python Code"):
                    # SECURITY SANBOX RATIONALE:
                    # All critical system modules (os, sys, subprocess) and built-in file operations

                    st.code(exec_res["code"], language="python")

                if isinstance(res_data, pd.DataFrame):
                    st.dataframe(res_data, use_container_width=True)
                    render_smart_charts(res_data, key_prefix="current_query")

            # Save to chat history
            st.session_state.chat_history.append(
                {"role": "user", "content": query_to_run}
            )
            st.session_state.chat_history.append(
                {
                    "role": "assistant",
                    "content": narrative,
                    # SECURITY SANBOX RATIONALE:
                    # Executing LLM-generated strings directly poses an arbitrary code execution risk.
                    "code": exec_res["code"],
                    "data": res_data if isinstance(res_data, pd.DataFrame) else None,
                }
            )
            st.rerun()
        else:
            with st.chat_message("assistant"):
                # SECURITY SANBOX RATIONALE:
                # are entirely absent from this allowlist namespace.

                st.error(f"❌ Failed to execute query. Error: {exec_res['error']}")
# --- TAB 2: Data Quality & Profiling ---
with tab2:
    st.header("📊 Data Quality & Profiling")

    # Metrics Summary Bar
    col_a, col_b = st.columns(2)
    with col_a:
        st.metric("Total Rows (Filtered)", f"{len(filtered_df):,}")
    with col_b:
        st.metric("Total Columns", len(filtered_df.columns))

    st.divider()

    # Schema Catalog Expander
    with st.expander("📋 View Full Dataset Schema & Catalog", expanded=False):
        st.json(schema_str)

    # General IQR Outlier Summary Table
    st.markdown("### 🎯 Dataset Outlier Overview (Summary)")
    outliers = detect_iqr_outliers(filtered_df)
    st.dataframe(pd.DataFrame(outliers).T, use_container_width=True)

    st.divider()

    # --- TASK D3: AUTOMATED ANOMALY & OUTLIER DETECTION ENGINE ---
    st.subheader("⚠️ Task D3: AI-Powered Anomaly & Outlier Narrative")

    active_df = filtered_df.copy()
    numeric_cols = active_df.select_dtypes(include=["number"]).columns.tolist()

    if numeric_cols:
        col_to_scan = st.selectbox(
            "Select Numeric Metric to Analyze for Statistical Anomalies:",
            numeric_cols,
        )

        if st.button("🔍 Run Anomaly Detection"):
            with st.spinner(
                "Scanning distribution with IQR and generating LLM report..."
            ):
                anomalies_df, narrative = analyze_anomalies_iqr(
                    active_df, col_to_scan, client
                )

                # Display LLM Narrative Card
                st.markdown("### 🤖 LLM Anomaly Insight")
                st.info(narrative)

                # Display Box Plot showing outliers visually
                st.markdown("### 📊 Distribution & Outlier Visualization")
                fig = px.box(
                    active_df,
                    y=col_to_scan,
                    points="outliers",
                    title=f"IQR Outlier Boxplot for '{col_to_scan}'",
                    template="plotly_dark",
                )
                st.plotly_chart(fig, use_container_width=True)

                # Display Anomaly Records Table
                st.markdown(f"### 📋 Flagged Anomaly Records ({len(anomalies_df)})")
                if not anomalies_df.empty:
                    st.dataframe(anomalies_df, use_container_width=True)
                else:
                    st.success("No statistical anomalies detected for this metric!")
    else:
        st.warning("No numeric columns available in the active dataset.")
# --- TAB 3: Raw Data Explorer ---
with tab3:
    st.subheader("Filtered Dataset View")
    st.dataframe(filtered_df, use_container_width=True)

    # ------------------------------------------- ML model train ------------------
    # --- TASK D2: PREDICTIVE ANALYTICS ---
st.subheader("📈 Task D2: Predictive Analytics (Linear Regression)")

active_df = filtered_df.copy()
numeric_cols = active_df.select_dtypes(include=["number"]).columns.tolist()

if len(numeric_cols) >= 2:
    col_x, col_y = st.columns(2)

    with col_x:
        # Default to 'age' if present, otherwise first numeric
        default_x = next(
            (i for i, col in enumerate(numeric_cols) if "age" in col.lower()), 0
        )
        x_var = st.selectbox(
            "Select Predictor Variable (X):", numeric_cols, index=default_x
        )

    with col_y:
        # Default to 'sales' or 'profit' if present
        default_y = next(
            (
                i
                for i, col in enumerate(numeric_cols)
                if "sale" in col.lower() or "profit" in col.lower()
            ),
            min(1, len(numeric_cols) - 1),
        )
        y_var = st.selectbox(
            "Select Target Variable (Y):", numeric_cols, index=default_y
        )

    if x_var == y_var:
        st.warning("Please select two different numeric columns for X and Y.")
    else:
        if st.button("🚀 Run Predictive Model"):
            with st.spinner(
                f"Fitting Scikit-Learn Regression model for {y_var} vs {x_var}..."
            ):
                model, narrative, metrics = run_predictive_regression(
                    active_df, x_var, y_var, client
                )

                # 1. Display Model Summary Cards
                st.markdown("### 🔢 Model Performance Summary")
                m1, m2, m3 = st.columns(3)
                m1.metric("Slope Coefficient (m)", f"{metrics['coefficient']}")
                m2.metric("Intercept (b)", f"{metrics['intercept']}")
                m3.metric("R² Score", f"{metrics['r2_score']}")

                # 2. Display LLM Narrative Explanation
                st.markdown("### 🤖 LLM Regression Interpretation")
                st.info(narrative)

                # 3. Display Regression Scatter Plot with OLS Trendline
                st.markdown("### 📊 Linear Fit Visualization")
                fig = px.scatter(
                    active_df,
                    x=x_var,
                    y=y_var,
                    trendline="ols",
                    trendline_color_override="red",
                    title=f"OLS Linear Regression: {y_var} vs {x_var}",
                    template="plotly_dark",
                )
                st.plotly_chart(fig, use_container_width=True)
else:
    st.warning("At least two numeric columns are required to run Task D2.")

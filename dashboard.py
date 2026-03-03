import pandas as pd
import streamlit as st
from datetime import timedelta
import io
import streamlit.components.v1 as components
import openpyxl

# ===============================
# Streamlit Setup
# ===============================
st.set_page_config(page_title="Weekly Cancellation Pivot", layout="wide")
st.title("📊 Weekly Cancellation Pivot (Aloha + Zoho)")


# ===============================
# Helpers
# ===============================
def normalize_columns(df):
    df.columns = [str(c).strip().lower() for c in df.columns]
    return df


def normalize_aloha(df):
    df = df.copy()

    df["appt_date"] = pd.to_datetime(df["appt. date"], errors="coerce")
    df["dob"] = pd.to_datetime(df["date of birth"], errors="coerce")

    date_str = df["appt_date"].dt.strftime("%Y-%m-%d")

    df["start_dt"] = pd.to_datetime(
        date_str + " " + df["appt. start time"].astype(str), errors="coerce"
    )
    df["end_dt"] = pd.to_datetime(
        date_str + " " + df["appt. end time"].astype(str), errors="coerce"
    )

    df["billing hours"] = pd.to_numeric(df["billing hours"], errors="coerce")
    return df


def add_week_bucket(df):
    df = df.copy()
    df = df.dropna(subset=["appt_date"])

    week_start = df["appt_date"] - pd.to_timedelta(df["appt_date"].dt.weekday, unit="D")
    week_end = week_start + timedelta(days=6)

    df["week_start"] = week_start.dt.normalize()

    df["week"] = (
        week_start.dt.month.astype(str)
        + "/"
        + week_start.dt.day.astype(str)
        + "/"
        + week_start.dt.year.astype(str)
        + " - "
        + week_end.dt.month.astype(str)
        + "/"
        + week_end.dt.day.astype(str)
        + "/"
        + week_end.dt.year.astype(str)
    )
    return df


def derive_billing_hours(df):
    df = df.copy()
    missing = df["billing hours"].isna()
    df.loc[missing, "billing hours"] = (
        (df.loc[missing, "end_dt"] - df.loc[missing, "start_dt"])
        .dt.total_seconds()
        .div(3600)
    )
    return df


def classify_cancel_bucket(status):
    s = str(status).lower()
    if "no show" in s or "last minute" in s:
        return "Last Minute Client Cancel/No Show"
    if "client" in s:
        return "Client Cancellation"
    if "staff" in s:
        return "Staff Cancellation"
    if "cancel" in s:
        return "Other Cancellation"
    return "Active"


def merge_case_coordinator(df_aloha, df_zoho):
    df = df_aloha.copy()
    z = df_zoho.copy()

    z["dob"] = pd.to_datetime(z["date of birth"], errors="coerce")

    df = df.merge(
        z[["medicaid id", "case coordinator name"]],
        left_on="insured id",
        right_on="medicaid id",
        how="left",
    )

    missing = df["case coordinator name"].isna()
    dob_lookup = (
        z[["dob", "case coordinator name"]]
        .dropna()
        .drop_duplicates("dob")
        .set_index("dob")["case coordinator name"]
    )

    df.loc[missing, "case coordinator name"] = df.loc[missing, "dob"].map(dob_lookup)
    return df.drop(columns=["medicaid id"], errors="ignore")


# ===============================
# Pivot Logic (Chronological)
# ===============================
def build_weekly_pivot(df):
    # Ensure numerical consistency
    pivot = pd.pivot_table(
        df,
        index=["week_start", "week", "case coordinator name"],
        columns=["completed_flag", "cancel_bucket"],
        values="billing hours",
        aggfunc="sum",
        fill_value=0,
    )

    pivot.columns = [f"{a} | {b}" for a, b in pivot.columns]

    yes_cols = [c for c in pivot.columns if c.startswith("Yes |")]
    blank_cols = [c for c in pivot.columns if c.startswith("(blank) |")]

    pivot["Yes Total"] = pivot[yes_cols].sum(axis=1)
    pivot["(blank) Total"] = pivot[blank_cols].sum(axis=1)
    pivot["Grand Total"] = pivot["Yes Total"] + pivot["(blank) Total"]

    cancel_cols = [
        c
        for c in pivot.columns
        if any(
            x in c
            for x in [
                "Client Cancellation",
                "Staff Cancellation",
                "Last Minute",
                "Other Cancellation",
            ]
        )
    ]

    pivot["Cancellation Total"] = pivot[cancel_cols].sum(axis=1)
    pivot["Cancellation Percentage"] = (
        pivot["Cancellation Total"] / pivot["Grand Total"] * 100
    ).fillna(0).round(2)

    week_totals = pivot.groupby(["week_start", "week"]).sum(numeric_only=True)
    week_totals["case coordinator name"] = "Total"
    week_totals["Cancellation Percentage"] = (
        week_totals["Cancellation Total"] / week_totals["Grand Total"] * 100
    ).fillna(0).round(2)
    week_totals = week_totals.reset_index()

    combined = pd.concat([pivot.reset_index(), week_totals], ignore_index=True)

    # Calculate WoW Differences
    combined = combined.sort_values(["case coordinator name", "week_start"])
    combined["Difference Yes Total"] = combined.groupby("case coordinator name")[
        "Yes Total"
    ].diff()
    combined["Difference Grand Total"] = combined.groupby("case coordinator name")[
        "Grand Total"
    ].diff()
    combined["Difference Total Cancellation"] = combined.groupby(
        "case coordinator name"
    )["Cancellation Total"].diff()

    combined["sort_helper"] = combined["case coordinator name"].apply(
        lambda x: 1 if x == "Total" else 0
    )

    return combined.sort_values(
        ["week_start", "sort_helper", "case coordinator name"]
    ).drop(columns=["sort_helper"])


# ===============================
# HTML Renderer
# ===============================
def render_html(df):
    df = df.copy()
    yes_cols = [c for c in df.columns if c.startswith("Yes |")]
    blank_cols = [c for c in df.columns if c.startswith("(blank) |")]

    def fmt(v, pct=False):
        if pd.isna(v) or v == "":
            return ""
        return f"{v:.2f}" if pct else f"{v:,.2f}"

    css = """
    <style>
      table { border-collapse: collapse; font-family: Arial; font-size: 11px; border:2px solid #333; background:white; width: 100%; }
      th, td { border:1px solid #999; padding:4px 6px; white-space:nowrap; }
      th { background:#f2f2f2; font-weight:700; text-align:center; position: sticky; top: 0; }
      td.num { text-align:right; }
      tr.week-header td { font-weight:700; background:#eee; border-top:2px solid #333; }
      tr.week-total td { background:#efe6c8; font-weight:700; border-top:2px solid #333; }
      td.total { font-weight:700; }
      td.diff { color:#00008B; font-style:italic; font-weight:700; }
    </style>
    """

    header = f"""
    <thead>
      <tr>
        <th rowspan="2">Week</th>
        <th rowspan="2">Staff Name</th>
        <th colspan="{len(yes_cols)}">Yes</th>
        <th rowspan="2">Yes Total</th>
        <th rowspan="2">Diff Yes</th>
        <th colspan="{len(blank_cols)}">(blank)</th>
        <th rowspan="2">(blank) Total</th>
        <th rowspan="2">Grand Total</th>
        <th rowspan="2">Cancel %</th>
        <th rowspan="2">Diff Grand</th>
        <th rowspan="2">Diff Cancel</th>
      </tr>
      <tr>
        {''.join(f"<th>{c.replace('Yes | ','')}</th>" for c in yes_cols)}
        {''.join(f"<th>{c.replace('(blank) | ','')}</th>" for c in blank_cols)}
      </tr>
    </thead>
    """

    body = []
    for week, g in df.groupby("week", sort=False):
        body.append(
            f"<tr class='week-header'><td>{week}</td><td colspan='{len(df.columns)}'></td></tr>"
        )

        for _, r in g.iterrows():
            row_class = "week-total" if r["case coordinator name"] == "Total" else ""
            cells = [f"<td></td>", f"<td>{r['case coordinator name']}</td>"]

            for c in yes_cols:
                cells.append(f"<td class='num'>{fmt(r[c])}</td>")
            cells.append(f"<td class='num total'>{fmt(r['Yes Total'])}</td>")
            cells.append(f"<td class='num diff'>{fmt(r['Difference Yes Total'])}</td>")

            for c in blank_cols:
                cells.append(f"<td class='num'>{fmt(r[c])}</td>")
            cells.append(f"<td class='num total'>{fmt(r['(blank) Total'])}</td>")

            cells.append(f"<td class='num total'>{fmt(r['Grand Total'])}</td>")
            cells.append(
                f"<td class='num'>{fmt(r['Cancellation Percentage'], True)}%</td>"
            )
            cells.append(
                f"<td class='num diff'>{fmt(r['Difference Grand Total'])}</td>"
            )
            cells.append(
                f"<td class='num diff'>{fmt(r['Difference Total Cancellation'])}</td>"
            )

            body.append(f"<tr class='{row_class}'>" + "".join(cells) + "</tr>")

    return css + "<table>" + header + "<tbody>" + "".join(body) + "</tbody></table>"


# ===============================
# Main App
# ===============================
aloha_file = st.file_uploader("Upload Aloha", type=["xlsx"])
zoho_file = st.file_uploader("Upload Zoho", type=["xlsx"])

if aloha_file and zoho_file:

    a = normalize_columns(pd.read_excel(aloha_file))
    z = normalize_columns(pd.read_excel(zoho_file))

    a = normalize_aloha(a)
    
    # --- SERVICE FILTER LOGIC ---
    if "service name" in a.columns:
        a = a[a["service name"].astype(str).str.strip().str.lower() == "direct service bt"].copy()
    
    a = add_week_bucket(a)
    a = derive_billing_hours(a)

    a["completed_flag"] = (
        a["completed"]
        .astype(str)
        .str.strip()
        .str.lower()
        .apply(lambda x: "Yes" if x == "yes" else "(blank)")
    )

    df = merge_case_coordinator(a, z)
    df["cancel_bucket"] = df["appointment status"].apply(classify_cancel_bucket)

    tab1, tab2 = st.tabs(["📊 Weekly Cancellation Pivot", "📈 Billing Growth Trend"])

    with tab1:
        if not df.empty:
            min_date = df["appt_date"].min()
            max_date = df["appt_date"].max()

            date_range_tab1 = st.date_input(
                "Filter Date Range (Tab 1)",
                value=(min_date, max_date),
                min_value=min_date,
                max_value=max_date,
                key="tab1_date_range",
            )

            df_tab1 = df.copy()
            if len(date_range_tab1) == 2:
                start_date, end_date = pd.to_datetime(date_range_tab1[0]), pd.to_datetime(date_range_tab1[1])
                df_tab1 = df_tab1[(df_tab1["appt_date"] >= start_date) & (df_tab1["appt_date"] <= end_date)].copy()

            if not df_tab1.empty:
                pivot = build_weekly_pivot(df_tab1)
                html = render_html(pivot)

                col1, col2 = st.columns(2)
                with col1:
                    merged_out = io.BytesIO()
                    with pd.ExcelWriter(merged_out, engine="openpyxl") as w:
                        df_tab1.to_excel(w, index=False, sheet_name="Merged_Data")
                    st.download_button("⬇️ Download Merged Excel", merged_out.getvalue(), "merged_bt_data.xlsx", use_container_width=True)

                with col2:
                    st.download_button("⬇️ Download Table (HTML)", html.encode("utf-8"), "pivot.html", "text/html", use_container_width=True)

                components.html(html, height=800, scrolling=True)
            else:
                st.warning("No data found for the selected date range.")
        else:
            st.warning("No 'Direct Service BT' records found in the uploaded Aloha file.")

    with tab2:
        st.subheader("Billing Growth by Case Coordinator")
        if not df.empty:
            view_type = st.radio("View Type", ["Weekly", "Monthly"], horizontal=True)
            
            trend_df = df.copy()
            if view_type == "Weekly":
                trend_df["period"] = trend_df["appt_date"] - pd.to_timedelta(trend_df["appt_date"].dt.weekday, unit="D")
            else:
                trend_df["period"] = trend_df["appt_date"].dt.to_period("M").dt.to_timestamp()

            summary = trend_df.groupby(["period", "case coordinator name"])["billing hours"].sum().reset_index().sort_values("period")
            pivot_view = summary.pivot(index="period", columns="case coordinator name", values="billing hours").fillna(0)

            st.line_chart(pivot_view)
            growth = pivot_view.pct_change() * 100
            st.markdown(f"### 📈 {view_type} Growth %")
            st.dataframe(growth.round(2), use_container_width=True)

else:
    st.info("Upload both files to generate the dashboard.")
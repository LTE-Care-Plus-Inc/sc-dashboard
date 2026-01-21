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
    df.columns = df.columns.str.strip().str.lower()
    return df


def normalize_aloha(df):
    df = df.copy()
    df["appt_date"] = pd.to_datetime(df["appt. date"], errors="coerce")
    df["dob"] = pd.to_datetime(df["date of birth"], errors="coerce")

    df["start_dt"] = pd.to_datetime(
        df["appt. date"].astype(str) + " " + df["appt. start time"].astype(str),
        errors="coerce"
    )
    df["end_dt"] = pd.to_datetime(
        df["appt. date"].astype(str) + " " + df["appt. end time"].astype(str),
        errors="coerce"
    )

    df["billing hours"] = pd.to_numeric(df["billing hours"], errors="coerce")
    return df


def add_week_bucket(df):
    df = df.copy()
    week_start = df["appt_date"] - pd.to_timedelta(df["appt_date"].dt.weekday, unit="D")
    week_end = week_start + timedelta(days=6)

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
# Pivot
# ===============================
def build_weekly_pivot(df):
    pivot = pd.pivot_table(
        df,
        index=["week", "case coordinator name"],
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
        c for c in pivot.columns
        if any(x in c for x in ["Client Cancellation", "Staff Cancellation", "Last Minute", "Other Cancellation"])
    ]
    
    pivot["Cancellation Total"] = pivot[cancel_cols].sum(axis=1)
    pivot["Cancellation Percentage"] = (pivot["Cancellation Total"] / pivot["Grand Total"] * 100).round(2)

    # --- ADD DIFFERENCE LOGIC ---
    # We sort by name then week to ensure the diff() compares the same person across time
    pivot = pivot.reset_index().sort_values(["case coordinator name", "week"])
    
    pivot["Difference Yes Total"] = pivot.groupby("case coordinator name")["Yes Total"].diff()
    pivot["Difference Grand Total"] = pivot.groupby("case coordinator name")["Grand Total"].diff()
    pivot["Difference Total Cancellation"] = pivot.groupby("case coordinator name")["Cancellation Total"].diff()

    # Re-sort to keep the week-based view for the UI
    return pivot.sort_values(["week", "case coordinator name"])


# ===============================
# HTML Renderer (UPDATED to include new columns)
# ===============================
def render_html(df):
    df = df.copy()
    yes_cols = [c for c in df.columns if c.startswith("Yes |")]
    blank_cols = [c for c in df.columns if c.startswith("(blank) |")]

    def fmt(v, pct=False):
        if pd.isna(v) or v == "": return ""
        return f"{v:.2f}" if pct else f"{v:,.2f}"

    css = """
    <style>
      table { border-collapse: collapse; font-family: Arial; font-size: 12px; border:2px solid #333; background:white; }
      th, td { border:1px solid #999; padding:4px 8px; white-space:nowrap; }
      th { background:#f2f2f2; font-weight:700; text-align:center; }
      td.num { text-align:right; }
      tr.week-header td { font-weight:700; background:#eee; border-top:2px solid #333; }
      tr.grand td { border-top:3px solid #333; background:#efe6c8; font-weight:700; }
      td.total { background:#f9f9f9; font-weight:700; }
      td.diff { color: #00008B; font-style: italic; } /* Distinguish diff columns */
    </style>
    """

    header = f"""
    <thead>
      <tr>
        <th rowspan="2">Week</th>
        <th rowspan="2">Staff Name</th>
        <th colspan="{len(yes_cols)}">Yes</th>
        <th rowspan="2">Yes Total</th>
        <th rowspan="2" style="color:blue">Diff Yes</th>
        <th colspan="{len(blank_cols)}">(blank)</th>
        <th rowspan="2">(blank) Total</th>
        <th rowspan="2">Grand Total</th>
        <th rowspan="2">Cancel %</th>
        <th rowspan="2" style="color:blue">Diff Grand</th>
        <th rowspan="2" style="color:blue">Diff Cancel</th>
      </tr>
      <tr>
        {''.join(f"<th>{c.replace('Yes | ','')}</th>" for c in yes_cols)}
        {''.join(f"<th>{c.replace('(blank) | ','')}</th>" for c in blank_cols)}
      </tr>
    </thead>
    """

    body = []
    for week, g in df.groupby("week"):
        # Week Summary Row
        body.append(f"<tr class='week-header'><td>{week}</td><td colspan='{len(df.columns)}'></td></tr>")
        
        for _, r in g.iterrows():
            cells = [
                f"<td></td>",
                f"<td>{r['case coordinator name']}</td>"
            ]
            # Yes Section
            for c in yes_cols: cells.append(f"<td class='num'>{fmt(r[c])}</td>")
            cells.append(f"<td class='num total'>{fmt(r['Yes Total'])}</td>")
            cells.append(f"<td class='num diff'>{fmt(r['Difference Yes Total'])}</td>")
            
            # Blank Section
            for c in blank_cols: cells.append(f"<td class='num'>{fmt(r[c])}</td>")
            cells.append(f"<td class='num total'>{fmt(r['(blank) Total'])}</td>")
            
            # Totals Section
            cells.append(f"<td class='num total'>{fmt(r['Grand Total'])}</td>")
            cells.append(f"<td class='num'>{fmt(r['Cancellation Percentage'], True)}</td>")
            
            # Differences
            cells.append(f"<td class='num diff'>{fmt(r['Difference Grand Total'])}</td>")
            cells.append(f"<td class='num diff'>{fmt(r['Difference Total Cancellation'])}</td>")
            
            body.append("<tr>" + "".join(cells) + "</tr>")

    # Grand Total row (usually diffs aren't summed in a grand total, so we leave those blank)
    gt = df.select_dtypes("number").sum()
    grand_cells = [
        "<td>Grand Total</td><td></td>"
    ]
    for c in yes_cols: grand_cells.append(f"<td class='num total'>{fmt(gt.get(c, 0))}</td>")
    grand_cells.append(f"<td class='num total'>{fmt(gt['Yes Total'])}</td>")
    grand_cells.append("<td></td>") # No diff for grand total
    for c in blank_cols: grand_cells.append(f"<td class='num total'>{fmt(gt.get(c, 0))}</td>")
    grand_cells.append(f"<td class='num total'>{fmt(gt['(blank) Total'])}</td>")
    grand_cells.append(f"<td class='num total'>{fmt(gt['Grand Total'])}</td>")
    grand_cells.append(f"<td class='num total'>{fmt(df['Cancellation Percentage'].mean(), True)}</td>")
    grand_cells.extend(["<td></td>", "<td></td>"])

    body.append("<tr class='grand'>" + "".join(grand_cells) + "</tr>")

    return css + "<table>" + header + "<tbody>" + "".join(body) + "</tbody></table>"


# ===============================
# Upload
# ===============================
aloha_file = st.file_uploader("Upload Aloha", type=["xlsx"])
zoho_file = st.file_uploader("Upload Zoho", type=["xlsx"])

if aloha_file and zoho_file:
    a = normalize_columns(pd.read_excel(aloha_file))
    z = normalize_columns(pd.read_excel(zoho_file))

    a = normalize_aloha(a)
    a = add_week_bucket(a)
    a = derive_billing_hours(a)

    a["completed_flag"] = (
        a["completed"].astype(str).str.strip().str.lower()
        .apply(lambda x: "Yes" if x == "yes" else "(blank)")
    )

    df = merge_case_coordinator(a, z)
    df["cancel_bucket"] = df["appointment status"].apply(classify_cancel_bucket)

    pivot = build_weekly_pivot(df)

    html = render_html(pivot)
    components.html(html, height=900, scrolling=True)

    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as w:
        pivot.to_excel(w, index=False)
    out.seek(0)

    st.download_button("⬇️ Download Excel", out, "weekly_pivot.xlsx")

else:
    st.info("Upload both files.")

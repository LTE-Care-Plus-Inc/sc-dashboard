import pandas as pd
import streamlit as st
from datetime import timedelta
import io
import streamlit.components.v1 as components

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
        errors="coerce",
    )
    df["end_dt"] = pd.to_datetime(
        df["appt. date"].astype(str) + " " + df["appt. end time"].astype(str),
        errors="coerce",
    )

    df["billing hours"] = pd.to_numeric(df["billing hours"], errors="coerce")
    return df


def add_week_bucket(df):
    """
    Monday → Sunday, Windows/Mac safe formatting
    """
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

    # Primary: Insured ID → Medicaid ID
    df = df.merge(
        z[["medicaid id", "case coordinator name"]],
        left_on="insured id",
        right_on="medicaid id",
        how="left",
    )

    # Fallback: DOB
    missing = df["case coordinator name"].isna()
    dob_lookup = (
        z[["dob", "case coordinator name"]]
        .dropna()
        .drop_duplicates("dob")
        .set_index("dob")["case coordinator name"]
    )
    df.loc[missing, "case coordinator name"] = df.loc[missing, "dob"].map(dob_lookup)

    return df.drop(columns=["medicaid id"], errors="ignore")


def build_weekly_pivot(df):
    pivot = pd.pivot_table(
        df,
        index=["week", "case coordinator name"],
        columns=["completed_flag", "cancel_bucket"],
        values="billing hours",
        aggfunc="sum",
        fill_value=0,
    )

    # Flatten columns
    pivot.columns = [f"{c0} | {c1}" for c0, c1 in pivot.columns]

    # Totals
    yes_cols = [c for c in pivot.columns if c.startswith("Yes |")]
    blank_cols = [c for c in pivot.columns if c.startswith("(blank) |")]

    pivot["Yes Total"] = pivot[yes_cols].sum(axis=1) if yes_cols else 0
    pivot["(blank) Total"] = pivot[blank_cols].sum(axis=1) if blank_cols else 0
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

    pivot["Cancellation Percentage"] = (
        pivot[cancel_cols].sum(axis=1) / pivot["Grand Total"] * 100
    ).round(2)

    return pivot.reset_index()


# ===============================
# HTML Renderer (Two Headers + Forced Light Theme)
# ===============================
def pivot_to_two_header_html(pivot_df: pd.DataFrame) -> str:
    df = pivot_df.copy()

    # Keep week only once per group
    df["week"] = df["week"].where(df["week"].ne(df["week"].shift()), "")

    yes_cols = [c for c in df.columns if c.startswith("Yes |")]
    blank_cols = [c for c in df.columns if c.startswith("(blank) |")]

    display_cols = (
        ["week", "case coordinator name"]
        + yes_cols
        + (["Yes Total"] if "Yes Total" in df.columns else [])
        + blank_cols
        + (["(blank) Total"] if "(blank) Total" in df.columns else [])
        + (["Grand Total"] if "Grand Total" in df.columns else [])
        + (["Cancellation Percentage"] if "Cancellation Percentage" in df.columns else [])
    )
    df = df[display_cols]

    # ---------- FORMATTERS ----------
    def fmt(v, is_pct=False):
        if v == "" or pd.isna(v):
            return ""
        try:
            v = float(v)
        except Exception:
            return str(v)
        return f"{v:.2f}" if is_pct else f"{v:,.2f}"

    # ---------- GRAND TOTAL ROW ----------
    numeric_cols = [c for c in df.columns if c not in ["week", "case coordinator name"]]

    grand_totals = {}
    for c in numeric_cols:
        if c == "Cancellation Percentage":
            # recompute from totals if possible
            if "Grand Total" in df.columns:
                cancel_cols = [
                    col for col in df.columns
                    if any(x in col for x in [
                        "Client Cancellation",
                        "Staff Cancellation",
                        "Last Minute",
                        "Other Cancellation"
                    ])
                ]
                total_cancel = df[cancel_cols].sum().sum()
                grand_total = df["Grand Total"].sum()
                grand_totals[c] = (total_cancel / grand_total * 100) if grand_total else 0
            else:
                grand_totals[c] = ""
        else:
            grand_totals[c] = df[c].sum()

    # ---------- CSS (DARKER LINES) ----------
    css = """
    <style>
      .wrap {
        background: white;
        padding: 12px;
        overflow-x: auto;
      }

      table.pivot {
        border-collapse: collapse;
        font-family: Arial, sans-serif;
        font-size: 13px;
        background: white !important;
        color: black !important;
        border: 2px solid #333;              /* OUTER BORDER */
      }

      .pivot th, .pivot td {
        border: 1px solid #999;
        padding: 6px 10px;
        white-space: nowrap;
        background: white !important;
        color: black !important;
      }

      .pivot thead th {
        font-weight: 700;
        text-align: center;
        background: #f2f2f2 !important;
        border-bottom: 2px solid #333;       /* HEADER SEPARATOR */
      }

      .pivot td.week {
        font-weight: 700;
      }

      .pivot td.name {
        padding-left: 18px;
      }

      .pivot td.num {
        text-align: right;
      }

      /* Totals columns */
      .pivot th.total,
      .pivot td.totalcol {
        background: #efe6c8 !important;
        font-weight: 700;
      }

      /* Thick separator between week groups */
      tr.week-start td {
        border-top: 2px solid #333 !important;
      }

      /* Grand total row */
      tr.grand-total td {
        border-top: 3px solid #333 !important;
        font-weight: 700;
        background: #efe6c8 !important;
      }
    </style>
    """

    # ---------- HEADER ----------
    header1 = f"""
      <tr>
        <th rowspan="2" style="text-align:left;">Week W/ SC</th>
        <th rowspan="2" style="text-align:left;"></th>
        <th colspan="{len(yes_cols)}">Yes</th>
        <th rowspan="2" class="total">Yes Total</th>
        <th colspan="{len(blank_cols)}">(blank)</th>
        <th rowspan="2" class="total">(blank) Total</th>
        <th rowspan="2" class="total">Grand Total</th>
        <th rowspan="2" class="total">Cancellation %</th>
      </tr>
    """

    header2 = "<tr>"
    header2 += "".join(f"<th>{c.replace('Yes | ', '')}</th>" for c in yes_cols)
    header2 += "".join(f"<th>{c.replace('(blank) | ', '')}</th>" for c in blank_cols)
    header2 += "</tr>"

    # ---------- BODY ----------
    body_rows = []
    prev_week = None

    for _, r in df.iterrows():
        is_week_start = r["week"] != "" and r["week"] != prev_week
        prev_week = r["week"] if r["week"] else prev_week

        tr_class = "week-start" if is_week_start else ""

        tds = [
            f"<td class='week'>{r['week']}</td>",
            f"<td class='name'>{r['case coordinator name']}</td>",
        ]

        for c in numeric_cols:
            is_pct = c == "Cancellation Percentage"
            cls = "num totalcol" if c in ["Yes Total", "(blank) Total", "Grand Total"] else "num"
            tds.append(f"<td class='{cls}'>{fmt(r[c], is_pct)}</td>")

        body_rows.append(f"<tr class='{tr_class}'>" + "".join(tds) + "</tr>")

    # ---------- GRAND TOTAL ROW ----------
    gt_cells = [
        "<td class='week'>Grand Total</td>",
        "<td></td>",
    ]

    for c in numeric_cols:
        is_pct = c == "Cancellation Percentage"
        cls = "num totalcol"
        gt_cells.append(f"<td class='{cls}'>{fmt(grand_totals[c], is_pct)}</td>")

    body_rows.append("<tr class='grand-total'>" + "".join(gt_cells) + "</tr>")

    html = f"""
    {css}
    <div class="wrap">
      <table class="pivot">
        <thead>
          {header1}
          {header2}
        </thead>
        <tbody>
          {''.join(body_rows)}
        </tbody>
      </table>
    </div>
    """
    return html


# ===============================
# Upload Files
# ===============================
aloha_file = st.file_uploader("Upload Aloha Appointment Billing Info", type=["xlsx"])
zoho_file = st.file_uploader("Upload Zoho Case List", type=["xlsx"])

if aloha_file and zoho_file:
    df_aloha = normalize_columns(pd.read_excel(aloha_file))
    df_zoho = normalize_columns(pd.read_excel(zoho_file))

    # Prepare Aloha
    df_aloha = normalize_aloha(df_aloha)
    df_aloha = add_week_bucket(df_aloha)
    df_aloha = derive_billing_hours(df_aloha)

    # Completed: strict Yes vs (blank)
    df_aloha["completed_flag"] = (
        df_aloha["completed"]
        .astype(str)
        .str.strip()
        .str.lower()
        .apply(lambda x: "Yes" if x == "yes" else "(blank)")
    )

    # Merge Zoho
    df = merge_case_coordinator(df_aloha, df_zoho)

    # Cancellation bucket
    df["cancel_bucket"] = df["appointment status"].apply(classify_cancel_bucket)

    # Pivot
    pivot_df = build_weekly_pivot(df)

    # Render HTML (properly, not as code)
    html = pivot_to_two_header_html(pivot_df)
    components.html(html, height=750, scrolling=True)

    # Download HTML export
    st.download_button(
        "⬇️ Download HTML",
        data=html.encode("utf-8"),
        file_name="weekly_cancellation_pivot.html",
        mime="text/html",
    )

    # Export Excel (raw pivot)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        pivot_df.to_excel(writer, index=False, sheet_name="Pivot")
    output.seek(0)

    st.download_button(
        "⬇️ Download Excel",
        data=output,
        file_name="weekly_cancellation_pivot.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

else:
    st.info("Please upload both Aloha and Zoho Excel files to continue.")

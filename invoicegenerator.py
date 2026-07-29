import streamlit as st
import pandas as pd
from io import BytesIO
from fpdf import FPDF
from datetime import datetime
from num2words import num2words
from decimal import Decimal, ROUND_HALF_UP
import re

# ─────────────────────────────────────────────────────
# Decimal helpers — all rate/percentage math below is
# done in Decimal, not float. Floats like 0.15 or sums
# such as 1113.10 cannot always be represented exactly
# in binary, so a "clean" value like 166.965 can quietly
# become 166.96499999999997 before it's ever rounded.
# Doing the multiplication itself in Decimal avoids that,
# and round2() then always rounds half up (5 in the 3rd
# decimal rounds the 2nd decimal up) instead of Python's
# built-in round(), which uses banker's rounding.
# ─────────────────────────────────────────────────────
def D(value):
    return Decimal(str(value))

def round2(value):
    return float(D(value).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))

# ─────────────────────────────────────────────────────
# Page config
# ─────────────────────────────────────────────────────
st.set_page_config(page_title="Payroll Summary & Invoice", layout="wide")
st.title("🧾 Invoice Generator")

# ─────────────────────────────────────────────────────
# Upload Payroll Summary Excel
# ─────────────────────────────────────────────────────
uploaded = st.file_uploader(
    "Upload Payroll Summary Excel (with ACTIVE / ABSCOND sections)",
    type=["xlsx"]
)

if not uploaded:
    st.stop()

# ─────────────────────────────────────────────────────
# Detect department from filename
# ─────────────────────────────────────────────────────
filename = uploaded.name.upper()

# Capture department name before first underscore
dept_match = re.match(r"([A-Z ]+?)_", filename)

department = (
    dept_match.group(1).strip()
    if dept_match
    else "UNKNOWN"
)


# ─────────────────────────────────────────────────────
# Invoice Month Controls
# ─────────────────────────────────────────────────────
invoice_month_base = st.text_input(
    "Invoice Month",
    value=f"Month of {datetime.now().strftime('%B %Y')}"
)

is_new_hire = st.checkbox("New Hire Invoice")

invoice_month = (
    f"{invoice_month_base} - NEW HIRE"
    if is_new_hire
    else invoice_month_base
)


# ─────────────────────────────────────────────────────
# Read raw sheet to detect sections
# ─────────────────────────────────────────────────────
raw = pd.read_excel(uploaded, sheet_name=0, header=None)

def find_row(label):
    rows = raw.index[
        raw.iloc[:, 0].astype(str).str.strip().str.upper() == label
    ].tolist()
    return rows[0] if rows else None

active_row = find_row("ACTIVE EMPLOYEES")
abscond_row = find_row("ABSCOND / RESIGN")

if active_row is None or abscond_row is None:
    st.error("ACTIVE EMPLOYEES / ABSCOND / RESIGN markers not found")
    st.stop()

header_row = active_row + 1
nrows = abscond_row - header_row - 1

if nrows <= 0:
    st.error("No active employee data detected")
    st.stop()

# ─────────────────────────────────────────────────────
# Read ACTIVE EMPLOYEES table
# ─────────────────────────────────────────────────────
df = pd.read_excel(
    uploaded,
    sheet_name=0,
    header=header_row,
    nrows=nrows
)

df.columns = [str(c).strip() for c in df.columns]

def numcol(name):
    if name in df.columns:
        return pd.to_numeric(df[name], errors="coerce").fillna(0.0)
    return pd.Series(0.0, index=df.index)

# ─────────────────────────────────────────────────────
# Payroll Calculations
# ─────────────────────────────────────────────────────
OT_total = (
    numcol("OT 1.5 (Amount)") +
    numcol("OT 2.0 (Amount)") +
    numcol("P.H 2.0 (Amount)") +
    numcol("OT 3.0 (Amount)")
)

gross = numcol("Gross Pay")

if gross.sum() == 0:
    gross = (
        numcol("Monthly Salary") +
        OT_total +
        numcol("Morning Shift") +
        numcol("Night Shift") +
        numcol("Performance Incentive") +
        numcol("Incentive Programme") +
        numcol("Recognition Award") +
        numcol("Annual Leave") +
        numcol("Backpay (BAC, BSC, BBB)") +
        numcol("Backpay Shift Allowance") +
        numcol("Backpay OT")
    )

wages = round2((gross - OT_total).sum())
ot_sum = round2(OT_total.sum())

emp_stat = round2((
    numcol("EPF ER") +
    numcol("Socso ER") +
    numcol("EIS ER")
).sum())

hrdf = round2(numcol("HRDF").sum())
medical = round2(numcol("Medical Fee").sum())

active_headcount = len(df)

# ─────────────────────────────────────────────────────
# Sidebar Controls
# ─────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Invoice Settings")
    mgmt_rate = st.number_input("Management Fee %", value=15.0) / 100
    sst_rate = st.number_input("SST %", value=8.0) / 100
    insurance_fee = st.number_input("Insurance Fee / Head (RM)", value=50.0)

insurance_qty = st.number_input(
    "Insurance Quantity",
    min_value=0,
    step=1,
    value=int(active_headcount)
)

insurance_amount = round2(D(insurance_qty) * D(insurance_fee))
mgmt_fee = round2((D(wages) + D(ot_sum) + D(emp_stat) + D(hrdf)) * D(mgmt_rate))

# ─────────────────────────────────────────────────────
# Invoice Table
# ─────────────────────────────────────────────────────
invoice_df = pd.DataFrame([
    (1, "Wages", 1, wages, wages),
    (2, "Overtime (OT)", 1, ot_sum, ot_sum),
    (3, "Employer Statutory (EPF + SOCSO + EIS)", 1, emp_stat, emp_stat),
    (4, "HRDF", 1, hrdf, hrdf),
    (5, "Medical Fee (Exclude Management Fee)", 1, medical, medical),
    (6, "Insurance Claim (Exclude Management Fee)", insurance_qty, insurance_fee, insurance_amount),
    (7, f"{int(mgmt_rate*100)}% Management Fee", 1, mgmt_fee, mgmt_fee),
], columns=["No.", "Description", "Qty", "U.Price", "Amount"])

subtotal = round2(sum(D(a) for a in invoice_df["Amount"]))
sst = round2(D(subtotal) * D(sst_rate))
total = round2(D(subtotal) + D(sst))

# ─────────────────────────────────────────────────────
# Preview
# ─────────────────────────────────────────────────────
st.subheader(f"🧾 Invoice Preview — {department}")
st.markdown(f"**{invoice_month}**")

st.dataframe(
    invoice_df.style.format({
        "U.Price": "{:,.2f}",
        "Amount": "{:,.2f}"
    }),
    use_container_width=True
)

st.markdown(f"**Sub-Total:** RM {subtotal:,.2f}")
st.markdown(f"**SST ({int(sst_rate*100)}%):** RM {sst:,.2f}")
st.markdown(f"## **TOTAL (Inclusive SST): RM {total:,.2f}**")

# ─────────────────────────────────────────────────────
# PDF Export
# ─────────────────────────────────────────────────────
def export_pdf(df, invoice_month, department):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)

    # ───────────── Page Title ─────────────
    pdf.set_font("Arial", "B", 16)
    pdf.cell(0, 12, "SUMMARY", ln=True, align="C")
    pdf.ln(6)


    # Header
    # ───────────── Header Blocks (Aligned) ─────────────

    pdf.set_font("Arial", "", 9)

    # Capture starting Y for alignment
    header_y = pdf.get_y()

    # Left: Company details
    pdf.multi_cell(
        90, 5,
        "DEXCOM MALAYSIA SDN BHD\n"
        "PMT 818, Persiaran Cassia Selatan 3, Taman Perindustrian Batu Kawan,\n"
        "Batu Kawan,\n"
        "Bandar Cassia, Batu Kawan, 14110 Simpang Ampat, Pulau Pinang, Malaysia.\n"
        "Attention: HR Department"
    )

    # Capture bottom of left block
    left_block_bottom = pdf.get_y()

    # Right: Invoice meta (aligned to top of left block)
    pdf.set_xy(120, header_y)
    pdf.set_font("Arial", "B", 9)

    line_gap = 6
    pdf.cell(0, line_gap, "Summary No.:", ln=True)
    pdf.set_x(120)
    pdf.cell(0, line_gap, "Date:", ln=True)
    pdf.set_x(120)
    pdf.cell(0, line_gap, "Terms:", ln=True)
    pdf.set_x(120)
    pdf.cell(0, line_gap, "PO Number:", ln=True)

    # Move cursor below both blocks safely
    pdf.set_y(max(left_block_bottom, pdf.get_y()) + 10)



    # Move safely below header blocks
    pdf.set_y(left_block_bottom + 10)

    # Invoice title
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 8, f"Payroll Invoice - {department}", ln=True, align="C")

    # Month (BOLD)
    pdf.set_font("Arial", "B", 11)
    pdf.cell(0, 8, invoice_month, ln=True, align="C")
    pdf.ln(4)

    # Table
    pdf.set_font("Arial", "B", 9)
    widths = [10, 85, 15, 30, 30]
    headers = ["No.", "Description", "Qty", "U.Price", "Amount"]

    for i, h in enumerate(headers):
        pdf.cell(widths[i], 8, h, 1, align="C")
    pdf.ln()

    pdf.set_font("Arial", "", 9)
    for _, r in df.iterrows():
        pdf.cell(widths[0], 8, str(int(r["No."])), 1, align="C")
        pdf.cell(widths[1], 8, r["Description"], 1)
        pdf.cell(widths[2], 8, str(int(r["Qty"])), 1, align="C")
        pdf.cell(widths[3], 8, f"{r['U.Price']:,.2f}", 1, align="R")
        pdf.cell(widths[4], 8, f"{r['Amount']:,.2f}", 1, align="R")
        pdf.ln()

    # Totals
    pdf.ln(4)
    pdf.cell(140, 6, "Sub - Total", 0)
    pdf.cell(40, 6, f"RM {subtotal:,.2f}", 0, ln=True, align="R")

    pdf.cell(140, 6, f"SST {int(sst_rate*100)}%", 0)
    pdf.cell(40, 6, f"RM {sst:,.2f}", 0, ln=True, align="R")

    pdf.set_font("Arial", "B", 9)
    pdf.cell(140, 7, "Total (Inclusive SST)", 0)
    pdf.cell(40, 7, f"RM {total:,.2f}", 0, ln=True, align="R")

    # Amount in words
    pdf.ln(6)
    ringgit = int(total)
    sen = int(round((total - ringgit) * 100))

    words = (
    f"{num2words(ringgit, lang='en').title()} Ringgit"
    + (f" And {num2words(sen, lang='en').title()} Sen" if sen else "")
)


    pdf.set_font("Arial", "", 9)
    pdf.multi_cell(0, 5, f"Ringgit Malaysia:\n{words} Only.")
    pdf.ln(14)

    # Signatures
    pdf.cell(80, 6, "______________________________", 0, 0, "C")
    pdf.cell(30, 6, "", 0)
    pdf.cell(80, 6, "______________________________", 0, 1, "C")

    pdf.cell(80, 6, "Prepared By", 0, 0, "C")
    pdf.cell(30, 6, "", 0)
    pdf.cell(80, 6, "Authorised Signature(s)", 0, 1, "C")

    return BytesIO(pdf.output(dest="S").encode("latin-1"))

# ─────────────────────────────────────────────────────
# Download Button
# ─────────────────────────────────────────────────────
st.download_button(
    "📄 Download Invoice (PDF)",
    data=export_pdf(invoice_df, invoice_month, department),
    file_name=f"Invoice_{department}_{invoice_month.replace(' ', '_')}.pdf",
    mime="application/pdf",
    use_container_width=True
)

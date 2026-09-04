import io
import numpy as np
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Commercial Valuation Engine", layout="wide")

st.title("Commercial Valuation & Contract Pricing Engine")
st.write(
    "Download the multi-sheet benchmark template, adjust operational shifts "
    "and headcount, and upload to calculate NPV, IRR, and breakeven pricing."
)

# OpenPyXL Styles matching commercial_valuation.py
header_fill = PatternFill(start_color="1F4E78",
                          end_color="1F4E78", fill_type="solid")
header_fill_blue = PatternFill(
    start_color="2E75B6", end_color="2E75B6", fill_type="solid")
section_fill_blue = PatternFill(
    start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")

thin_border = Border(
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"),
    bottom=Side(style="thin", color="D9D9D9"),
)
header_border = Border(
    left=Side(style="thin", color="000000"),
    right=Side(style="thin", color="000000"),
    top=Side(style="medium", color="000000"),
    bottom=Side(style="medium", color="000000"),
)

headers_def = [
    ("Project_Name", "Project / Tender identifier.",
     "@", Alignment(horizontal="left")),
    ("Managers", "Number of dedicated site managers.",
     "#,##0", Alignment(horizontal="center")),
    ("Supervisors", "Number of operational supervisors.",
     "#,##0", Alignment(horizontal="center")),
    ("Workers", "Number of cleaning line workers.",
     "#,##0", Alignment(horizontal="center")),
    ("Shift_Option", "1=8h/7d (240h), 2=8h/6d (208h), 3=12h/7d (360h), 4=12h/6d (312h)",
     "0", Alignment(horizontal="center")),
    ("Contract_Years", "Contract tenor in years.",
     "0", Alignment(horizontal="center")),
    ("Capex_Initial", "Upfront machinery/equipment (EGP).",
     "#,##0", Alignment(horizontal="right")),
    ("Markup_Pct", "Target commercial markup %.",
     "0.00%", Alignment(horizontal="right")),
    ("WACC", "Discount rate / cost of capital %.",
     "0.00%", Alignment(horizontal="right")),
    ("Tax_Rate", "Corporate income tax rate %.",
     "0.00%", Alignment(horizontal="right")),
    ("DSO_Days", "Days Sales Outstanding (days).",
     "0", Alignment(horizontal="center")),
    ("DPO_Days", "Days Payable Outstanding (days).",
     "0", Alignment(horizontal="center")),
]

data_rows = [
    ("Project_ABC_Cairo", 1, 5, 36, 1, 3, 2170200.0, 0.1285, 0.4575, 0.225, 60, 30),
    ("Branch_Carrefour_Alex", 1, 2, 18, 2, 3,
     950000.0, 0.1285, 0.4575, 0.225, 45, 30),
    ("Logistics_Hub_Giza", 2, 8, 75, 3, 5,
     4800000.0, 0.1500, 0.4000, 0.225, 60, 30),
    ("Tech_Campus_NewCairo", 1, 3, 25, 1, 3,
     1300000.0, 0.1350, 0.4575, 0.225, 60, 30),
    ("Retail_Hypermarket_Zayed", 1, 4, 30, 4, 3,
     1750000.0, 0.1285, 0.4575, 0.225, 60, 30),
]


def create_template_bytes() -> bytes:
    wb = openpyxl.Workbook()
    ws_inputs = wb.active
    ws_inputs.title = "INPUTS"
    ws_ref = wb.create_sheet(title="REFERENCE_GUIDE")

    for col_idx, (col_name, _, _, _) in enumerate(headers_def, start=1):
        cell = ws_inputs.cell(row=1, column=col_idx, value=col_name)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True)
        cell.border = header_border

    for row_idx, row_data in enumerate(data_rows, start=2):
        for col_idx, val in enumerate(row_data, start=1):
            cell = ws_inputs.cell(row=row_idx, column=col_idx, value=val)
            cell.border = thin_border
            cell.font = Font(name="Calibri", size=10)
            cell.number_format = headers_def[col_idx - 1][2]
            cell.alignment = headers_def[col_idx - 1][3]

    ws_inputs.row_dimensions[1].height = 28

    ws_ref["A1"] = "Commercial Pricing Model - Parameter & Benchmark Guide"
    ws_ref["A1"].font = Font(name="Calibri", size=14,
                             bold=True, color="1F4E78")

    guide_tables = [
        ("Shift Option Reference (from INPUTS sheet in model)", 3, [
            ("Shift_Option Code", "Working Schedule",
             "Monthly Hours / Worker", "Standard Application"),
            (1, "8 Hours / 7 Days a week", 240,
             "Standard retail, malls, and hypermarkets"),
            (2, "8 Hours / 6 Days a week", 208,
             "Corporate headquarters and administrative offices"),
            (3, "12 Hours / 7 Days a week", 360,
             "24/7 continuous industrial facilities & factories"),
            (4, "12 Hours / 6 Days a week", 312,
             "Warehouses and logistics distribution centers"),
        ]),
        ("Benchmark Salary & Burden Structure (EGP / Month)", 11, [
            ("Designation", "Net Base Salary",
             "Loaded Total Labor Cost", "Statutory Burdens Included"),
            ("Manager", 12000, 18176,
             "GOSI (18.75% Co / 11% Emp) + Income Tax + Medical + Uniform"),
            ("Supervisor", 9000, 15887, "GOSI + Income Tax + Medical + Uniform"),
            ("Worker", 7000, 11607, "GOSI + Income Tax + Medical + Uniform"),
        ]),
        ("Operating Cost Factoring Ratios (Calibrated from B.E.P Sheet)", 18, [
            ("Cost Category", "Ratio to Base Payroll",
             "Excel Sheet Source", "Commercial Description"),
            ("Statutory Burdens", "43.13%", "B.E.P (Row 13-17)",
             "Company GOSI (18.75%), Labor Tax, Health, Uniforms"),
            ("Labor Transportation", "41.33%", "B.E.P (Row 29)",
             "Dedicated outsourced passenger buses/vans"),
            ("Cleaning Consumables", "4.58%", "B.E.P (Row 30)",
             "Chemicals, trash bags, mop heads, dispensers"),
            ("Machinery Maintenance", "1.67%", "B.E.P (Row 28)",
             "Scrubbers, sweepers, burnishers preventative servicing"),
            ("SG&A Overhead", "9.77%", "B.E.P (Row 31) / O.H1",
             "Head office admin, operations supervisors, corporate HQ"),
            ("Commercial / Mun. Taxes", "11.62%", "B.E.P (Row 32)",
             "Municipal licenses, contract stamp duties, stamp taxes"),
        ]),
        ("Financial & Capital Budgeting Defaults", 28, [
            ("Financial Parameter", "Benchmark Value", "Notes & Regulatory Basis"),
            ("WACC (Discount Rate)", "45.75%",
             "Calibrated in WACC sheet: Ke = 45.75% (Rf=24.84%, Beta=1.5, ERP=13.94%)"),
            ("Corporate Tax Rate", "22.50%", "Egyptian Corporate Income Tax Code"),
            ("Target Markup", "12.85%",
             "Contract commercial margin over total operating and capital cost"),
            ("DSO (Receivables Lag)", "60 Days", "Client billing payment cycle"),
            ("DPO (Payables Lag)", "30 Days",
             "Vendor material and subcontractor credit terms"),
            ("Capex Benchmark", "51,670 EGP / Head",
             "Standard machinery & setup investment if not quoted directly"),
        ]),
    ]

    for title, start_row, rows in guide_tables:
        ws_ref.cell(row=start_row, column=1, value=title).font = Font(
            name="Calibri", size=11, bold=True, color="1F4E78")
        for r_idx, row_data in enumerate(rows):
            curr_r = start_row + 1 + r_idx
            for c_idx, val in enumerate(row_data):
                cell = ws_ref.cell(row=curr_r, column=c_idx + 1, value=val)
                if r_idx == 0:
                    cell.fill = section_fill_blue
                    cell.font = Font(name="Calibri", size=10, bold=True)
                else:
                    cell.border = thin_border
                    cell.font = Font(name="Calibri", size=10)

    for ws in [ws_inputs, ws_ref]:
        for col in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 14)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def run_valuation_pipeline(df_in: pd.DataFrame) -> pd.DataFrame:
    shift_hours_map = {1: 240.0, 2: 208.0, 3: 360.0, 4: 312.0}
    results = []

    for _, row in df_in.iterrows():
        name = str(row["Project_Name"])
        mgr, sup, wrk = int(row["Managers"]), int(
            row["Supervisors"]), int(row["Workers"])
        shift, tenor = int(row["Shift_Option"]), int(row["Contract_Years"])
        markup, wacc, tax_rate = float(row["Markup_Pct"]), float(
            row["WACC"]), float(row["Tax_Rate"])
        dso, dpo = float(row["DSO_Days"]), float(row["DPO_Days"])

        total_hc = mgr + sup + wrk
        shift_mult = shift_hours_map.get(shift, 240.0) / 240.0

        c_base = ((mgr * 12000.0) + (sup * 9000.0 * 1.166) +
                  (wrk * 7000.0 * 1.166)) * shift_mult
        c_labor = c_base * (1.0 + 0.4313)
        c_transport = c_base * 0.4133
        c_consumables = c_base * 0.0458
        c_maintenance = c_base * 0.0167
        c_sga_overhead = c_base * 0.0977
        c_opex_monthly = c_labor + c_transport + \
            c_consumables + c_maintenance + c_sga_overhead

        capex = float(row["Capex_Initial"]) if pd.notna(
            row["Capex_Initial"]) else total_hc * 51670.0
        monthly_depr = capex / (tenor * 12.0)
        c_com_tax = c_base * 0.1162

        c_total_monthly = c_opex_monthly + monthly_depr + c_com_tax
        monthly_price = c_total_monthly * (1.0 + markup)
        annual_rev = monthly_price * 12.0

        annual_ebit = (monthly_price - c_opex_monthly - monthly_depr) * 12.0
        nwc_drag_annual = (annual_rev * (dso / 365.0) -
                           (c_opex_monthly * 12.0) * (dpo / 365.0)) / tenor
        annual_fcf = (annual_ebit * (1.0 - tax_rate)) + \
            (monthly_depr * 12.0) - nwc_drag_annual

        t_steps = np.arange(1, tenor + 1, dtype=np.float64)
        pv_fcf = np.sum(annual_fcf / np.power(1.0 + wacc, t_steps))
        npv = pv_fcf - capex

        cf_stream = [-capex] + [annual_fcf] * tenor
        rate = 0.30
        for _ in range(25):
            disc = np.power(1.0 + rate, -np.arange(len(cf_stream)))
            f_val = np.sum(cf_stream * disc)
            f_prime = np.sum(-np.arange(len(cf_stream)) * cf_stream *
                             np.power(1.0 + rate, -np.arange(len(cf_stream)) - 1.0))
            if abs(f_prime) < 1e-8:
                break
            step = f_val / f_prime
            rate -= step
            if abs(step) < 1e-5:
                break
        irr = rate * 100.0 if np.isfinite(rate) else np.nan

        sum_disc = np.sum(1.0 / np.power(1.0 + wacc, t_steps))
        cost_burden = (
            (c_opex_monthly * 12.0) * (1.0 - tax_rate)
            - (monthly_depr * 12.0) * tax_rate
            - (c_opex_monthly * 12.0) * (dpo / (365.0 * tenor))
        )
        rev_coeff = (12.0 * (1.0 - tax_rate) -
                     (dso / (365.0 * tenor)) * 12.0) / 12.0
        be_monthly_price = (capex + sum_disc * cost_burden) / \
            (sum_disc * rev_coeff * 12.0)

        c5_price_p50 = 10089.86 * np.power(total_hc, 1.0874)

        results.append({
            "Project_Name": name,
            "Total_Headcount": total_hc,
            "Contract_Years": tenor,
            "Monthly_Cost_EGP": round(c_total_monthly, 2),
            "Quoted_Monthly_Price_EGP": round(monthly_price, 2),
            "Annual_Revenue_EGP": round(annual_rev, 2),
            "Total_Capex_EGP": round(capex, 2),
            "Annual_FCF_EGP": round(annual_fcf, 2),
            "NPV_EGP": round(npv, 2),
            "IRR_Pct": round(irr, 2),
            "Breakeven_Monthly_Price_EGP": round(be_monthly_price, 2),
            "Class5_CER_Benchmark_P50": round(c5_price_p50, 2),
            "Class5_Lower_Bound": round(c5_price_p50 * 0.70, 2),
            "Class5_Upper_Bound": round(c5_price_p50 * 1.50, 2),
        })

    return pd.DataFrame(results)


def export_results_excel(df_out: pd.DataFrame, df_in: pd.DataFrame) -> bytes:
    wb_out = openpyxl.Workbook()
    ws_out1 = wb_out.active
    ws_out1.title = "VALUATION_OUTPUTS"
    ws_out2 = wb_out.create_sheet(title="INPUT_SOURCE")

    ws_out1.append(list(df_out.columns))
    for cell in ws_out1[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for _, r in df_out.iterrows():
        ws_out1.append(list(r))

    for r in range(2, ws_out1.max_row + 1):
        for c in range(1, ws_out1.max_column + 1):
            cell = ws_out1.cell(row=r, column=c)
            cell.border = thin_border
            if c == 1:
                cell.alignment = Alignment(horizontal="left")
            elif c in [2, 3]:
                cell.alignment = Alignment(horizontal="center")
                cell.number_format = "#,##0"
            elif c == 10:
                cell.alignment = Alignment(horizontal="right")
                cell.number_format = "0.00"
            else:
                cell.alignment = Alignment(horizontal="right")
                cell.number_format = "#,##0.00"

    ws_out2.append(list(df_in.columns))
    for cell in ws_out2[1]:
        cell.fill = header_fill_blue
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for _, r in df_in.iterrows():
        ws_out2.append(list(r))

    for r in range(2, ws_out2.max_row + 1):
        for c in range(1, ws_out2.max_column + 1):
            cell = ws_out2.cell(row=r, column=c)
            cell.border = thin_border
            if c == 1:
                cell.alignment = Alignment(horizontal="left")
            elif c in [2, 3, 4, 5, 6, 11, 12]:
                cell.alignment = Alignment(horizontal="center")
            elif c in [8, 9, 10]:
                cell.alignment = Alignment(horizontal="right")
                cell.number_format = "0.00%"
            else:
                cell.alignment = Alignment(horizontal="right")
                cell.number_format = "#,##0"

    for ws in [ws_out1, ws_out2]:
        for col in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 14)

    buf = io.BytesIO()
    wb_out.save(buf)
    return buf.getvalue()


# Streamlit Layout & Controls
col1, col2 = st.columns([1, 2])
with col1:
    st.subheader("1. Download Benchmark Template")
    template_data = create_template_bytes()
    st.download_button(
        label="Download Pricing_Inputs_Template.xlsx",
        data=template_data,
        file_name="Pricing_Inputs_Template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

with col2:
    st.subheader("2. Upload & Run Pipeline")
    uploaded_file = st.file_uploader(
        "Upload filled template (.xlsx)", type=["xlsx"])

input_df = None
if uploaded_file is not None:
    input_df = pd.read_excel(uploaded_file, sheet_name="INPUTS")
elif st.checkbox("Load benchmark scenario presets", value=True):
    input_df = pd.DataFrame(data_rows, columns=[h[0] for h in headers_def])

if input_df is not None:
    st.write("### Active Input Tender Parameters")
    st.dataframe(input_df, use_container_width=True)

    results_df = run_valuation_pipeline(input_df)

    st.write("### Commercial Valuation Outputs")
    st.dataframe(
        results_df.style.format({
            "Total_Headcount": "{:,}",
            "Contract_Years": "{:,}",
            "Monthly_Cost_EGP": "{:,.2f}",
            "Quoted_Monthly_Price_EGP": "{:,.2f}",
            "Annual_Revenue_EGP": "{:,.2f}",
            "Total_Capex_EGP": "{:,.2f}",
            "Annual_FCF_EGP": "{:,.2f}",
            "NPV_EGP": "{:,.2f}",
            "IRR_Pct": "{:.2f}%",
            "Breakeven_Monthly_Price_EGP": "{:,.2f}",
            "Class5_CER_Benchmark_P50": "{:,.2f}",
            "Class5_Lower_Bound": "{:,.2f}",
            "Class5_Upper_Bound": "{:,.2f}",
        }),
        use_container_width=True,
    )

    output_excel = export_results_excel(results_df, input_df)
    st.download_button(
        label="Download Commercial_Valuation_Results.xlsx",
        data=output_excel,
        file_name="Commercial_Valuation_Results.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

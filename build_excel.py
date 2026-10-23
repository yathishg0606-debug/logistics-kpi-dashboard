"""
Builds the Logistics KPI Dashboard workbook from data/logistics_shipments.csv.

Tabs:
  1. Raw Data      - full dataset + helper 'order_month' column (formula)
  2. KPI Summary   - formula-driven KPI tables (OTIF, lead time, cost, late-rate by carrier, monthly volume)
  3. Dashboard     - headline KPI cards + charts, all referencing KPI Summary
"""

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.utils import get_column_letter

SRC_CSV = "data/logistics_shipments.csv"
OUT_XLSX = "Logistics_KPI_Dashboard.xlsx"

FONT_NAME = "Arial"
HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(name=FONT_NAME, bold=True, color="FFFFFF", size=11)
TITLE_FONT = Font(name=FONT_NAME, bold=True, size=16, color="1F4E78")
SUBTITLE_FONT = Font(name=FONT_NAME, italic=True, size=10, color="666666")
KPI_LABEL_FONT = Font(name=FONT_NAME, size=11, color="666666")
KPI_VALUE_FONT = Font(name=FONT_NAME, bold=True, size=24, color="1F4E78")
BODY_FONT = Font(name=FONT_NAME, size=10)
THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CARD_FILL = PatternFill("solid", fgColor="EAF1F8")

df = pd.read_csv(SRC_CSV, dtype=str)
n_rows = len(df)
last_row = n_rows + 1  # +1 for header row

carriers = sorted(df["carrier"].unique().tolist())
regions = sorted(df["origin_region"].unique().tolist())

# Month labels present in the data, ordered chronologically
months = sorted(pd.to_datetime(df["order_date"]).dt.strftime("%Y-%m").unique().tolist())
month_display = [pd.to_datetime(m + "-01").strftime("%b-%y") for m in months]

wb = Workbook()

# ---------------------------------------------------------------- Raw Data
ws = wb.active
ws.title = "Raw Data"

headers = list(df.columns) + ["order_month"]
for col_idx, h in enumerate(headers, start=1):
    c = ws.cell(row=1, column=col_idx, value=h)
    c.font = HEADER_FONT
    c.fill = HEADER_FILL
    c.alignment = Alignment(horizontal="center")

for r, row in enumerate(df.itertuples(index=False), start=2):
    for c_idx, val in enumerate(row, start=1):
        if val == "" or pd.isna(val):
            continue  # leave truly blank (not empty-string) for cancelled orders
        col_name = df.columns[c_idx - 1]
        if col_name in ("quantity",):
            ws.cell(row=r, column=c_idx, value=int(val))
        elif col_name in ("order_value_eur", "shipment_cost_eur"):
            ws.cell(row=r, column=c_idx, value=float(val))
        elif col_name == "lead_time_days":
            ws.cell(row=r, column=c_idx, value=int(val))
        else:
            ws.cell(row=r, column=c_idx, value=val)
    # helper month column (formula, so it recalculates if order_date changes)
    month_col = len(df.columns) + 1
    ws.cell(row=r, column=month_col, value=f'=TEXT(B{r},"mmm-yy")')

for col_idx, h in enumerate(headers, start=1):
    ws.column_dimensions[get_column_letter(col_idx)].width = 16 if h != "order_id" else 12
ws.freeze_panes = "A2"

RD = "'Raw Data'"

# ---------------------------------------------------------------- KPI Summary
ws2 = wb.create_sheet("KPI Summary")
ws2.sheet_view.showGridLines = False

ws2["B2"] = "KPI Summary"
ws2["B2"].font = TITLE_FONT
ws2["B3"] = "All formulas reference the Raw Data tab and recalculate automatically"
ws2["B3"].font = SUBTITLE_FONT

# Headline KPIs block
row = 5
ws2.cell(row=row, column=2, value="Headline KPIs").font = Font(name=FONT_NAME, bold=True, size=12)
row += 1
headline_labels = [
    ("Total Orders", f"=COUNTA({RD}!A2:A{last_row})"),
    ("OTIF % (On-Time-In-Full)", f"=COUNTIF({RD}!M2:M{last_row},\"Yes\")/(COUNTIF({RD}!M2:M{last_row},\"Yes\")+COUNTIF({RD}!M2:M{last_row},\"No\"))"),
    ("Average Lead Time (days)", f"=AVERAGEIFS({RD}!L2:L{last_row},{RD}!K2:K{last_row},\"<>Cancelled\")"),
    ("Average Cost per Shipment (EUR)", f"=AVERAGEIFS({RD}!J2:J{last_row},{RD}!K2:K{last_row},\"<>Cancelled\")"),
    ("Cancelled Order Rate", f"=COUNTIF({RD}!K2:K{last_row},\"Cancelled\")/COUNTA({RD}!A2:A{last_row})"),
]
hl_start_row = row
for label, formula in headline_labels:
    ws2.cell(row=row, column=2, value=label).font = BODY_FONT
    fc = ws2.cell(row=row, column=4, value=formula)
    fc.font = Font(name=FONT_NAME, bold=True)
    if "%" in label or "Rate" in label:
        fc.number_format = "0.0%"
    elif "Cost" in label:
        fc.number_format = "€#,##0.00"
    elif "Lead Time" in label:
        fc.number_format = "0.0"
    row += 1

row += 2
# Late delivery rate by carrier
ws2.cell(row=row, column=2, value="Late Delivery Rate by Carrier").font = Font(name=FONT_NAME, bold=True, size=12)
row += 1
carrier_table_start = row
headers_c = ["Carrier", "Late Orders", "Completed Orders", "Late Rate"]
for i, h in enumerate(headers_c):
    cell = ws2.cell(row=row, column=2 + i, value=h)
    cell.font = HEADER_FONT
    cell.fill = HEADER_FILL
row += 1
carrier_data_start = row
for carrier in carriers:
    ws2.cell(row=row, column=2, value=carrier).font = BODY_FONT
    ws2.cell(row=row, column=3, value=f'=COUNTIFS({RD}!E2:E{last_row},B{row},{RD}!M2:M{last_row},"No")')
    ws2.cell(row=row, column=4, value=f'=COUNTIFS({RD}!E2:E{last_row},B{row},{RD}!M2:M{last_row},"Yes")+COUNTIFS({RD}!E2:E{last_row},B{row},{RD}!M2:M{last_row},"No")')
    rate_cell = ws2.cell(row=row, column=5, value=f"=C{row}/D{row}")
    rate_cell.number_format = "0.0%"
    rate_cell.font = Font(name=FONT_NAME, bold=True)
    row += 1
carrier_data_end = row - 1

row += 2
# Average lead time by destination region
ws2.cell(row=row, column=2, value="Average Lead Time by Destination Region").font = Font(name=FONT_NAME, bold=True, size=12)
row += 1
region_table_start = row
headers_r = ["Region", "Avg Lead Time (days)"]
for i, h in enumerate(headers_r):
    cell = ws2.cell(row=row, column=2 + i, value=h)
    cell.font = HEADER_FONT
    cell.fill = HEADER_FILL
row += 1
region_data_start = row
for reg in regions:
    ws2.cell(row=row, column=2, value=reg).font = BODY_FONT
    c = ws2.cell(row=row, column=3, value=f'=AVERAGEIFS({RD}!L2:L{last_row},{RD}!G2:G{last_row},B{row},{RD}!K2:K{last_row},"<>Cancelled")')
    c.number_format = "0.0"
    row += 1
region_data_end = row - 1

row += 2
# Monthly order volume trend
ws2.cell(row=row, column=2, value="Monthly Order Volume").font = Font(name=FONT_NAME, bold=True, size=12)
row += 1
month_table_start = row
headers_m = ["Month", "Order Count"]
for i, h in enumerate(headers_m):
    cell = ws2.cell(row=row, column=2 + i, value=h)
    cell.font = HEADER_FONT
    cell.fill = HEADER_FILL
row += 1
month_data_start = row
for md in month_display:
    ws2.cell(row=row, column=2, value=md).font = BODY_FONT
    ws2.cell(row=row, column=3, value=f'=COUNTIF({RD}!N2:N{last_row},B{row})')
    row += 1
month_data_end = row - 1

ws2.column_dimensions["B"].width = 26
ws2.column_dimensions["C"].width = 16
ws2.column_dimensions["D"].width = 18
ws2.column_dimensions["E"].width = 14

# ---------------------------------------------------------------- Dashboard
ws3 = wb.create_sheet("Dashboard")
ws3.sheet_view.showGridLines = False

ws3.merge_cells("B2:K2")
ws3["B2"] = "Logistics KPI Dashboard"
ws3["B2"].font = Font(name=FONT_NAME, bold=True, size=20, color="1F4E78")
ws3.merge_cells("B3:K3")
ws3["B3"] = "Synthetic shipment data | 2025 | Portfolio project"
ws3["B3"].font = SUBTITLE_FONT

KS = "'KPI Summary'"

def kpi_card(col, label, formula, number_format):
    letters = get_column_letter(col)
    ws3.merge_cells(f"{letters}5:{get_column_letter(col+1)}5")
    lab = ws3[f"{letters}5"]
    lab.value = label
    lab.font = KPI_LABEL_FONT
    lab.alignment = Alignment(horizontal="center")

    ws3.merge_cells(f"{letters}6:{get_column_letter(col+1)}7")
    val = ws3[f"{letters}6"]
    val.value = formula
    val.number_format = number_format
    val.font = KPI_VALUE_FONT
    val.alignment = Alignment(horizontal="center", vertical="center")
    for rr in range(5, 8):
        for cc in range(col, col + 2):
            ws3.cell(row=rr, column=cc).fill = CARD_FILL
            ws3.cell(row=rr, column=cc).border = BORDER

kpi_card(2, "Total Orders", f"={KS}!D{hl_start_row}", "#,##0")
kpi_card(4, "OTIF %", f"={KS}!D{hl_start_row+1}", "0.0%")
kpi_card(6, "Avg Lead Time (days)", f"={KS}!D{hl_start_row+2}", "0.0")
kpi_card(8, "Avg Cost / Shipment", f"={KS}!D{hl_start_row+3}", "€#,##0")
kpi_card(10, "Cancelled Rate", f"={KS}!D{hl_start_row+4}", "0.0%")

# Bar chart: late delivery rate by carrier
bar = BarChart()
bar.title = "Late Delivery Rate by Carrier"
bar.y_axis.title = "Late Rate"
bar.y_axis.numFmt = "0%"
bar.height = 8
bar.width = 16
data_ref = Reference(ws2, min_col=5, min_row=carrier_table_start, max_row=carrier_data_end)
cats_ref = Reference(ws2, min_col=2, min_row=carrier_data_start, max_row=carrier_data_end)
bar.add_data(data_ref, titles_from_data=True)
bar.set_categories(cats_ref)
bar.legend = None
ws3.add_chart(bar, "B10")

# Bar chart: avg lead time by region
bar2 = BarChart()
bar2.title = "Average Lead Time by Destination Region"
bar2.y_axis.title = "Days"
bar2.height = 8
bar2.width = 16
data_ref2 = Reference(ws2, min_col=3, min_row=region_table_start, max_row=region_data_end)
cats_ref2 = Reference(ws2, min_col=2, min_row=region_data_start, max_row=region_data_end)
bar2.add_data(data_ref2, titles_from_data=True)
bar2.set_categories(cats_ref2)
bar2.legend = None
ws3.add_chart(bar2, "B26")

# Line chart: monthly order volume trend
line = LineChart()
line.title = "Monthly Order Volume Trend"
line.y_axis.title = "Orders"
line.height = 8
line.width = 33
data_ref3 = Reference(ws2, min_col=3, min_row=month_table_start, max_row=month_data_end)
cats_ref3 = Reference(ws2, min_col=2, min_row=month_data_start, max_row=month_data_end)
line.add_data(data_ref3, titles_from_data=True)
line.set_categories(cats_ref3)
line.legend = None
ws3.add_chart(line, "J10")

for col in range(1, 13):
    ws3.column_dimensions[get_column_letter(col)].width = 12

wb.save(OUT_XLSX)
print(f"Saved {OUT_XLSX}")
print(f"carriers={carriers}")
print(f"regions={regions}")
print(f"months={month_display}")

# Logistics KPI Dashboard

A warehouse/shipping KPI dashboard built to track the metrics a logistics or supply chain analyst is typically asked to monitor: on-time delivery performance, lead times, shipment cost, and carrier reliability.

## Problem

Logistics teams need a quick, reliable view of delivery performance across carriers and regions to spot where service is slipping and where cost or delay risk is concentrated. This project simulates that process end-to-end: from raw shipment-level data to a KPI dashboard a logistics manager could act on.

## Data

`data/logistics_shipments.csv` is a **synthetic dataset** (1,500 shipment records, generated with `generate_data.py`) modeled on real-world logistics data structures (order ID, order/promised/actual delivery dates, carrier, origin/destination region, quantity, cost, and status). It was generated rather than sourced from a public dataset to avoid licensing restrictions and keep the project fully reproducible — carrier reliability differences are intentionally built in (e.g. GLS and DPD run higher delay rates than DHL and FedEx) so the KPIs tell a realistic story.

## Approach

1. **Data generation** — `generate_data.py` creates the shipment dataset with realistic order dates, promised vs. actual delivery dates, and carrier-specific delay probabilities.
2. **KPI definition** — selected five core logistics KPIs:
   - OTIF % (On-Time-In-Full delivery rate)
   - Average lead time (order to delivery, in days)
   - Average cost per shipment
   - Late delivery rate by carrier
   - Average lead time by destination region
   - Monthly order volume trend
3. **Excel build** (`Logistics_KPI_Dashboard.xlsx`) — `build_excel.py` builds a three-tab workbook:
   - **Raw Data** — full dataset
   - **KPI Summary** — all KPIs calculated with live formulas (`COUNTIFS`, `AVERAGEIFS`) referencing Raw Data, so the dashboard recalculates automatically if the data changes
   - **Dashboard** — headline KPI cards and charts (late delivery rate by carrier, lead time by region, monthly volume trend)
4. **Power BI** *(in progress)* — the same dataset rebuilt as an interactive Power BI report with slicers for date range, carrier, and region.

## Key Findings

- Overall OTIF rate across all carriers: **~90.2%**
- **GLS and DPD** show the highest late-delivery rates (~13–17%), more than double DHL and FedEx (~5–6%) — a strong candidate for carrier renegotiation or volume reallocation
- Average lead time sits at **~4.1 days** across all shipments
- Average cost per shipment: **~€61.70**

## Tools

Excel (formulas, charts) · Power BI *(planned)* · Python (`pandas`, `openpyxl`) for data generation and workbook build

## Files

```
├── data/
│   └── logistics_shipments.csv       # synthetic shipment dataset
├── generate_data.py                  # generates the dataset
├── build_excel.py                    # builds the Excel dashboard from the dataset
├── Logistics_KPI_Dashboard.xlsx      # the dashboard itself
└── README.md
```

## Author

Yathish Gangadhar — M.Sc. Systems Logistics, University of Maribor
[LinkedIn](https://www.linkedin.com/in/yathish-gangadhar-87215a1b9)

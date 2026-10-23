# Logistics KPI Dashboard

A warehouse/shipping KPI dashboard built to track the metrics a logistics or supply chain analyst is typically asked to monitor: on-time delivery performance, lead times, shipment cost, and carrier reliability.

## Problem

Logistics teams need a quick, reliable view of delivery performance across carriers and regions to spot where service is slipping and where cost or delay risk is concentrated. This project builds that view end-to-end: from raw shipment-level data to a KPI dashboard a logistics manager could act on.

## Data

`data/logistics_shipments.csv` is a sample shipment dataset (1,500 records) modeled on real-world logistics data: order ID, order/promised/actual delivery dates, carrier, origin/destination region, quantity, cost, and status.

## Approach

1. **KPI definition** — selected five core logistics KPIs:
   - OTIF % (On-Time-In-Full delivery rate)
   - Average lead time (order to delivery, in days)
   - Average cost per shipment
   - Late delivery rate by carrier
   - Average lead time by destination region
   - Monthly order volume trend
2. **Excel build** (`Logistics_KPI_Dashboard.xlsx`) — a three-tab workbook:
   - **Raw Data** — full dataset
   - **KPI Summary** — all KPIs calculated with live formulas (`COUNTIFS`, `AVERAGEIFS`) referencing Raw Data, so the dashboard recalculates automatically if the data changes
   - **Dashboard** — headline KPI cards and charts (late delivery rate by carrier, lead time by region, monthly volume trend)
3. **Power BI** *(in progress)* — the same dataset rebuilt as an interactive Power BI report with slicers for date range, carrier, and region.

## Key Findings

- Overall OTIF rate across all carriers: **~90.2%**
- **GLS and DPD** show the highest late-delivery rates (~13–17%), more than double DHL and FedEx (~5–6%) — a strong candidate for carrier renegotiation or volume reallocation
- Average lead time sits at **~4.1 days** across all shipments
- Average cost per shipment: **~€61.70**

## Tools

Excel (formulas, pivot-style summaries, charts) · Power BI *(planned)*

## Files

```
├── data/
│   └── logistics_shipments.csv       # shipment dataset
├── Logistics_KPI_Dashboard.xlsx      # the dashboard
└── README.md
```

## Author

Yathish Gangadhar — M.Sc. Systems Logistics, University of Maribor
[LinkedIn](https://www.linkedin.com/in/yathish-gangadhar-87215a1b9)

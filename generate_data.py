"""
Generates a synthetic (but realistic) logistics/shipment dataset for the
Warehouse & Logistics KPI Dashboard portfolio project.

Columns mirror what a real logistics ops/analyst dataset looks like:
order_id, order_date, promised_delivery_date, actual_delivery_date,
carrier, origin_region, destination_region, shipment_cost, order_value,
quantity, order_status

Data is entirely synthetic / randomly generated for demonstration purposes.
"""

import csv
import random
from datetime import date, timedelta

random.seed(42)

CARRIERS = ["DHL", "FedEx", "UPS", "DPD", "GLS"]
REGIONS = ["Central Europe", "Western Europe", "Southern Europe", "Northern Europe", "Eastern Europe"]
STATUSES_WEIGHTED = [("Delivered", 0.88), ("Delayed", 0.09), ("Cancelled", 0.03)]

START_DATE = date(2025, 1, 1)
END_DATE = date(2025, 12, 31)
N_ORDERS = 1500

# Carriers have different reliability profiles (used to bias delay odds)
CARRIER_DELAY_BIAS = {
    "DHL": 0.06,
    "FedEx": 0.07,
    "UPS": 0.08,
    "DPD": 0.14,
    "GLS": 0.16,
}

def random_date(start, end):
    delta = (end - start).days
    return start + timedelta(days=random.randint(0, delta))

def weighted_status(carrier):
    delay_p = CARRIER_DELAY_BIAS[carrier]
    r = random.random()
    if r < 0.03:
        return "Cancelled"
    elif r < 0.03 + delay_p:
        return "Delayed"
    else:
        return "Delivered"

rows = []
for i in range(1, N_ORDERS + 1):
    order_id = f"ORD-{10000 + i}"
    order_date = random_date(START_DATE, END_DATE)
    carrier = random.choice(CARRIERS)
    origin = random.choice(REGIONS)
    destination = random.choice(REGIONS)

    promised_days = random.choice([2, 3, 4, 5, 7])
    promised_delivery_date = order_date + timedelta(days=promised_days)

    status = weighted_status(carrier)

    if status == "Cancelled":
        actual_delivery_date = ""
        lead_time_days = ""
        on_time = ""
    else:
        if status == "Delayed":
            extra_days = random.randint(1, 6)
            actual_delivery_date = promised_delivery_date + timedelta(days=extra_days)
        else:
            # Delivered on time or a bit early
            jitter = random.randint(-1, 0)
            actual_delivery_date = promised_delivery_date + timedelta(days=jitter)
            if actual_delivery_date < order_date:
                actual_delivery_date = order_date + timedelta(days=1)

        lead_time_days = (actual_delivery_date - order_date).days
        on_time = "Yes" if actual_delivery_date <= promised_delivery_date else "No"
        actual_delivery_date = actual_delivery_date.isoformat()

    quantity = random.randint(1, 50)
    unit_cost = round(random.uniform(2.5, 40.0), 2)
    shipment_cost = round(unit_cost * quantity * random.uniform(0.08, 0.15), 2)
    order_value = round(unit_cost * quantity, 2)

    rows.append({
        "order_id": order_id,
        "order_date": order_date.isoformat(),
        "promised_delivery_date": promised_delivery_date.isoformat(),
        "actual_delivery_date": actual_delivery_date,
        "carrier": carrier,
        "origin_region": origin,
        "destination_region": destination,
        "quantity": quantity,
        "order_value_eur": order_value,
        "shipment_cost_eur": shipment_cost,
        "order_status": status,
        "lead_time_days": lead_time_days,
        "on_time_flag": on_time,
    })

out_path = "data/logistics_shipments.csv"
with open(out_path, "w", newline="") as f:
    writer = csv.DictWriter(f, fieldnames=rows[0].keys())
    writer.writeheader()
    writer.writerows(rows)

print(f"Wrote {len(rows)} rows to {out_path}")

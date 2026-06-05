import sqlite3
import random
import csv
from datetime import datetime, timedelta

import os
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

random.seed(42)

CLAIMANTS = [
    ("C001", "Alice Johnson"), ("C002", "Bob Martinez"), ("C003", "Carol White"),
    ("C004", "David Chen"), ("C005", "Emma Singh"), ("C006", "Frank Lee"),
    ("C007", "Grace Kim"), ("C008", "Henry Patel"), ("C009", "Irene Brown"),
    ("C010", "James Wilson"), ("C011", "Karen Davis"), ("C012", "Louis Tremblay"),
    ("C013", "Maria Garcia"), ("C014", "Nathan Scott"), ("C015", "Olivia Thompson"),
]

PAYMENT_TYPES = ["Weekly Benefit", "Pension", "Rehabilitation Allowance", "Short-Term Disability", "Medical Reimbursement"]
REGIONS = ["Vancouver", "Surrey", "Burnaby", "Richmond", "Abbotsford", "Kelowna", "Victoria", "Prince George"]

def random_date(start, end):
    return start + timedelta(days=random.randint(0, (end - start).days))

def generate_transactions(n=300):
    start = datetime(2024, 1, 1)
    end = datetime(2024, 12, 31)
    rows = []

    for i in range(1, n + 1):
        claimant_id, claimant_name = random.choice(CLAIMANTS)
        payment_type = random.choice(PAYMENT_TYPES)
        region = random.choice(REGIONS)
        date = random_date(start, end)

        # Normal payment range
        amount = round(random.uniform(800, 3000), 2)
        is_flagged = 0
        flag_reason = ""

        r = random.random()

        #fraud patterns
        if r < 0.04:  # Duplicate payment
            amount = round(random.uniform(800, 3000), 2)
            is_flagged = 1
            flag_reason = "Duplicate Payment"
        elif r < 0.08:  # Overpayment
            amount = round(random.uniform(5000, 12000), 2)
            is_flagged = 1
            flag_reason = "Amount Exceeds Threshold"
        elif r < 0.11:  # Deceased claimant still receiving
            is_flagged = 1
            flag_reason = "Inactive Claimant Status"
        elif r < 0.13:  # Multiple payments same day same claimant
            is_flagged = 1
            flag_reason = "Multiple Payments Same Day"

        rows.append({
            "transaction_id": f"TXN{i:05d}",
            "claimant_id": claimant_id,
            "claimant_name": claimant_name,
            "payment_type": payment_type,
            "amount": amount,
            "payment_date": date.strftime("%Y-%m-%d"),
            "region": region,
            "is_flagged": is_flagged,
            "flag_reason": flag_reason,
            "processed_by": f"EMP{random.randint(1,10):03d}",
        })

    return rows

def setup_db(rows):
    conn = sqlite3.connect(os.path.join(BASE_DIR, "data", "payments.db"))
    cur = conn.cursor()

    cur.execute("DROP TABLE IF EXISTS transactions")
    cur.execute("""
        CREATE TABLE transactions (
            transaction_id TEXT PRIMARY KEY,
            claimant_id TEXT,
            claimant_name TEXT,
            payment_type TEXT,
            amount REAL,
            payment_date TEXT,
            region TEXT,
            is_flagged INTEGER,
            flag_reason TEXT,
            processed_by TEXT
        )
    """)

    cur.executemany("""
        INSERT INTO transactions VALUES (
            :transaction_id, :claimant_id, :claimant_name, :payment_type,
            :amount, :payment_date, :region, :is_flagged, :flag_reason, :processed_by
        )
    """, rows)

    conn.commit()
    conn.close()
    print(f"Inserted {len(rows)} transactions into payments.db")

def export_csv(rows):
    path = os.path.join(BASE_DIR, "data", "raw_transactions.csv")
    with open(path, "w", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=rows[0].keys())
        writer.writeheader()
        writer.writerows(rows)
    print(f"Exported CSV to {path}")

if __name__ == "__main__":
    rows = generate_transactions(300)
    setup_db(rows)
    export_csv(rows)

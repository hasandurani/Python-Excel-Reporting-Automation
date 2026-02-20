import pandas as pd
import numpy as np
from datetime import datetime

print("Starting the Automated Reporting Bot...")

print("Fetching raw data from system...")
np.random.seed(42)
dates = pd.date_range(start="2026-02-01", periods=100, freq="h")
products = np.random.choice(["Laptops", "Smartphones", "Headphones", "Monitors"], 100)
regions = np.random.choice(["Islamabad", "Lahore", "Karachi"], 100)
sales = np.random.randint(5000, 150000, 100)
qty = np.random.randint(1, 10, 100)

df_raw = pd.DataFrame({
    "Date": dates,
    "Region": regions,
    "Product": products,
    "Quantity": qty,
    "Revenue_PKR": sales
})

print("⚙️ Processing Data & Creating Pivot Tables...")

pivot_report = pd.pivot_table(
    df_raw, 
    values="Revenue_PKR", 
    index="Region", 
    columns="Product", 
    aggfunc="sum", 
    fill_value=0
)

df_raw['Day'] = df_raw['Date'].dt.date
daily_summary = df_raw.groupby('Day').agg(
    Total_Orders=('Product', 'count'),
    Total_Revenue=('Revenue_PKR', 'sum')
).reset_index()

print("📊 Generating Formatted Excel Report...")

today_str = datetime.today().strftime('%Y-%m-%d')
excel_filename = f"Automated_Daily_Report_{today_str}.xlsx"


with pd.ExcelWriter(excel_filename, engine='openpyxl') as writer:
    pivot_report.to_excel(writer, sheet_name="Regional Summary")
    daily_summary.to_excel(writer, sheet_name="Daily Trends", index=False)
    df_raw.to_excel(writer, sheet_name="Raw System Data", index=False)

print(f"\n✅ BOOM! Task Completed.")
print(f"📁 Please check your folder for the file: '{excel_filename}'")
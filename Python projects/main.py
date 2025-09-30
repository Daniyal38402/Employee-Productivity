import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns

# plot style
sns.set_theme(style="whitegrid")
plt.rcParams["figure.figsize"] = (10,6)

# output folder for saved plots / tables
os.makedirs("outputs", exist_ok=True)



file_path = r"d:\Sales_Data.xlsx.xlsx" 

# read sheets
sales_df = pd.read_excel(file_path, sheet_name="Sales_Data")
state_df = pd.read_excel(file_path, sheet_name="State_list")
supervisor_df = pd.read_excel(file_path, sheet_name="Supervisor")

# quick checks
print("Sales sheet shape:", sales_df.shape)
print("Columns:", sales_df.columns.tolist())
print("State sheet cols:", state_df.columns.tolist())
print("Supervisor sheet cols:", supervisor_df.columns.tolist())


# 1) drop exact missing key rows
sales_df = sales_df.dropna(subset=["Order_Number", "State_Code", "Order_Date"], how="any")

# 2) drop duplicate orders if needed
sales_df = sales_df.drop_duplicates(subset=["Order_Number"])

# 3) merge State names (State_list assumed to have State_Code + State)
if "State_Code" in state_df.columns and "State_Code" in sales_df.columns:
    sales_df = sales_df.merge(state_df, on="State_Code", how="left")
else:
    print("Warning: State_Code column missing in one of sheets")

# 4) ensure date
sales_df["Order_Date"] = pd.to_datetime(sales_df["Order_Date"], errors="coerce")

# 5) numeric columns safe-convert
for c in ["Cost", "Sales", "Quantity", "Total_Cost", "Total_Sales"]:
    if c in sales_df.columns:
        sales_df[c] = pd.to_numeric(sales_df[c], errors="coerce")

# 6) create Total_Sales or Total_Cost if missing (best-effort)
if "Total_Sales" not in sales_df.columns:
    if {"Sales","Quantity"}.issubset(sales_df.columns):
        sales_df["Total_Sales"] = sales_df["Sales"] * sales_df["Quantity"]
        print("Computed Total_Sales = Sales * Quantity")
    else:
        print("Total_Sales missing and cannot compute (Sales/Quantity absent)")

if "Total_Cost" not in sales_df.columns:
    if {"Cost","Quantity"}.issubset(sales_df.columns):
        sales_df["Total_Cost"] = sales_df["Cost"] * sales_df["Quantity"]
        print("Computed Total_Cost = Cost * Quantity")
    else:
        print("Total_Cost missing and cannot compute (Cost/Quantity absent)")





sales_df["Year"] = sales_df["Order_Date"].dt.year
sales_df["Month"] = sales_df["Order_Date"].dt.month_name()
sales_df["Month_Num"] = sales_df["Order_Date"].dt.month
sales_df["Day"] = sales_df["Order_Date"].dt.day
sales_df["Weekday"] = sales_df["Order_Date"].dt.day_name()





print("Shape:", sales_df.shape)
print("\nMissing per column:\n", sales_df.isnull().sum())
print("\nNumeric summary:\n", sales_df[["Cost","Sales","Quantity","Total_Cost","Total_Sales"]].describe())

# Top categories / brands / states
print("\nTop Categories:\n", sales_df["Category"].value_counts().head(10))
print("\nTop Brands:\n", sales_df["Brand"].value_counts().head(10))
if "State" in sales_df.columns:
    print("\nTop States by count:\n", sales_df["State"].value_counts().head(10))






monthly = sales_df.groupby(["Year","Month_Num","Month"])["Total_Sales"].sum().reset_index()
monthly = monthly.sort_values(["Year","Month_Num"])

plt.figure()
sns.lineplot(data=monthly, x="Month_Num", y="Total_Sales", hue="Year", marker="o")
plt.xticks(ticks=range(1,13), labels=[
    "Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"
])
plt.title("Monthly Sales Trend by Year")
plt.xlabel("Month")
plt.ylabel("Total Sales")
plt.tight_layout()
plt.savefig("outputs/monthly_sales_trend.png", bbox_inches="tight")
plt.show()











if "State" in sales_df.columns:
    state_sales = sales_df.groupby("State")["Total_Sales"].sum().sort_values(ascending=False).head(20).reset_index()
    plt.figure()
    sns.barplot(data=state_sales, x="Total_Sales", y="State")
    plt.title("Top 20 States by Total Sales")
    plt.tight_layout()
    plt.savefig("outputs/state_sales_top20.png", bbox_inches="tight")
    plt.show()
else:
    print("No State column to plot")







if "Category" in sales_df.columns:
    cat_sales = sales_df.groupby("Category")["Total_Sales"].sum().reset_index().sort_values("Total_Sales", ascending=False)
    plt.figure()
    sns.barplot(data=cat_sales, x="Total_Sales", y="Category")
    plt.title("Sales by Category")
    plt.tight_layout()
    plt.savefig("outputs/category_sales.png", bbox_inches="tight")
    plt.show()









# possible column names
sup_candidates = ["Assigned Supervisor", "Supervisor", "Salesperson", "Supervisor Name"]
sup_col = next((c for c in sup_candidates if c in sales_df.columns), None)

if sup_col:
    sup_sales = sales_df.groupby(sup_col)["Total_Sales"].sum().reset_index().sort_values("Total_Sales", ascending=False)
    plt.figure(figsize=(10,8))
    sns.barplot(data=sup_sales.head(20), x="Total_Sales", y=sup_col)
    plt.title("Top Supervisors by Sales")
    plt.tight_layout()
    plt.savefig("outputs/supervisor_performance.png", bbox_inches="tight")
    plt.show()
else:
    # try to join with supervisor_df if available
    if "Supervisor" in supervisor_df.columns:
        print("No supervisor column found in sales. You have supervisor sheet — consider mapping.")
    else:
        print("Supervisor info not available")






if "Brand" in sales_df.columns:
    brand_sales = sales_df.groupby("Brand")["Total_Sales"].sum().sort_values(ascending=False).head(10).reset_index()
    plt.figure()
    sns.barplot(data=brand_sales, x="Total_Sales", y="Brand")
    plt.title("Top 10 Brands by Sales")
    plt.tight_layout()
    plt.savefig("outputs/top_brands.png", bbox_inches="tight")
    plt.show()






if {"Total_Sales","Total_Cost"}.issubset(sales_df.columns):
    sales_df["Profit"] = sales_df["Total_Sales"] - sales_df["Total_Cost"]
    profit_cat = sales_df.groupby("Category")["Profit"].sum().reset_index().sort_values("Profit", ascending=False)
    plt.figure()
    sns.barplot(data=profit_cat, x="Profit", y="Category")
    plt.title("Profit by Category")
    plt.tight_layout()
    plt.savefig("outputs/profit_by_category.png", bbox_inches="tight")
    plt.show()
else:
    print("Cannot compute Profit — Total_Sales or Total_Cost missing")







num_cols = ["Cost","Sales","Quantity","Total_Cost","Total_Sales"]
num_present = [c for c in num_cols if c in sales_df.columns]
if len(num_present) >= 2:
    plt.figure(figsize=(8,6))
    sns.heatmap(sales_df[num_present].corr(), annot=True)
    plt.title("Correlation Heatmap")
    plt.tight_layout()
    plt.savefig("outputs/correlation_heatmap.png", bbox_inches="tight")
    plt.show()
else:
    print("Not enough numeric columns for correlation")







# total sales by state
if "State" in sales_df.columns:
    sales_df.groupby("State")["Total_Sales"].sum().sort_values(ascending=False).to_csv("outputs/sales_by_state.csv")

# sales by category
if "Category" in sales_df.columns:
    sales_df.groupby("Category")["Total_Sales"].sum().sort_values(ascending=False).to_csv("outputs/sales_by_category.csv")

# supervisor summary
if "Profit" in sales_df.columns:
    sup_summary = sales_df.groupby(sup_col)[["Total_Sales", "Profit"]]
else:
    sup_summary = sales_df.groupby(sup_col)[["Total_Sales"]]

summary_sup = sales_df.groupby(sup_col).agg(
    total_sales=("Total_Sales","sum"),
    orders=("Order_Number","nunique")
).sort_values("total_sales", ascending=False)


# whole cleaned data
sales_df.to_csv("outputs/sales_data_cleaned.csv", index=False)
print("CSV exports saved in outputs/")

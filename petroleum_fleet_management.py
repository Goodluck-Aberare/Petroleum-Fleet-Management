
# ============================================================
# PETROLEUM TANKER LOGISTICS MANAGEMENT
# PROJECT: PETROLEUM PRODUCT DELIVERY SHORTAGE ANALYSIS
# AUTHOR: GOODLUCK ABERARE
# ============================================================


# STEP 1: IMPORT REQUIRED LIBRARIES

import os
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns


# STEP 2: LOCATE THE PROJECT FOLDER AND EXCEL FILE

project_folder = os.path.dirname(os.path.abspath(__file__))

file_path = os.path.join(
    project_folder,
    "Petroleum_tanker_logistics.xlsx"
)

output_folder = os.path.join(project_folder, "output")

os.makedirs(output_folder, exist_ok=True)


# STEP 3: CHECK AND LOAD THE EXCEL DATASET

print("\n======================================")
print("PETROLEUM DELIVERY SHORTAGE ANALYSIS")
print("======================================")

if not os.path.exists(file_path):
    raise FileNotFoundError(
        f"Excel file not found: {file_path}\n"
        "Place Petroleum_tanker_logistics.xlsx "
        "in the same folder as this Python file."
    )

# Load the actual worksheet name in your workbook
df = pd.read_excel(
    file_path,
    sheet_name="Sheet1"
)

print("\nDataset loaded successfully!")


# STEP 4: CHECK THE DATASET STRUCTURE

print("\n========== DATASET OVERVIEW ==========")

print("\nFirst five records:")
print(df.head())

print("\nDataset shape:")
print(df.shape)

print("\nColumn names:")
print(df.columns.tolist())

print("\nDataset information:")
df.info()

print("\nMissing values:")
print(df.isnull().sum())


# STEP 5: CHECK REQUIRED COLUMNS

required_columns = [
    "Trip ID",
    "Driver ID",
    "Driver Name",
    "Truck ID",
    "Truck Registration Number",
    "Truck Capacity (Litres)",
    "Product",
    "Loading Depot",
    "Destination",
    "Loading Date",
    "Quantity Loaded (Litres)",
    "Quantity Delivered (Litres)",
    "Shortage (Litres)",
    "Distance (km)",
    "Fuel Used (Litres)",
    "Transport Fee (NGN)",
    "Trip Status",
    "Payment Status"
]

missing_columns = [
    column for column in required_columns
    if column not in df.columns
]

if missing_columns:
    raise ValueError(
        f"The following required columns are missing: "
        f"{missing_columns}"
    )


# STEP 6: DATA CLEANING

# Remove duplicate trip records
df = df.drop_duplicates(subset=["Trip ID"]).copy()

# Convert loading date to datetime
df["Loading Date"] = pd.to_datetime(
    df["Loading Date"],
    errors="coerce"
)

# Convert numerical columns to numeric values
numeric_columns = [
    "Truck Capacity (Litres)",
    "Quantity Loaded (Litres)",
    "Quantity Delivered (Litres)",
    "Shortage (Litres)",
    "Distance (km)",
    "Fuel Used (Litres)",
    "Transport Fee (NGN)"
]

for column in numeric_columns:
    df[column] = pd.to_numeric(
        df[column],
        errors="coerce"
    )

# Remove rows without essential trip information
df = df.dropna(
    subset=[
        "Trip ID",
        "Driver Name",
        "Truck Registration Number",
        "Quantity Loaded (Litres)",
        "Quantity Delivered (Litres)",
        "Trip Status"
    ]
).copy()

# Standardize trip status
df["Trip Status"] = (
    df["Trip Status"].astype(str).str.strip()
)

print("\nData cleaning completed.")

print("Number of records after cleaning:", len(df))


# STEP 7: CALCULATE DELIVERY SHORTAGES

# Shortage means quantity loaded minus quantity delivered.
# This is a delivery difference, not proof of product loss.

df["Shortage (Litres)"] = (
    df["Quantity Loaded (Litres)"]
    - df["Quantity Delivered (Litres)"]
)

# Calculate shortage percentage
df["Shortage Percentage"] = (
    df["Shortage (Litres)"]
    / df["Quantity Loaded (Litres)"].where(
        df["Quantity Loaded (Litres)"] > 0
    )
) * 100

print("\nDelivery shortage calculations completed.")


# STEP 8: SEPARATE COMPLETED TRIPS

completed_trips = df[
    df["Trip Status"].str.lower() == "completed"
].copy()

# Exclude records with zero or negative loaded quantities
completed_trips = completed_trips[
    completed_trips["Quantity Loaded (Litres)"] > 0
].copy()

if completed_trips.empty:
    raise ValueError(
        "No completed trips with positive loaded quantities "
        "were found. Check your Trip Status column."
    )

print("\n========== TRIP STATUS SUMMARY ==========")

print(df["Trip Status"].value_counts())


# STEP 9: OVERALL DELIVERY SHORTAGE ANALYSIS

total_trips = len(completed_trips)

total_loaded = completed_trips[
    "Quantity Loaded (Litres)"
].sum()

total_delivered = completed_trips[
    "Quantity Delivered (Litres)"
].sum()

total_shortage = completed_trips[
    "Shortage (Litres)"
].sum()

overall_shortage_percentage = (
    total_shortage / total_loaded * 100
    if total_loaded > 0 else 0
)

print("\n========== OVERALL DELIVERY ANALYSIS ==========")

print(f"Completed trips: {total_trips}")

print(f"Total petroleum loaded: {total_loaded:,.2f} litres")

print(
    f"Total petroleum delivered: "
    f"{total_delivered:,.2f} litres"
)

print(
    f"Total delivery difference: "
    f"{total_shortage:,.2f} litres"
)

print(
    f"Overall shortage percentage: "
    f"{overall_shortage_percentage:.2f}%"
)


# STEP 10: DRIVER PERFORMANCE ANALYSIS

driver_analysis = (
    completed_trips.groupby(
        ["Driver ID", "Driver Name"]
    )
    .agg(
        Total_Trips=("Trip ID", "count"),
        Total_Loaded=(
            "Quantity Loaded (Litres)", "sum"
        ),
        Total_Delivered=(
            "Quantity Delivered (Litres)", "sum"
        ),
        Total_Shortage=(
            "Shortage (Litres)", "sum"
        )
    )
    .reset_index()
)

driver_analysis["Shortage_Percentage"] = (
    driver_analysis["Total_Shortage"]
    / driver_analysis["Total_Loaded"].where(
        driver_analysis["Total_Loaded"] > 0
    )
) * 100

driver_analysis = driver_analysis.sort_values(
    by="Total_Shortage",
    ascending=False
)

print("\n========== DRIVER SHORTAGE ANALYSIS ==========")

print(driver_analysis.head(10).to_string(index=False))


# STEP 11: TRUCK PERFORMANCE ANALYSIS

truck_analysis = (
    completed_trips.groupby(
        ["Truck ID", "Truck Registration Number"]
    )
    .agg(
        Total_Trips=("Trip ID", "count"),
        Total_Loaded=(
            "Quantity Loaded (Litres)", "sum"
        ),
        Total_Delivered=(
            "Quantity Delivered (Litres)", "sum"
        ),
        Total_Shortage=(
            "Shortage (Litres)", "sum"
        )
    )
    .reset_index()
)

truck_analysis["Shortage_Percentage"] = (
    truck_analysis["Total_Shortage"]
    / truck_analysis["Total_Loaded"].where(
        truck_analysis["Total_Loaded"] > 0
    )
) * 100

truck_analysis = truck_analysis.sort_values(
    by="Total_Shortage",
    ascending=False
)

print("\n========== TRUCK SHORTAGE ANALYSIS ==========")

print(truck_analysis.head(10).to_string(index=False))


# STEP 12: DEPOT SHORTAGE ANALYSIS

depot_analysis = (
    completed_trips.groupby("Loading Depot")
    .agg(
        Total_Trips=("Trip ID", "count"),
        Total_Loaded=(
            "Quantity Loaded (Litres)", "sum"
        ),
        Total_Delivered=(
            "Quantity Delivered (Litres)", "sum"
        ),
        Total_Shortage=(
            "Shortage (Litres)", "sum"
        )
    )
    .reset_index()
)

depot_analysis = depot_analysis.sort_values(
    by="Total_Shortage",
    ascending=False
)

print("\n========== DEPOT SHORTAGE ANALYSIS ==========")

print(depot_analysis.to_string(index=False))


# STEP 13: DESTINATION SHORTAGE ANALYSIS

destination_analysis = (
    completed_trips.groupby("Destination")
    .agg(
        Total_Trips=("Trip ID", "count"),
        Total_Loaded=(
            "Quantity Loaded (Litres)", "sum"
        ),
        Total_Delivered=(
            "Quantity Delivered (Litres)", "sum"
        ),
        Total_Shortage=(
            "Shortage (Litres)", "sum"
        )
    )
    .reset_index()
)

destination_analysis = destination_analysis.sort_values(
    by="Total_Shortage",
    ascending=False
)

print("\n========== DESTINATION SHORTAGE ANALYSIS ==========")

print(destination_analysis.to_string(index=False))


# STEP 14: PREPARE VISUALIZATION SETTINGS

sns.set_theme(style="whitegrid")


# CHART 1: TOTAL LOADED VS TOTAL DELIVERED

plt.figure(figsize=(9, 6))

plt.bar(
    ["Total Loaded", "Total Delivered"],
    [total_loaded, total_delivered]
)

plt.title("Petroleum Loaded vs Petroleum Delivered")
plt.xlabel("Delivery Category")
plt.ylabel("Quantity in Litres")

plt.tight_layout()

plt.savefig(
    os.path.join(output_folder, "loaded_vs_delivered.png")
)

plt.close()


# CHART 2: TOP 10 DRIVERS BY DELIVERY SHORTAGE

top_drivers = driver_analysis.head(10)

plt.figure(figsize=(11, 6))

sns.barplot(
    data=top_drivers,
    x="Total_Shortage",
    y="Driver Name"
)

plt.title("Top 10 Drivers by Delivery Difference")
plt.xlabel("Shortage in Litres")
plt.ylabel("Driver Name")

plt.tight_layout()

plt.savefig(
    os.path.join(output_folder, "driver_shortage.png")
)

plt.close()


# CHART 3: TOP 10 TRUCKS BY DELIVERY SHORTAGE

top_trucks = truck_analysis.head(10)

plt.figure(figsize=(11, 6))

sns.barplot(
    data=top_trucks,
    x="Total_Shortage",
    y="Truck Registration Number"
)

plt.title("Top 10 Trucks by Delivery Difference")
plt.xlabel("Shortage in Litres")
plt.ylabel("Truck Registration Number")

plt.tight_layout()

plt.savefig(
    os.path.join(output_folder, "truck_shortage.png")
)

plt.close()


# CHART 4: DELIVERY SHORTAGE BY DEPOT

plt.figure(figsize=(11, 6))

sns.barplot(
    data=depot_analysis,
    x="Total_Shortage",
    y="Loading Depot"
)

plt.title("Delivery Difference by Loading Depot")
plt.xlabel("Shortage in Litres")
plt.ylabel("Loading Depot")

plt.tight_layout()

plt.savefig(
    os.path.join(output_folder, "depot_shortage.png")
)

plt.close()


# CHART 5: DELIVERY SHORTAGE BY DESTINATION

plt.figure(figsize=(11, 6))

sns.barplot(
    data=destination_analysis,
    x="Total_Shortage",
    y="Destination"
)

plt.title("Delivery Difference by Destination")
plt.xlabel("Shortage in Litres")
plt.ylabel("Destination")

plt.tight_layout()

plt.savefig(
    os.path.join(output_folder, "destination_shortage.png")
)

plt.close()


# STEP 15: CREATE A SUMMARY REPORT

summary = pd.DataFrame({
    "Metric": [
        "Total Completed Trips",
        "Total Quantity Loaded (Litres)",
        "Total Quantity Delivered (Litres)",
        "Total Delivery Difference (Litres)",
        "Overall Shortage Percentage"
    ],
    "Value": [
        total_trips,
        total_loaded,
        total_delivered,
        total_shortage,
        round(overall_shortage_percentage, 2)
    ]
})


# STEP 16: EXPORT ALL ANALYSIS RESULTS TO EXCEL

report_path = os.path.join(
    output_folder,
    "Petroleum_Delivery_Shortage_Analysis.xlsx"
)

with pd.ExcelWriter(
    report_path,
    engine="openpyxl"
) as writer:

    summary.to_excel(
        writer,
        sheet_name="Summary",
        index=False
    )

    completed_trips.to_excel(
        writer,
        sheet_name="Completed Trips",
        index=False
    )

    driver_analysis.to_excel(
        writer,
        sheet_name="Driver Analysis",
        index=False
    )

    truck_analysis.to_excel(
        writer,
        sheet_name="Truck Analysis",
        index=False
    )

    depot_analysis.to_excel(
        writer,
        sheet_name="Depot Analysis",
        index=False
    )

    destination_analysis.to_excel(
        writer,
        sheet_name="Destination Analysis",
        index=False
    )

    df["Trip Status"].value_counts().rename_axis(
        "Trip Status"
    ).reset_index(
        name="Number of Trips"
    ).to_excel(
        writer,
        sheet_name="Trip Status",
        index=False
    )


# STEP 17: DISPLAY COMPLETION MESSAGE

print("\n======================================")
print("PROJECT COMPLETED SUCCESSFULLY!")
print("======================================")

print("\nFive analysis charts have been saved.")

print("\nExcel analysis report:")
print(report_path)

print("\nOutput folder:")
print(output_folder)

print("\nThank you for using the Petroleum")
print("Tanker Logistics Analysis Program.")

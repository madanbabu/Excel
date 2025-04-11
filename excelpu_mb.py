import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
from datetime import datetime

# Set plotting style
sns.set(style="whitegrid")

def parse_document_data():
    """
    Parse the provided document data into a structured DataFrame.
    In a real scenario, replace this with pd.read_excel('file.xlsx').
    """
    # Monthly data from the document
    monthly_data = [
        {"Years": 2023, "Months": "Jan", "Product Line": "Classic Cars", "Value": 41191.78},
        {"Years": 2023, "Months": "Jan", "Product Line": "Trains", "Value": 4933.55},
        {"Years": 2023, "Months": "Jan", "Product Line": "Vintage Cars", "Value": 46826.84},
        {"Years": 2023, "Months": "Feb", "Product Line": "Classic Cars", "Value": 20464.41},
        {"Years": 2023, "Months": "Feb", "Product Line": "Motorcycles", "Value": 25783.76},
        {"Years": 2023, "Months": "Feb", "Product Line": "Planes", "Value": 39205.31},
        {"Years": 2023, "Months": "Feb", "Product Line": "Ships", "Value": 27050.38},
        {"Years": 2023, "Months": "Feb", "Product Line": "Trains", "Value": 4330.10},
        {"Years": 2023, "Months": "Feb", "Product Line": "Vintage Cars", "Value": 24002.23},
        {"Years": 2023, "Months": "Mar", "Product Line": "Classic Cars", "Value": 105026.68},
        {"Years": 2023, "Months": "Mar", "Product Line": "Motorcycles", "Value": 12639.15},
        {"Years": 2023, "Months": "Mar", "Product Line": "Vintage Cars", "Value": 46931.01},
        {"Years": 2023, "Months": "Apr", "Product Line": "Classic Cars", "Value": 59873.60},
        {"Years": 2023, "Months": "Apr", "Product Line": "Motorcycles", "Value": 23475.59},
        {"Years": 2023, "Months": "Apr", "Product Line": "Planes", "Value": 36563.34},
        {"Years": 2023, "Months": "Apr", "Product Line": "Ships", "Value": 27399.45},
        {"Years": 2023, "Months": "Apr", "Product Line": "Trains", "Value": 4756.47},
        {"Years": 2023, "Months": "Apr", "Product Line": "Vintage Cars", "Value": 20750.34},
        {"Years": 2023, "Months": "May", "Product Line": "Classic Cars", "Value": 98179.48},
        {"Years": 2023, "Months": "May", "Product Line": "Motorcycles", "Value": 22097.32},
        {"Years": 2023, "Months": "May", "Product Line": "Vintage Cars", "Value": 47033.58},
        {"Years": 2023, "Months": "Jun", "Product Line": "Classic Cars", "Value": 50256.79},
        {"Years": 2023, "Months": "Jun", "Product Line": "Motorcycles", "Value": 2642.01},
        {"Years": 2023, "Months": "Jun", "Product Line": "Planes", "Value": 32115.65},
        {"Years": 2023, "Months": "Jun", "Product Line": "Ships", "Value": 29044.82},
        {"Years": 2023, "Months": "Jun", "Product Line": "Trains", "Value": 10071.06},
        {"Years": 2023, "Months": "Jun", "Product Line": "Vintage Cars", "Value": 26582.51},
        {"Years": 2023, "Months": "Jul", "Product Line": "Classic Cars", "Value": 94055.58},
        {"Years": 2023, "Months": "Jul", "Product Line": "Motorcycles", "Value": 37924.23},
        {"Years": 2023, "Months": "Jul", "Product Line": "Vintage Cars", "Value": 29652.09},
        {"Years": 2023, "Months": "Aug", "Product Line": "Classic Cars", "Value": 48406.61},
        {"Years": 2023, "Months": "Aug", "Product Line": "Motorcycles", "Value": 44164.91},
        {"Years": 2023, "Months": "Aug", "Product Line": "Planes", "Value": 33938.33},
        {"Years": 2023, "Months": "Aug", "Product Line": "Ships", "Value": 25607.08},
        {"Years": 2023, "Months": "Aug", "Product Line": "Trains", "Value": 7374.94},
        {"Years": 2023, "Months": "Aug", "Product Line": "Vintage Cars", "Value": 23285.46},
        {"Years": 2023, "Months": "Sep", "Product Line": "Classic Cars", "Value": 137666.87},
        {"Years": 2023, "Months": "Sep", "Product Line": "Motorcycles", "Value": 3155.58},
        {"Years": 2023, "Months": "Sep", "Product Line": "Ships", "Value": 20564.19},
        {"Years": 2023, "Months": "Sep", "Product Line": "Trains", "Value": 6034.52},
        {"Years": 2023, "Months": "Sep", "Product Line": "Vintage Cars", "Value": 53045.90},
        {"Years": 2023, "Months": "Oct", "Product Line": "Classic Cars", "Value": 241145.43},
        {"Years": 2023, "Months": "Oct", "Product Line": "Motorcycles", "Value": 64235.65},
        {"Years": 2023, "Months": "Oct", "Product Line": "Planes", "Value": 69180.74},
        {"Years": 2023, "Months": "Oct", "Product Line": "Ships", "Value": 35980.37},
        {"Years": 2023, "Months": "Oct", "Product Line": "Trains", "Value": 10233.50},
        {"Years": 2023, "Months": "Oct", "Product Line": "Vintage Cars", "Value": 100603.25},
        {"Years": 2023, "Months": "Nov", "Product Line": "Classic Cars", "Value": 452924.37},
        {"Years": 2023, "Months": "Nov", "Product Line": "Motorcycles", "Value": 109345.50},
        {"Years": 2023, "Months": "Nov", "Product Line": "Planes", "Value": 54133.27},
        {"Years": 2023, "Months": "Nov", "Product Line": "Ships", "Value": 79174.80},
        {"Years": 2023, "Months": "Nov", "Product Line": "Trains", "Value": 22523.40},
        {"Years": 2023, "Months": "Nov", "Product Line": "Vintage Cars", "Value": 184673.40},
        {"Years": 2023, "Months": "Dec", "Product Line": "Classic Cars", "Value": 135593.69},
        {"Years": 2023, "Months": "Dec", "Product Line": "Motorcycles", "Value": 25431.88},
        {"Years": 2023, "Months": "Dec", "Product Line": "Planes", "Value": 7120.96},
        {"Years": 2023, "Months": "Dec", "Product Line": "Trains", "Value": 2544.75},
        {"Years": 2023, "Months": "Dec", "Product Line": "Vintage Cars", "Value": 47601.15},
        {"Years": 2024, "Months": "Jan", "Product Line": "Classic Cars", "Value": 122791.55},
        {"Years": 2024, "Months": "Jan", "Product Line": "Motorcycles", "Value": 41200.52},
        {"Years": 2024, "Months": "Jan", "Product Line": "Planes", "Value": 35453.54},
        {"Years": 2024, "Months": "Jan", "Product Line": "Ships", "Value": 29479.55},
        {"Years": 2024, "Months": "Jan", "Product Line": "Trains", "Value": 7582.86},
        {"Years": 2024, "Months": "Jan", "Product Line": "Vintage Cars", "Value": 74917.40},
        {"Years": 2024, "Months": "Feb", "Product Line": "Classic Cars", "Value": 133034.82},
        {"Years": 2024, "Months": "Feb", "Product Line": "Motorcycles", "Value": 49066.50},
        {"Years": 2024, "Months": "Feb", "Product Line": "Planes", "Value": 37659.93},
        {"Years": 2024, "Months": "Feb", "Product Line": "Ships", "Value": 30611.52},
        {"Years": 2024, "Months": "Feb", "Product Line": "Trains", "Value": 4968.01},
        {"Years": 2024, "Months": "Feb", "Product Line": "Vintage Cars", "Value": 25041.84},
        {"Years": 2024, "Months": "Mar", "Product Line": "Classic Cars", "Value": 81144.00},
        {"Years": 2024, "Months": "Mar", "Product Line": "Ships", "Value": 21576.79},
        {"Years": 2024, "Months": "Mar", "Product Line": "Trains", "Value": 9641.82},
        {"Years": 2024, "Months": "Mar", "Product Line": "Vintage Cars", "Value": 55981.22},
        {"Years": 2024, "Months": "Apr", "Product Line": "Classic Cars", "Value": 91815.89},
        {"Years": 2024, "Months": "Apr", "Product Line": "Motorcycles", "Value": 36269.07},
        {"Years": 2024, "Months": "Apr", "Product Line": "Planes", "Value": 23032.84},
        {"Years": 2024, "Months": "Apr", "Product Line": "Vintage Cars", "Value": 55030.32},
        {"Years": 2024, "Months": "May", "Product Line": "Classic Cars", "Value": 83382.16},
        {"Years": 2024, "Months": "May", "Product Line": "Motorcycles", "Value": 46848.95},
        {"Years": 2024, "Months": "May", "Product Line": "Planes", "Value": 40656.20},
        {"Years": 2024, "Months": "May", "Product Line": "Ships", "Value": 29421.29},
        {"Years": 2024, "Months": "May", "Product Line": "Trains", "Value": 10199.47},
        {"Years": 2024, "Months": "May", "Product Line": "Vintage Cars", "Value": 26749.18},
        {"Years": 2024, "Months": "Jun", "Product Line": "Classic Cars", "Value": 87163.19},
        {"Years": 2024, "Months": "Jun", "Product Line": "Motorcycles", "Value": 47237.41},
        {"Years": 2024, "Months": "Jun", "Product Line": "Planes", "Value": 36650.28},
        {"Years": 2024, "Months": "Jun", "Product Line": "Ships", "Value": 24229.71},
        {"Years": 2024, "Months": "Jun", "Product Line": "Vintage Cars", "Value": 47985.84},
        {"Years": 2024, "Months": "Jul", "Product Line": "Classic Cars", "Value": 151237.37},
        {"Years": 2024, "Months": "Jul", "Product Line": "Motorcycles", "Value": 22774.00},
        {"Years": 2024, "Months": "Jul", "Product Line": "Planes", "Value": 40915.32},
        {"Years": 2024, "Months": "Jul", "Product Line": "Ships", "Value": 22532.62},
        {"Years": 2024, "Months": "Jul", "Product Line": "Trains", "Value": 10802.20},
        {"Years": 2024, "Months": "Jul", "Product Line": "Vintage Cars", "Value": 40712.49},
        {"Years": 2024, "Months": "Aug", "Product Line": "Classic Cars", "Value": 216895.82},
        {"Years": 2024, "Months": "Aug", "Product Line": "Motorcycles", "Value": 62704.93},
        {"Years": 2024, "Months": "Aug", "Product Line": "Planes", "Value": 33413.68},
        {"Years": 2024, "Months": "Aug", "Product Line": "Ships", "Value": 26643.29},
        {"Years": 2024, "Months": "Aug", "Product Line": "Trains", "Value": 8793.18},
        {"Years": 2024, "Months": "Aug", "Product Line": "Vintage Cars", "Value": 78139.05},
        {"Years": 2024, "Months": "Sep", "Product Line": "Classic Cars", "Value": 97006.21},
        {"Years": 2024, "Months": "Sep", "Product Line": "Motorcycles", "Value": 42471.05},
        {"Years": 2024, "Months": "Sep", "Product Line": "Planes", "Value": 34785.62},
        {"Years": 2024, "Months": "Sep", "Product Line": "Ships", "Value": 25151.04},
        {"Years": 2024, "Months": "Sep", "Product Line": "Trains", "Value": 6445.32},
        {"Years": 2024, "Months": "Sep", "Product Line": "Vintage Cars", "Value": 73100.32},
        {"Years": 2024, "Months": "Oct", "Product Line": "Classic Cars", "Value": 223856.81},
        {"Years": 2024, "Months": "Oct", "Product Line": "Motorcycles", "Value": 39413.96},
        {"Years": 2024, "Months": "Oct", "Product Line": "Planes", "Value": 37596.91},
        {"Years": 2024, "Months": "Oct", "Product Line": "Ships", "Value": 43811.37},
        {"Years": 2024, "Months": "Oct", "Product Line": "Trains", "Value": 15183.37},
        {"Years": 2024, "Months": "Oct", "Product Line": "Vintage Cars", "Value": 116160.29},
        {"Years": 2024, "Months": "Nov", "Product Line": "Classic Cars", "Value": 372231.89},
        {"Years": 2024, "Months": "Nov", "Product Line": "Motorcycles", "Value": 151711.86},
        {"Years": 2024, "Months": "Nov", "Product Line": "Planes", "Value": 121130.70},
        {"Years": 2024, "Months": "Nov", "Product Line": "Ships", "Value": 63900.85},
        {"Years": 2024, "Months": "Nov", "Product Line": "Trains", "Value": 22271.23},
        {"Years": 2024, "Months": "Nov", "Product Line": "Vintage Cars", "Value": 233990.34},
        {"Years": 2024, "Months": "Dec", "Product Line": "Classic Cars", "Value": 101697.38},
        {"Years": 2024, "Months": "Dec", "Product Line": "Motorcycles", "Value": 20846.98},
        {"Years": 2024, "Months": "Dec", "Product Line": "Planes", "Value": 61376.78},
        {"Years": 2024, "Months": "Dec", "Product Line": "Ships", "Value": 24079.94},
        {"Years": 2024, "Months": "Dec", "Product Line": "Trains", "Value": 20636.39},
        {"Years": 2024, "Months": "Dec", "Product Line": "Vintage Cars", "Value": 83615.48},
    ]
    
    # Yearly data from the document
    yearly_data = [
        {"Years": 2023, "Product Line": "Classic Cars", "Value": 130000},
        {"Years": 2023, "Product Line": "Motorcycles", "Value": 400000},
        {"Years": 2023, "Product Line": "Planes", "Value": 300000},
        {"Years": 2023, "Product Line": "Ships", "Value": 250000},
        {"Years": 2023, "Product Line": "Trains", "Value": 100000},
        {"Years": 2023, "Product Line": "Vintage Cars", "Value": 600000},
        {"Years": 2024, "Product Line": "Classic Cars", "Value": 150000},
        {"Years": 2024, "Product Line": "Motorcycles", "Value": 500000},
        {"Years": 2024, "Product Line": "Planes", "Value": 450000},
        {"Years": 2024, "Product Line": "Ships", "Value": 350000},
        {"Years": 2024, "Product Line": "Trains", "Value": 100000},
        {"Years": 2024, "Product Line": "Vintage Cars", "Value": 1000000},
    ]
    
    return pd.DataFrame(monthly_data), pd.DataFrame(yearly_data)

def clean_data(monthly_df, yearly_df):
    """Clean and validate the DataFrames."""
    try:
        # Convert 'Value' to numeric
        monthly_df['Value'] = pd.to_numeric(monthly_df['Value'], errors='coerce')
        yearly_df['Value'] = pd.to_numeric(yearly_df['Value'], errors='coerce')
        
        # Convert 'Years' to integer
        monthly_df['Years'] = pd.to_numeric(monthly_df['Years'], errors='coerce').astype('Int64')
        yearly_df['Years'] = pd.to_numeric(yearly_df['Years'], errors='coerce').astype('Int64')
        
        # Check for missing values
        if monthly_df.isnull().any().any():
            print("Warning: Missing values detected in monthly data")
            monthly_df = monthly_df.fillna(0)
        if yearly_df.isnull().any().any():
            print("Warning: Missing values detected in yearly data")
            yearly_df = yearly_df.fillna(0)
        
        # Define month order
        month_order = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
        monthly_df['Months'] = pd.Categorical(monthly_df['Months'], categories=month_order, ordered=True)
        
        return monthly_df, yearly_df
    except Exception as e:
        print(f"Error cleaning data: {e}")
        return None, None

def analyze_totals(monthly_df, yearly_df):
    """Compare monthly sums with provided yearly totals."""
    monthly_sums = monthly_df.groupby(['Years', 'Product Line'])['Value'].sum().unstack()
    yearly_totals = yearly_df.pivot(index='Years', columns='Product Line', values='Value')
    
    print("\nMonthly Data Totals by Year and Product Line:")
    print(monthly_sums)
    
    print("\nProvided Yearly Totals:")
    print(yearly_totals)
    
    # Check for discrepancies
    print("\nDiscrepancies (Monthly Sums vs. Yearly Totals):")
    for year in monthly_sums.index:
        for product in monthly_sums.columns:
            monthly_val = monthly_sums.loc[year, product] if pd.notnull(monthly_sums.loc[year, product]) else 0
            yearly_val = yearly_totals.loc[year, product] if pd.notnull(yearly_totals.loc[year, product]) else 0
            diff = monthly_val - yearly_val
            if abs(diff) > 1e-2:  # Allow small floating-point differences
                print(f"Year {year}, {product}: Monthly Sum = {monthly_val:.2f}, Yearly Total = {yearly_val:.2f}, Difference = {diff:.2f}")

def plot_monthly_trends(monthly_df):
    """Plot monthly trends for each product line."""
    pivot_df = monthly_df.pivot_table(index=['Years', 'Months'], columns='Product Line', values='Value', aggfunc='sum').reset_index()
    
    plt.figure(figsize=(14, 8))
    for product in pivot_df.columns[2:]:  # Skip Years and Months
        sns.lineplot(data=pivot_df, x='Months', y=product, hue='Years', marker='o', label=product)
    
    plt.title('Monthly Sales Trends by Product Line (2023 vs. 2024)')
    plt.xlabel('Month')
    plt.ylabel('Value')
    plt.legend(title='Product Line', bbox_to_anchor=(1.05, 1), loc='upper left')
    plt.tight_layout()
    plt.show()

def plot_stacked_bar(monthly_df):
    """Plot stacked bar chart of product line contributions by month."""
    pivot_df = monthly_df.pivot_table(index=['Years', 'Months'], columns='Product Line', values='Value', aggfunc='sum').fillna(0)
    
    fig, axes = plt.subplots(2, 1, figsize=(12, 10), sharex=True)
    for idx, year in enumerate([2023, 2024]):
        year_data = pivot_df.loc[year]
        year_data.plot(kind='bar', stacked=True, ax=axes[idx])
        axes[idx].set_title(f'Product Line Contributions - {year}')
        axes[idx].set_ylabel('Value')
        axes[idx].legend(title='Product Line', bbox_to_anchor=(1.05, 1), loc='upper left')
    
    plt.xlabel('Month')
    plt.tight_layout()
    plt.show()

def analyze_performance(monthly_df):
    """Analyze top performers, growth, and outliers."""
    # Total value by product line
    total_by_product = monthly_df.groupby('Product Line')['Value'].sum().sort_values(ascending=False)
    print("\nTotal Value by Product Line (Monthly Data):")
    print(total_by_product)
    
    # Year-over-year growth
    yearly_totals = monthly_df.groupby('Years')['Value'].sum()
    growth = ((yearly_totals[2024] - yearly_totals[2023]) / yearly_totals[2023]) * 100
    print(f"\nYear-over-Year Growth (2023 to 2024): {growth:.2f}%")
    
    # Monthly averages by product line
    monthly_avg = monthly_df.groupby(['Product Line', 'Months'])['Value'].mean().unstack()
    print("\nAverage Monthly Value by Product Line:")
    print(monthly_avg)
    
    # Detect outliers (values > Q3 + 1.5*IQR or < Q1 - 1.5*IQR)
    print("\nPotential Outliers (by Product Line):")
    for product in monthly_df['Product Line'].unique():
        product_data = monthly_df[monthly_df['Product Line'] == product]['Value']
        q1 = product_data.quantile(0.25)
        q3 = product_data.quantile(0.75)
        iqr = q3 - q1
        outliers = product_data[(product_data < (q1 - 1.5 * iqr)) | (product_data > (q3 + 1.5 * iqr))]
        if not outliers.empty:
            print(f"{product}:")
            print(monthly_df.loc[outliers.index, ['Years', 'Months', 'Value']])

def main():
    """Main function to run the analysis."""
    # Load and clean data
    monthly_df, yearly_df = parse_document_data()
    monthly_df, yearly_df = clean_data(monthly_df, yearly_df)
    
    if monthly_df is None or yearly_df is None:
        print("Failed to process data. Exiting.")
        return
    
    # Perform analyses
    analyze_totals(monthly_df, yearly_df)
    analyze_performance(monthly_df)
    
    # Generate visualizations
    plot_monthly_trends(monthly_df)
    plot_stacked_bar(monthly_df)

if __name__ == "__main__":
    main()

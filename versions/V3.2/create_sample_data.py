import pandas as pd
import random
from datetime import datetime, timedelta

# Generate sample sales data
random.seed(42)

# Create date range for 6 months
start_date = datetime(2025, 1, 1)
dates = [start_date + timedelta(days=x) for x in range(180)]

# Product categories and names
categories = ['Electronics', 'Clothing', 'Food & Beverage', 'Home & Garden', 'Sports']
products = {
    'Electronics': ['Laptop', 'Smartphone', 'Headphones', 'Tablet', 'Monitor'],
    'Clothing': ['T-Shirt', 'Jeans', 'Jacket', 'Sneakers', 'Dress'],
    'Food & Beverage': ['Coffee', 'Tea', 'Snacks', 'Water', 'Juice'],
    'Home & Garden': ['Plant', 'Lamp', 'Cushion', 'Curtain', 'Rug'],
    'Sports': ['Yoga Mat', 'Dumbbells', 'Basketball', 'Running Shoes', 'Water Bottle']
}

regions = ['North', 'South', 'East', 'West', 'Central']
sales_reps = ['Alice Chen', 'Bob Wang', 'Carol Li', 'David Zhang', 'Emma Liu']

# Generate 500 rows of sales data
data = []
for _ in range(500):
    category = random.choice(categories)
    product = random.choice(products[category])
    region = random.choice(regions)
    sales_rep = random.choice(sales_reps)
    date = random.choice(dates)
    quantity = random.randint(1, 20)
    
    # Set base prices for categories
    base_prices = {
        'Electronics': random.randint(200, 1500),
        'Clothing': random.randint(20, 200),
        'Food & Beverage': random.randint(2, 30),
        'Home & Garden': random.randint(15, 150),
        'Sports': random.randint(10, 100)
    }
    
    unit_price = base_prices[category]
    total_sales = unit_price * quantity
    cost = unit_price * 0.6  # 60% cost ratio
    profit = total_sales - (cost * quantity)
    
    data.append({
        'Date': date,
        'Region': region,
        'Sales Rep': sales_rep,
        'Category': category,
        'Product': product,
        'Quantity': quantity,
        'Unit Price': unit_price,
        'Total Sales': total_sales,
        'Profit': profit
    })

# Create DataFrame
df = pd.DataFrame(data)
df = df.sort_values('Date')

# Save to Excel with multiple sheets
with pd.ExcelWriter('sample_sales_data.xlsx', engine='openpyxl') as writer:
    # Main data sheet
    df.to_excel(writer, sheet_name='Sales Data', index=False)
    
    # Summary by category
    category_summary = df.groupby('Category').agg({
        'Total Sales': 'sum',
        'Profit': 'sum',
        'Quantity': 'sum'
    }).round(2)
    category_summary.to_excel(writer, sheet_name='Category Summary')
    
    # Summary by region
    region_summary = df.groupby('Region').agg({
        'Total Sales': 'sum',
        'Profit': 'sum',
        'Quantity': 'sum'
    }).round(2)
    region_summary.to_excel(writer, sheet_name='Region Summary')
    
    # Monthly trend
    df['Month'] = df['Date'].dt.to_period('M')
    monthly_summary = df.groupby('Month').agg({
        'Total Sales': 'sum',
        'Profit': 'sum',
        'Quantity': 'sum'
    }).round(2)
    monthly_summary.to_excel(writer, sheet_name='Monthly Trend')

print("✓ Sample data created: sample_sales_data.xlsx")
print(f"  - {len(df)} rows of sales transactions")
print(f"  - Date range: {df['Date'].min().date()} to {df['Date'].max().date()}")
print(f"  - Total sales: ${df['Total Sales'].sum():,.2f}")
print(f"  - 4 sheets: Sales Data, Category Summary, Region Summary, Monthly Trend")

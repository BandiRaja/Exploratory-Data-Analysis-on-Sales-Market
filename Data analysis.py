import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Load Excel file
df = pd.read_excel("C:\\Users\\ACER\\OneDrive\\Desktop\\AmazonSales_RawData.xlsx", sheet_name='Walmart')

# Convert order date and create YearMonth
df['Order Date'] = pd.to_datetime(df['Order Date'])
df['YearMonth'] = df['Order Date'].dt.to_period('M').astype(str)

# Set seaborn style
sns.set(style="whitegrid")

# 1. Monthly Sales & Profit Trend
monthly = df.groupby('YearMonth')[['Sales', 'Profit']].sum().reset_index()
plt.figure(figsize=(14, 6))
sns.lineplot(x='YearMonth', y='Sales', data=monthly, label='Sales', marker='o')
sns.lineplot(x='YearMonth', y='Profit', data=monthly, label='Profit', marker='o')
plt.xticks(rotation=45)
plt.title('Monthly Sales and Profit Trend')
plt.xlabel('Month')
plt.ylabel('Amount ($)')
plt.legend()
plt.tight_layout()
plt.show()

# 2. Category-wise Sales
plt.figure(figsize=(10, 5))
cat_sales = df.groupby('Category')['Sales'].sum().sort_values(ascending=False)
sns.barplot(x=cat_sales.values, y=cat_sales.index, palette='Blues_r')
plt.title('Total Sales by Category')
plt.xlabel('Sales')
plt.ylabel('Category')
plt.tight_layout()
plt.show()

# 3. Top 10 Products by Sales
plt.figure(figsize=(10, 6))
top_products = df.groupby('Product Name')['Sales'].sum().sort_values(ascending=False).head(10)
sns.barplot(x=top_products.values, y=top_products.index, palette='Greens_r')
plt.title('Top 10 Products by Sales')
plt.xlabel('Sales')
plt.ylabel('Product Name')
plt.tight_layout()
plt.show()

# 4. Top 10 States by Sales
plt.figure(figsize=(12, 6))
top_states = df.groupby('State')['Sales'].sum().sort_values(ascending=False).head(10)
sns.barplot(x=top_states.values, y=top_states.index, palette='Oranges_r')
plt.title('Top 10 States by Sales')
plt.xlabel('Sales')
plt.ylabel('State')
plt.tight_layout()
plt.show()

# 5. Correlation Heatmap
plt.figure(figsize=(6, 4))
sns.heatmap(df[['Sales', 'Profit', 'Quantity']].corr(), annot=True, cmap='coolwarm')
plt.title('Correlation between Sales, Profit, and Quantity')
plt.tight_layout()
plt.show()

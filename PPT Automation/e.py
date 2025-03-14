import pandas as pd

# Sample data for the Excel files
data1 = {'Name': ['John', 'Alice', 'Bob'],
         'Age': [28, 24, 35],
         'City': ['New York', 'Los Angeles', 'Chicago']}

data2 = {'Product': ['Laptop', 'Mouse', 'Keyboard'],
         'Price': [1000, 50, 75],
         'Stock': [10, 50, 30]}

# Convert the data into DataFrames
df1 = pd.DataFrame(data1)
df2 = pd.DataFrame(data2)

# Save to two separate Excel files
df1.to_csv('People.csv', index=False)
df2.to_csv('Products.csv', index=False)

print("Two Excel files created: 'People.xlsx' and 'Products.xlsx'")

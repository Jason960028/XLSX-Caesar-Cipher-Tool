import pandas as pd

# Define sample data for the Excel file.
data = {
    "Name": ["Alice", "Bob", "Charlie"],
    "Age": [25, 30, 35],
    "Email": ["alice@example.com", "bob@example.com", "charlie@example.com"],
    "Gender": ["M","F","M"]
}

# Create a DataFrame using the sample data.
df = pd.DataFrame(data)

# Save the DataFrame to an Excel file without the index.
df.to_excel("sample.xlsx", index=False)

print("Sample Excel file 'input.xlsx' has been created.")

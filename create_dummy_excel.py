
import pandas as pd

# Create dummy data
data = {
    'Date': ['2026-01-01', '2026-01-02', '2026-01-03'],
    'Sample ID': ['S001', 'S002', 'S003'],
    'Value': [10.5, 20.1, 15.3]
}

df = pd.DataFrame(data)
df.to_excel('test_data.xlsx', index=False)
print("Dummy excel created.")

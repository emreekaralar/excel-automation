import pandas as pd
import plotly.express as px
import openpyxl
import warnings

warnings.filterwarnings("ignore", category=UserWarning, module='openpyxl')


# Load the Excel file
path = 'MSD Timeline Template - Stay vs Go Renewal New Acquisition v.2.xlsm'
df = pd.read_excel(path, sheet_name=None)

# Assuming the data is in the first sheet
sheet_name = list(df.keys())[1]
data = df[sheet_name]

# Manually identified indices of colored rows
colored_indices = [17, 21, 24, 25, 28, 29, 32, 39, 42, 43, 44, 47, 51, 52,55, 58, 62, 65, 66, 67, 71, 75, 78, 82, 89, 90, 91, 92, 97, 98, 99, 100, 101, 102, 103, 108, 109  ]  # Replace with actual indices

# Select only the colored rows
filtered_data = data.iloc[colored_indices]

# Filter the data to include only the relevant columns (Activity, Start Date, End Date)
filtered_data = filtered_data[['Activity', 'Start Date', 'End Date']]

# Drop rows with missing values in these columns
filtered_data.dropna(subset=['Activity', 'Start Date', 'End Date'], inplace=True)

# Create Gantt chart
fig = px.timeline(filtered_data, x_start='Start Date', x_end='End Date', y='Activity')

# Update layout
fig.update_layout(title='Gantt Chart', xaxis_title='Date', yaxis_title='Activity')

# Lock the rows
fig.update_yaxes(fixedrange=True)

# Show plot
fig.show()

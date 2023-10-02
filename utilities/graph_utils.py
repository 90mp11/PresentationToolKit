import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

def add_jitter(series, jitter_amount=0.1):
    return series + (np.random.rand(len(series)) - 0.5) * jitter_amount

# Read the CSV file
project_df = pd.read_csv('./raw/PROJECT.csv')

# Map 'Estimated Effort' and 'Estimated Impact' to numerical values
effort_mapping = {'Hours': 1, 'Days': 2, 'Weeks': 3, 'Months': 4, 'Quarters': 5}
impact_mapping = {'Very Low': 1, 'Low': 2, 'Medium': 3, 'High': 4, 'Very High': 5}
project_df['Numerical Effort'] = project_df['Estimated Effort'].map(effort_mapping)
project_df['Numerical Impact'] = project_df['Estimated Impact'].map(impact_mapping)

# Map 'Priority' to colors
priority_color_mapping = {'P1 🔥': 'red', 'P2': 'orange', 'P3': 'green', 'P4': 'blue'}
project_df['Color'] = project_df['Priority'].map(priority_color_mapping)

# Create the scatter plot with jittered points, varying colors, and circle markers
plt.figure(figsize=(14, 10))

# Plot each priority level separately to create a custom legend
for priority, color in priority_color_mapping.items():
    subset = project_df[project_df['Priority'] == priority]
    plt.scatter(add_jitter(subset['Numerical Effort']), add_jitter(subset['Numerical Impact']), c=color, label=priority, s=100, alpha=0.6, marker='o')

# Add titles and labels
plt.title('Project Comparison: Estimated Effort vs Estimated Impact (Jittered)')
plt.xlabel('Estimated Effort (Numerical)')
plt.ylabel('Estimated Impact (Numerical)')
plt.grid(True)

# Add a legend to explain the colors
plt.legend(title='Priority')

# Display the plot
plt.show()

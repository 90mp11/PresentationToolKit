import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

def add_jitter(series, jitter_amount=0.1):
    """Add jitter to a Pandas series to avoid overlapping points in a scatter plot."""
    return series + (np.random.rand(len(series)) - 0.5) * jitter_amount

def load_and_prepare_data(csv_file_path):
    """Load the project data from a CSV file and map categorical values to numerical values."""
    effort_mapping = {'Hours': 1, 'Days': 2, 'Weeks': 3, 'Months': 4, 'Quarters': 5}
    impact_mapping = {'Very Low': 1, 'Low': 2, 'Medium': 3, 'High': 4, 'Very High': 5}
    priority_color_mapping = {'P1 🔥': 'red', 'P2': 'orange', 'P3': 'green', 'P4': 'blue'}

    # Read the CSV file
    project_df = pd.read_csv(csv_file_path)

    # Map 'Estimated Effort' and 'Estimated Impact' to numerical values
    project_df['Numerical Effort'] = project_df['Estimated Effort'].map(effort_mapping)
    project_df['Numerical Impact'] = project_df['Estimated Impact'].map(impact_mapping)

    # Map 'Priority' to colors
    project_df['Color'] = project_df['Priority'].map(priority_color_mapping)

    return project_df, priority_color_mapping

def create_scatter_plot(df, priority_color_mapping, jitter_amount=0.1):
    """Create a scatter plot with jittered points, varying colors, and circle markers."""
    plt.figure(figsize=(14, 10))

    # Plot each priority level separately to create a custom legend
    for priority, color in priority_color_mapping.items():
        subset = df[df['Priority'] == priority]
        plt.scatter(add_jitter(subset['Numerical Effort'], jitter_amount), 
                    add_jitter(subset['Numerical Impact'], jitter_amount), 
                    c=color, label=priority, s=100, alpha=0.6, marker='o')

    # Add titles and labels
    plt.title('Project Comparison: Estimated Effort vs Estimated Impact (Jittered)')
    plt.xlabel('Estimated Effort (Numerical)')
    plt.ylabel('Estimated Impact (Numerical)')
    plt.grid(True)

    # Add a legend to explain the colors
    plt.legend(title='Priority')

    # Display the plot
    plt.show()

def create_age_bar_chart(df, output_path='age_bar_chart.png'):
    # Create the horizontal bar chart
    plt.figure(figsize=(10, 6))
    bars = plt.barh(df['AssignedTo'], df['Age_BusinessDays'], color='skyblue')
    plt.xlabel('Total Age in Business Days')
    plt.ylabel('Assigned To')
    plt.title('Total Age of Open Tickets by Assigned Person')
    plt.gca().invert_yaxis()  # Invert the y-axis to have the longest bar on top

    # Add labels on each bar with the number of tickets
    for bar, count in zip(bars, df['TicketCount']):
        width = bar.get_width()
        plt.text(width, bar.get_y() + bar.get_height()/2, f'{count} tickets', 
                 va='center', ha='left', fontsize=10, color='black')

    # Save the chart as an image
    plt.savefig(output_path, bbox_inches='tight')
    plt.close()

# Example usage of the refactored functions
if __name__ == "__main__":
    csv_file_path = './raw/PROJECT.csv'
    project_df, priority_color_mapping = load_and_prepare_data(csv_file_path)
    create_scatter_plot(project_df, priority_color_mapping)

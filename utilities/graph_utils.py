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

def plot_resolved_items_per_month(df, output_path='resolved_month_chart.png'):
    """
    Plots a stacked bar chart showing the number of resolved items per engineer per month.
    
    df: DataFrame with resolved items per month per engineer.
    """
    # Pivot the data to get engineers (Closed by) as columns and YearMonth as rows
    pivot_table = df.pivot(index='YearMonth', columns='Closed by', values='ResolvedCount').fillna(0)

    # Plot a stacked bar chart
    pivot_table.plot(kind='bar', stacked=True, figsize=(10, 7), cmap='tab20')
    
    plt.title('Resolved Items Per Month (Closed by)')
    plt.xlabel('Month')
    plt.ylabel('Number of Resolved Items')
    plt.legend(title='Closed By', bbox_to_anchor=(1.05, 1), loc='upper left')
    plt.tight_layout()
    plt.savefig(output_path, bbox_inches='tight')
    plt.close()

def plot_grouped_resolved_items_per_month(df, output_path='grouped_resolved_month_chart.png'):
    """
    Plots a grouped bar chart showing the number of resolved items per engineer per month.
    
    df: DataFrame with resolved items per month per engineer.
    output_path: Path to save the generated chart image.
    """
    # Pivot the data to get engineers (Closed by) as columns and YearMonth as rows
    pivot_table = df.pivot(index='YearMonth', columns='Closed by', values='ResolvedCount').fillna(0)

    # Plot a grouped bar chart
    ax = pivot_table.plot(kind='bar', stacked=False, figsize=(10, 7), cmap='tab20')

    plt.title('')
    plt.xlabel('')
    plt.ylabel('')
    plt.legend(title='Closed By', bbox_to_anchor=(1.05, 1), loc='upper left')
    plt.tight_layout()
    plt.savefig(output_path, bbox_inches='tight')
    plt.close()

def plot_engineer_grouped_resolved_items(df, output_path='engineer_grouped_resolved_chart.png'):
    """
    Plots a bar chart grouped by engineer, showing the number of resolved items per engineer per month.
    
    df: DataFrame with resolved items per month per engineer.
    output_path: Path to save the generated chart image.
    """
    # Pivot the data so that engineers are the index, and each month is a separate column
    pivot_table = df.pivot(index='Closed by', columns='YearMonth', values='ResolvedCount').fillna(0)

    # Plot the grouped bar chart
    ax = pivot_table.plot(kind='bar', stacked=False, figsize=(12, 8), cmap='tab20')

    plt.title('')
    plt.xlabel('')
    plt.ylabel('Number of Tickets Closed')
    plt.legend(title='Month', bbox_to_anchor=(1.05, 1), loc='upper left')
    plt.tight_layout()
    plt.savefig(output_path, bbox_inches='tight')
    plt.close()

def plot_resolution_time_by_engineer(df, output_path='resolution_time_chart.png'):
    """
    Plots a bar chart showing the average resolution time by engineer.
    
    df: DataFrame with columns ['AssignedTo', 'TimeToResolve_BusinessDays'].
    output_path: Path to save the generated chart image.
    """
    # Group by engineer and calculate the average resolution time
    grouped_df = df.groupby('Closed by')['TimeToResolve_BusinessDays'].mean().reset_index()
    
    # Sort the data for better visualization
    grouped_df = grouped_df.sort_values(by='TimeToResolve_BusinessDays', ascending=False)

    # Plot the bar chart
    plt.figure(figsize=(10, 6))
    bars = plt.barh(grouped_df['Closed by'], grouped_df['TimeToResolve_BusinessDays'], color='skyblue')
    plt.xlabel('Average Resolution Time in Business Days')
    plt.ylabel('')
    plt.title('')
    plt.gca().invert_yaxis()  # Invert the y-axis to have the longest bar on top
    
    # Add labels on each bar
    for bar in bars:
        width = bar.get_width()
        plt.text(width, bar.get_y() + bar.get_height()/2, f'{width:.1f} days', 
                 va='center', ha='left', fontsize=10, color='black')
    
    # Save the chart as an image
    plt.savefig(output_path, bbox_inches='tight')
    plt.close()

# Example usage of the refactored functions
if __name__ == "__main__":
    csv_file_path = './raw/PROJECT.csv'
    project_df, priority_color_mapping = load_and_prepare_data(csv_file_path)
    create_scatter_plot(project_df, priority_color_mapping)

def plot_claim_time_summary(df, output_path='claim_time_summary_chart.png'):
    """
    Plot the summary of claim times, showing average claim time and count of tickets exceeding 2 business days.
    
    df: Summary dataframe with average claim time, count exceeding 2 days, and total tickets.
    output_path: Path to save the generated chart image.
    """
    fig, ax1 = plt.subplots(figsize=(10, 6))

    ax2 = ax1.twinx()
    
    width = 0.4
    df.set_index('AssignedTo', inplace=True)
    
    df['avg_claim_time'].plot(kind='bar', color='blue', ax=ax1, width=width, position=1)
    df['exceed_two_days'].plot(kind='bar', color='red', ax=ax2, width=width, position=0)

    ax1.set_ylabel('Average Claim Time (Business Days)', color='blue')
    ax2.set_ylabel('Tickets Exceeding 2 Days', color='red')

    # Set y-axis limits and enforce integer ticks on the right axis
    ax1.set_ylim(0, max(1, df['avg_claim_time'].max() * 1.2))
    ax2.set_ylim(0, max(1, df['exceed_two_days'].max() * 1.2))
    ax2.yaxis.set_major_locator(plt.MaxNLocator(integer=True))

    ax1.set_xlabel('Engineer')
    ax1.set_title('Claim Time Summary by Engineer')
    
    # Annotate bar heights
    for p in ax1.patches:
        ax1.annotate(f'{p.get_height():.1f}', (p.get_x() + p.get_width() / 2., p.get_height()),
                     ha='center', va='center', xytext=(0, 10), textcoords='offset points')
    
    for p in ax2.patches:
        ax2.annotate(f'{p.get_height():.0f}', (p.get_x() + p.get_width() / 2., p.get_height()),
                     ha='center', va='center', xytext=(0, 10), textcoords='offset points')

    plt.tight_layout()
    plt.savefig(output_path)
    plt.close()

from datetime import datetime, timedelta
import pandas as pd
import numpy as np
from bs4 import BeautifulSoup
import html
from typing import Union, List

def create_blank_dataframe(csv_file='./raw/DATA.csv'):
    df = pd.read_csv(csv_file)
    return df

def filter_dataframe_by_team(dataframe, team_name):
    filtered_df = dataframe[dataframe['Impacted Teams'].apply(lambda teams: team_name in teams)]
    return filtered_df

def filter_dataframe_by_status(dataframe, statuses):
    if isinstance(statuses, str):
        # If a single status is passed as a string, convert it to a list
        statuses = [statuses]
    
    # Filter the dataframe for rows where the Status column contains any of the specified statuses
    filtered_df = dataframe[dataframe['Status'].apply(lambda status: any(s in status for s in statuses))]
    
    return filtered_df

def filter_dataframe_by_release_group(dataframe, group):
    filtered_df = dataframe[dataframe['Release Group'].apply(lambda groups: group in groups)]
    return filtered_df

def filter_dataframe_by_staging(dataframe, staging):
    filtered_df = dataframe[dataframe['Staging'].apply(lambda statuses: staging in statuses)]
    return filtered_df

def filter_dataframe_by_contact_owner(dataframe, owner):
    filtered_df = dataframe[dataframe['AssignedTo'].apply(lambda owners: owner in owners)]
    return filtered_df

def combine_dataframe(df1, df2):
    combined_df = pd.concat([df1, df2], ignore_index=True)
    return combined_df

def impacted_teams_list(df):
        # Define a function to handle NaN values and split strings
    def process_teams(teams_str):
        if isinstance(teams_str, str):
            return [team.strip() for team in teams_str.strip('[]').replace('"', '').split(',')]
        else:
            return []

    # Apply the process_teams function to the Impacted Teams column
    df['Impacted Teams'] = df['Impacted Teams'].apply(process_teams)

    # Create a set to store unique Impacted Teams
    unique_teams = set()

    # Iterate through the rows and add Impacted Teams to the set
    for teams_list in df['Impacted Teams']:
        unique_teams.update(teams_list)

    # Convert the set back to a list if needed
    unique_teams_list = list(unique_teams)

    return (unique_teams_list)

def convert_html_to_text_with_newlines(html_str):
    """
    Convert a given HTML string to plain text while preserving newlines and fixing encoding issues.
    This version also ensures that each sentence or paragraph is on a new line.
    
    Parameters:
    - html_str (str): The HTML string to convert.
    
    Returns:
    - str: A plain text representation of the HTML string with newlines preserved.
    """

    if html_str is None or (isinstance(html_str, (float, int)) and np.isnan(html_str)) or (isinstance(html_str, str) and html_str.strip() == ""):
        return ""

    # Use BeautifulSoup to parse the HTML    
    soup = BeautifulSoup(html_str, 'html.parser')
    
    # Initialize an empty string to store the final text
    final_text = ""
    
    # Iterate over each element in the HTML
    for elem in soup.stripped_strings:
        # Add the text of the element to the final text
        final_text += elem
        # Add a newline if the text ends with a period or is contained within a <p> tag
        if elem.endswith('.') or (hasattr(elem, 'parent') and elem.parent.name == 'p'):
            final_text += "\n"
    
    # Unescape HTML entities
    final_text = html.unescape(final_text)
    
    # Remove any unwanted characters (e.g., â€‹ can appear due to encoding issues)
    final_text = final_text.replace("â€‹", "")
    final_text = final_text.replace("&%23160;", "")
    final_text = final_text.replace("&%2358;", ":")
    final_text = final_text.replace("\r\n", "\n")

    return final_text.strip()  # Strip removes leading/trailing white spaces

def calculate_total_age(creation_date_str, reference_date_str=None):
    """
    Calculate the total age in days from the creation date to the reference date.
    
    Parameters:
    - creation_date_str: The creation date as a string (format: '%Y-%m-%d').
    - reference_date_str: The reference date as a string (format: '%Y-%m-%d'). Defaults to today.
    
    Returns:
    - int: Total age in days.
    """
    # Convert the creation date string to a datetime object
    creation_date = pd.to_datetime(creation_date_str, dayfirst=True)
    
    # If no reference date is provided, use the current date
    if reference_date_str is None:
        reference_date = datetime.now()
    else:
        reference_date = pd.to_datetime(reference_date_str, dayfirst=True)
    
    # Calculate the difference in days
    total_age_days = (reference_date - creation_date).days

    return total_age_days

def calculate_business_days_age(creation_date_str, reference_date_str=None):
    """
    Calculate the number of business days from the creation date to the reference date,
    accounting for tickets opened on weekends.
    
    Parameters:
    - creation_date_str: The creation date as a string (format: '%Y-%m-%d').
    - reference_date_str: The reference date as a string (format: '%Y-%m-%d'). Defaults to today.
    
    Returns:
    - int: Number of business days.
    """
    # Convert the creation date string to a datetime object
    creation_date = pd.to_datetime(creation_date_str, dayfirst=True)
    
    # If no reference date is provided, use the current date
    if reference_date_str is None:
        reference_date = datetime.now()
    else:
        reference_date = pd.to_datetime(reference_date_str, dayfirst=True)
    
    # Check if either date is NaT (Not a Time)
    if pd.isna(creation_date) or pd.isna(reference_date):
        return np.nan
    
    # If the ticket was created on a weekend, adjust the creation date to the following Monday
    if creation_date.weekday() >= 5:  # 5 is Saturday, 6 is Sunday
        creation_date += timedelta(days=(7 - creation_date.weekday()))
    
    # Generate a range of all dates between creation_date and reference_date
    all_days = pd.date_range(start=creation_date, end=reference_date, freq='B')  # 'B' stands for business days

    # Calculate the number of business days
    business_days = len(all_days)
    
    # Adjust for partial days if the ticket was opened or closed outside of business hours
    if creation_date.time() > datetime.strptime("09:00:00", "%H:%M:%S").time():
        business_days -= 1
    if reference_date.time() < datetime.strptime("17:00:00", "%H:%M:%S").time():
        business_days -= 1
    
    # Ensure the count doesn't go below zero
    business_days = max(business_days, 0)
    
    return business_days

def calculate_and_group_ticket_ages(df):
    # Ensure we are working on a copy of the DataFrame
    df = df.copy()

    # Ensure the Age_BusinessDays column is calculated
    df['Age_BusinessDays'] = df['OriginalCreationDate'].apply(lambda x: calculate_business_days_age(x))

    # Group by AssignedTo and calculate both the sum of Age_BusinessDays and the count of tickets
    grouped_df = df.groupby('AssignedTo').agg({
        'Age_BusinessDays': 'sum',  # Sum of business days
        'OriginalCreationDate': 'count'  # Count of tickets
    }).reset_index()

    # Rename 'OriginalCreationDate' to 'TicketCount' for clarity
    grouped_df = grouped_df.rename(columns={'OriginalCreationDate': 'TicketCount'})

    # Sort the grouped data by Age_BusinessDays in descending order
    sorted_df = grouped_df.sort_values(by='Age_BusinessDays', ascending=False)

    return sorted_df

import pandas as pd
from typing import Union, List

def filter_by_lead_team(df: pd.DataFrame, teams: Union[str, List[str]] = None) -> pd.DataFrame:
    """
    Filters the DataFrame to include only rows where the 'Lead Team' column
    matches the specified values. Defaults to 'Passive Engineering' and 'Regional Technical Leads'
    if no teams are provided.
    
    Parameters:
    df (pd.DataFrame): The input DataFrame to filter.
    teams (Union[str, List[str]], optional): A string or list of strings specifying the teams to filter by.
                                            Defaults to ['Passive Engineering', 'Regional Technical Leads'].
    
    Returns:
    pd.DataFrame: The original DataFrame if 'Lead Team' column is not found,
                  otherwise a new DataFrame filtered by the specified 'Lead Team' values.
    """
    # Check if 'Lead Team' column exists
    if 'Lead Team' not in df.columns:
        print("Warning: 'Lead Team' column not found. Returning the original DataFrame.")
        return df
    
    # Set default teams if none are provided
    if teams is None:
        teams = ["Passive Engineering", "Regional Technical Leads"]
    
    # Ensure that teams is always a list
    if isinstance(teams, str):
        teams = [teams]
    
    # Filter the DataFrame
    filtered_df = df[df['Lead Team'].isin(teams)]
    
    return filtered_df

def calculate_time_to_resolve(df):
    """
    Calculate the time taken to resolve each ticket in business days and total days.
    Returns a dataframe with additional columns for these calculations.
    """
    df = df.copy()
    
    # Ensure the columns are in datetime format
    df['Resolved Time'] = pd.to_datetime(df['Completed Time'], errors='coerce', dayfirst=True)
    df['Creation Time'] = pd.to_datetime(df['OriginalCreationDate'], errors='coerce', dayfirst=True)
    
    # Drop rows with NaT in either 'Resolved Time' or 'Creation Time'
    df = df.dropna(subset=['Resolved Time', 'Creation Time'])
    
    # Calculate time to resolve in total days
    df['TimeToResolve_Days'] = (df['Resolved Time'] - df['Creation Time']).dt.days
    
    # Calculate time to resolve in business days using a helper function
    df['TimeToResolve_BusinessDays'] = df.apply(lambda row: calculate_business_days_age(row['OriginalCreationDate'], row['Completed Time']), axis=1)
    
    return df

def calculate_claim_time(df):
    """
    Calculate the time taken to claim each ticket in business days and total days.
    Returns a dataframe with additional columns for these calculations, including unclaimed tickets.
    """
    df = df.copy()
    
    # Ensure the columns are in datetime format
    df['Claimed Date'] = pd.to_datetime(df['Claimed Date'], errors='coerce', dayfirst=True)
    df['Creation Time'] = pd.to_datetime(df['OriginalCreationDate'], errors='coerce', dayfirst=True)
    
    # Keep rows where 'Creation Time' is valid, but allow NaT in 'Claimed Date'
    df = df.dropna(subset=['Creation Time'])

    # Calculate time to claim in business days, but only where 'Claimed Date' is not NaT
    df['TimeToClaim_BusinessDays'] = df.apply(
        lambda row: calculate_business_days_age(row['Creation Time'], row['Claimed Date']) 
        if pd.notna(row['Claimed Date']) else None, 
        axis=1
    )
    
    return df

def analyze_claim_times(df):
    """
    Analyze the time to claim tickets by engineer.
    Returns a summary dataframe with average claim time, count of tickets exceeding 2 business days,
    total ticket count per engineer, and the number of unclaimed tickets.
    """
    # Count total tickets, including unclaimed
    total_tickets = df.groupby('AssignedTo').size()
    
    # Count unclaimed tickets (where 'Claimed Date' is NaT)
    unclaimed_tickets = df[df['Claimed Date'].isna()].groupby('AssignedTo').size()
    
    # Analyze claim times only for the claimed tickets
    summary = df.dropna(subset=['Claimed Date']).groupby('AssignedTo').agg(
        avg_claim_time=('TimeToClaim_BusinessDays', 'mean'),
        exceed_two_days=('TimeToClaim_BusinessDays', lambda x: (x > 2).sum()),
        total_tickets=('TimeToClaim_BusinessDays', 'count')
    ).reset_index()
    
    # Merge in total tickets and unclaimed tickets
    summary = summary.merge(total_tickets.rename('total_tickets_assigned'), on='AssignedTo', how='left')
    summary = summary.merge(unclaimed_tickets.rename('unclaimed_tickets'), on='AssignedTo', how='left')
    
    # Fill any NaN values in unclaimed_tickets with 0
    summary['unclaimed_tickets'] = summary['unclaimed_tickets'].fillna(0).astype(int)
    
    return summary

def pre_filter_creation_time(df, start_date='2024-01-01', end_date='2024-12-31'):
    """
    Filters the DataFrame for created tickets within the specified date range
    
    Returns the filtered DataFrame.
    """

    # Convert start_date and end_date to datetime
    start_date = pd.to_datetime(start_date)
    end_date = pd.to_datetime(end_date)

    # Filter for rows where Status is 'Resolved' and within the date range
    mask = pd.to_datetime(df['OriginalCreationDate'], errors='coerce', dayfirst=True).between(start_date, end_date)
    df_filtered = df[mask].copy()

    return df_filtered

def filter_and_aggregate_resolution_time(df, start_date='2024-01-01', end_date='2024-12-31', field='AssignedTo'):
    """
    Filters the DataFrame for resolved tickets within the specified date range and
    calculates the average resolution time by engineer.
    
    Returns both the filtered DataFrame and the grouped DataFrame.
    """
    # Calculate the time to resolve for each ticket
    df = calculate_time_to_resolve(df)

    # Convert start_date and end_date to datetime
    start_date = pd.to_datetime(start_date)
    end_date = pd.to_datetime(end_date)

    # Filter for rows where Status is 'Resolved' and within the date range
    mask = (df['Status'] == 'Resolved') & \
           (pd.to_datetime(df['Completed Time'], errors='coerce', dayfirst=True).between(start_date, end_date))
    df_filtered = df[mask].copy()

    # Group by engineer and calculate the average resolution time in business days
    df_grouped = df_filtered.groupby(field)['TimeToResolve_BusinessDays'].mean().reset_index()

    return df_filtered, df_grouped

def filter_dataframe_by_names(df: pd.DataFrame, selected_names: list, field = 'Closed by') -> pd.DataFrame:
    """
    Filters the DataFrame by the selected names in the 'ClosedBy' column.

    :param df: The input DataFrame to filter.
    :param selected_names: A list of names to filter by.
    :return: A filtered DataFrame containing only rows where 'ClosedBy' matches a name in selected_names.
    """
    if 'Closed by' not in df.columns:
        raise ValueError("DataFrame must contain a 'ClosedBy' column.")
    
    # Filter the DataFrame
    filtered_df = df[df[field].isin(selected_names)]
    
    return filtered_df
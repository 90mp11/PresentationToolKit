from datetime import datetime, timedelta
import pandas as pd
import numpy as np
from bs4 import BeautifulSoup
import html

def create_blank_dataframe(csv_file='./raw/DATA.csv'):
    df = pd.read_csv(csv_file)
    return df

def filter_dataframe_by_team(dataframe, team_name):
    filtered_df = dataframe[dataframe['Impacted Teams'].apply(lambda teams: team_name in teams)]
    return filtered_df

def filter_dataframe_by_status(dataframe, status):
    filtered_df = dataframe[dataframe['Status'].apply(lambda statuses: status in statuses)]
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

from bs4 import BeautifulSoup
import html

from bs4 import BeautifulSoup
import html

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
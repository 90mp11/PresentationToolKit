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

def filter_dataframe_by_staging(dataframe, staging):
    filtered_df = dataframe[dataframe['Staging'].apply(lambda statuses: staging in statuses)]
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
            final_text += '\n'
    
    # Unescape HTML entities
    final_text = html.unescape(final_text)
    
    # Remove any unwanted characters (e.g., â€‹ can appear due to encoding issues)
    final_text = final_text.replace("â€‹", "")
    final_text = final_text.replace("&%23160;", "")
    final_text = final_text.replace("&%2358;", ":")
    
    return final_text.strip()  # Strip removes leading/trailing white spaces

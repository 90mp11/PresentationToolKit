import pandas as pd

def create_blank_dataframe(csv_file='./raw/DATA.csv'):
    df = pd.read_csv(csv_file)
    return df

def filter_dataframe_by_team(dataframe, team_name):
    filtered_df = dataframe[dataframe['Impacted Teams'].apply(lambda teams: team_name in teams)]
    return filtered_df

def filter_dataframe_by_status(dataframe, status):
    filtered_df = dataframe[dataframe['Status'].apply(lambda statuses: status in statuses)]
    return filtered_df

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
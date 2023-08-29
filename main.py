import argparse
import data_utils as du
import presentation_utils as pu
import constants as const

# Define command-line arguments
parser = argparse.ArgumentParser(description="Run different queries")
parser.add_argument( # --engineering
    "--engineering",
    action="store_true",
    help="Create Batch of Presentations, one for each listed Team Member",
)
parser.add_argument( # --impact
    "--impact",
    action="store_true",
    help="Create Presentations for each Impacted Team",
)
parser.add_argument( # --allimpacted
    "--allimpacted",
    action="store_true",
    help="Create One Presentation with sections for each Impacted Team",
)
parser.add_argument( # --who
    "--who",
    action="store_true",
    help="Load prompt to filter by an individual's data",
)
parser.add_argument( # --output
    "--output",
    action="store_true",
    help="saves Dataframe as csv",
)

# Parse the command-line 
args = parser.parse_args()

# Determine which commands to execute based on the command-line arguments
if args.engineering:
    for eng in const.ENGINEERS:
        prs = pu.create_blank_presentation(const.FILE_LOCATIONS['pptx_template'])
        df = du.create_blank_dataframe(const.FILE_LOCATIONS['project_csv'])
        pu.create_ProjectOwner_slides(df, prs, eng)
        pu.create_Objective_slides(df, prs, eng)
        pu.save_exit(prs, "_"+eng)
    exit(1)
elif args.who:
    name_filter = input("Enter the custom workspace name: ")
    prs = pu.create_blank_presentation(const.FILE_LOCATIONS['pptx_template'])
    df = du.create_blank_dataframe(const.FILE_LOCATIONS['project_csv'])
    pu.create_ProjectOwner_slides(df, prs, name_filter)
    pu.create_Objective_slides(df, prs, name_filter)
    pu.save_exit(prs, "_"+name_filter)
    exit(1)
elif args.impact:
    df = du.create_blank_dataframe(const.FILE_LOCATIONS['project_csv'])
    impacted = du.impacted_teams_list(df)
    for imp in impacted:
        prs = pu.create_blank_presentation(const.FILE_LOCATIONS['pptx_template'])
        pu.create_Impacted_section(df, prs, no_section=True, impacted_team=imp)
        pu.save_exit(prs, "_"+imp)
    exit(1)
elif args.allimpacted:
    df = du.create_blank_dataframe(const.FILE_LOCATIONS['project_csv'])
    impacted = du.impacted_teams_list(df)
    prs = pu.create_blank_presentation(const.FILE_LOCATIONS['pptx_template'])
    for imp in impacted:
        pu.create_Impacted_section(df, prs, no_section=True, impacted_team=imp)
    pu.save_exit(prs, "_AllImpacts")
    exit(1)
elif args.output:
    df = du.create_blank_dataframe(const.FILE_LOCATIONS['project_csv'])
    df.to_csv("./dataframe.csv")
    exit(1)
else:
    df = du.create_blank_dataframe(const.FILE_LOCATIONS['project_csv'])
    prs = pu.create_blank_presentation(const.FILE_LOCATIONS['pptx_template'])
    pu.create_ProjectOwner_slides(df, prs)
    pu.create_Objective_slides(df, prs)
    pu.create_title_slide(prs, "Impacted Teams")
    impacted = du.impacted_teams_list(df)
    for imp in impacted:
        pu.create_Impacted_section(df, prs, no_section=True, impacted_team=imp)
    pu.save_exit(prs)
    exit(1)
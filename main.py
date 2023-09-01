import argparse
import data_utils as du
import presentation_utils as pu
import constants as const

# Define command-line arguments
parser = argparse.ArgumentParser(description="Project Reporting tool developed by the Passive Engineering Team. For specific help, please contact Matt Proctor")
parser.add_argument( # --engineering
    "--engineering",
    action="store_true",
    help="Create a batch of Presentations, one for each listed Team Member in the constants.py file",
)
parser.add_argument( # --impact
    "--impact",
    action="store_true",
    help="Create individual presentations for each of the unique teams listed in the 'Impacted Teams' column",
)
parser.add_argument( # --allimpacted
    "--allimpacted",
    action="store_true",
    help="Create single presentation to contain impacts against each of the unique teams listed in the 'Impacted Teams' column",
)
parser.add_argument( # --who
    "--who",
    action="store_true",
    help="Creates a single presentation filtered based on user input to the 'Primary Owner' field - this works on incomplete names",
)
parser.add_argument( # --output
    "--output",
    action="store_true",
    help="Used for testing / debug",
)
parser.add_argument( # --onhold
    "--onhold",
    action="store_true",
    help="Exports all the on-hold projects to a single presentation",
)
parser.add_argument( # --objective
    "--objective",
    action="store_true",
    help="Exports the Objective view into a single presentation",
)
parser.add_argument( # --projects
    "--projects",
    action="store_true",
    help="Exports the Objective view into a single presentation",
)
parser.add_argument( # --docs
    "--docs",
    action="store_true",
    help="Imports the Document Change Log and outputs the Monthly Release Board slides",
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
    print('debug')
    exit(1)
elif args.objective:
    df = du.create_blank_dataframe(const.FILE_LOCATIONS['project_csv'])
    prs = pu.create_blank_presentation(const.FILE_LOCATIONS['pptx_template'])
    pu.create_Objective_slides(df, prs)
    pu.save_exit(prs, "_Objective")
    exit(1)
elif args.onhold:
    df = du.create_blank_dataframe(const.FILE_LOCATIONS['project_csv'])
    prs = pu.create_blank_presentation(const.FILE_LOCATIONS['pptx_template'])
    pu.create_OnHold_slides(df, prs, no_section=True)
    pu.save_exit(prs, "_OnHold")
elif args.projects:
    df = du.create_blank_dataframe(const.FILE_LOCATIONS['project_csv'])
    prs = pu.create_blank_presentation(const.FILE_LOCATIONS['pptx_template'])
    pu.create_project_section(df, prs)
    pu.save_exit(prs, "_Projects")
    exit(1)
elif args.docs:
    df = du.create_blank_dataframe(const.FILE_LOCATIONS['document_csv'])
    prs = pu.create_blank_presentation(const.FILE_LOCATIONS['pptx_template'])
    pu.create_document_release_section(df, prs)
    pu.save_exit(prs, "_DocumentBoard")
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
    #pu.create_OnHold_slides(df, prs)
    pu.save_exit(prs)
    exit(1)
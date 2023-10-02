import argparse
import utilities.data_utils as du
import utilities.presentation_utils as pu
import utilities.constants as const

def main():
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
            pu.save_exit(prs, "_"+eng, const.FILE_LOCATIONS['output_folder'])
        exit(1)
    elif args.who:
        name_filter = input("Name: ")
        prs = pu.create_blank_presentation(const.FILE_LOCATIONS['pptx_template'])
        df = du.create_blank_dataframe(const.FILE_LOCATIONS['project_csv'])
        pu.create_ProjectOwner_slides(df, prs, name_filter)
        pu.create_Objective_slides(df, prs, name_filter)
        pu.save_exit(prs, "_"+name_filter, const.FILE_LOCATIONS['output_folder'])
        exit(1)
    elif args.impact:
        df = du.create_blank_dataframe(const.FILE_LOCATIONS['project_csv'])
        impacted = du.impacted_teams_list(df)
        for imp in impacted:
            prs = pu.create_blank_presentation(const.FILE_LOCATIONS['pptx_template'])
            pu.create_Impacted_section(df, prs, no_section=True, impacted_team=imp)
            pu.save_exit(prs, "_"+imp, const.FILE_LOCATIONS['output_folder'])
        exit(1)
    elif args.allimpacted:
        df = du.create_blank_dataframe(const.FILE_LOCATIONS['project_csv'])
        impacted = du.impacted_teams_list(df)
        prs = pu.create_blank_presentation(const.FILE_LOCATIONS['pptx_template'])
        for imp in impacted:
            pu.create_Impacted_section(df, prs, no_section=True, impacted_team=imp)
        pu.save_exit(prs, "_AllImpacts", const.FILE_LOCATIONS['output_folder'])
        exit(1)
    elif args.output:
        prs = pu.create_blank_presentation(const.FILE_LOCATIONS['pptx_template'])
        slide = prs.slides.add_slide(prs.slide_masters[1].slide_layouts[7]) #Doc Release Board Template
        pu.placeholder_identifier(slide)
        exit(1)
    elif args.objective:
        df = du.create_blank_dataframe(const.FILE_LOCATIONS['project_csv'])
        prs = pu.create_blank_presentation(const.FILE_LOCATIONS['pptx_template'])
        pu.create_Objective_slides(df, prs)
        pu.save_exit(prs, "_Objective", const.FILE_LOCATIONS['output_folder'])
        exit(1)
    elif args.onhold:
        df = du.create_blank_dataframe(const.FILE_LOCATIONS['project_csv'])
        prs = pu.create_blank_presentation(const.FILE_LOCATIONS['pptx_template'])
        pu.create_OnHold_slides(df, prs, no_section=True)
        pu.save_exit(prs, "_OnHold", const.FILE_LOCATIONS['output_folder'])
    elif args.projects:
        df = du.create_blank_dataframe(const.FILE_LOCATIONS['project_csv'])
        prs = pu.create_blank_presentation(const.FILE_LOCATIONS['pptx_template'])
        pu.create_project_section(df, prs)
        pu.save_exit(prs, "_Projects", const.FILE_LOCATIONS['output_folder'])
        exit(1)
    elif args.docs:
        name_filter = input("Date: ")
        df = du.create_blank_dataframe(const.FILE_LOCATIONS['document_csv'])
        prs = pu.create_blank_presentation(const.FILE_LOCATIONS['pptx_template'])
        pu.create_document_release_section(df, prs, name_filter)
        pu.save_exit(prs, "_DocumentBoard", const.FILE_LOCATIONS['output_folder'])
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
        pu.save_exit(prs, folder=const.FILE_LOCATIONS['output_folder'])
        exit(1)

def person_filter(project_csv, output_folder, person):
    df = du.create_blank_dataframe(project_csv)
    prs = pu.create_blank_presentation(const.FILE_LOCATIONS['pptx_template'])
    pu.create_ProjectOwner_slides(df, prs, person)
    pu.create_Objective_slides(df, prs, person)
    output_path = pu.save_exit(prs, "_"+person, output_folder)
    return output_path

def allimpacted(project_csv, output_folder):
    df = du.create_blank_dataframe(project_csv)
    impacted = du.impacted_teams_list(df)
    prs = pu.create_blank_presentation(const.FILE_LOCATIONS['pptx_template'])
    for imp in impacted:
        pu.create_Impacted_section(df, prs, no_section=True, impacted_team=imp)
    output_path = pu.save_exit(prs, "_AllImpacts", output_folder)
    return output_path

def objective(project_csv, output_folder):
    df = du.create_blank_dataframe(project_csv)
    prs = pu.create_blank_presentation(const.FILE_LOCATIONS['pptx_template'])
    pu.create_Objective_slides(df, prs)
    output_path = pu.save_exit(prs, "_Objective", output_folder)
    return output_path

def output_all(project_csv, output_folder):
    df = du.create_blank_dataframe(project_csv)
    prs = pu.create_blank_presentation(const.FILE_LOCATIONS['pptx_template'])
    pu.create_ProjectOwner_slides(df, prs)
    pu.create_Objective_slides(df, prs)
    pu.create_title_slide(prs, "Impacted Teams")
    impacted = du.impacted_teams_list(df)
    for imp in impacted:
        pu.create_Impacted_section(df, prs, no_section=True, impacted_team=imp)
    output_path = pu.save_exit(prs, folder = output_folder)
    return output_path

if __name__ == "__main__":
    main()


import argparse
import os
import utilities.data_utils as du
import utilities.presentation_utils as pu
import utilities.constants as const

def engineering_presentation():
    for eng in const.ENGINEERS:
        person_filter(person=eng)
    for rtl in const.RTL:
        person_filter(person=rtl)

def impact_presentation(impact_filter=None):
    if not impact_filter:
        impact_filter = input("Impacted Area: ")
    impact_slides(filter=impact_filter)

def allimpacted_presentation():
    allimpacted()

def who_presentation(name_filter=None):
    if not name_filter:
        name_filter = input("Name: ")
    person_filter(person=name_filter)

def onhold_presentation():
    df = du.create_blank_dataframe(const.FILE_LOCATIONS['project_csv'])
    prs = pu.create_blank_presentation(const.FILE_LOCATIONS['pptx_template'])
    pu.create_OnHold_slides(df, prs, no_section=True)
    pu.save_exit(prs, "PEA_Project_Report", "_OnHold", const.FILE_LOCATIONS['output_folder'])

def objective_presentation():
    objective()

def projects_presentation():
    df = du.create_blank_dataframe(const.FILE_LOCATIONS['project_csv'])
    prs = pu.create_blank_presentation(const.FILE_LOCATIONS['pptx_template'])
    pu.create_project_section(df, prs)
    pu.save_exit(prs, "PEA_Project_Report", "_Projects", const.FILE_LOCATIONS['output_folder'])

def docs_presentation(date_filter=None):
    if not date_filter:
        date_filter = input("Date: ")
    all_docs(name_filter=date_filter)

def document_changes_presentation():
    doc_changes(const.FILE_LOCATIONS['document_csv'], const.FILE_LOCATIONS['output_folder'])

def output_all_presentation():
    output_all()

def release_presentation(release_group=None, internal=False):
    if not release_group:
        release_group = input("Release Group: ")
    if release_group == "BDUK":
        release_board_slides_multi_filter(filter=["BDUK - P1", "BDUK - P2", "BDUK - P3", "BDUK - P4"], internal=internal)
    else:
        release_board_slides(filter=release_group, internal=internal)

def person_filter(project_csv=const.FILE_LOCATIONS['project_csv'], output_folder=const.FILE_LOCATIONS['output_folder'], person='Matt', save=True, prs=None):
    df = du.create_blank_dataframe(project_csv)
    if prs is None:
        prs = pu.create_blank_presentation(const.FILE_LOCATIONS['pptx_template'])
    pu.create_ProjectOwner_slides(df, prs, person)
    pu.create_Objective_slides(df, prs, person)
    output_path = ""
    if save:
        output_path = pu.save_exit(prs, "PEA_Project_Report", "_"+person, output_folder)
    return output_path

def impact_slides(project_csv=const.FILE_LOCATIONS['project_csv'], output_folder=const.FILE_LOCATIONS['output_folder'], filter=""):
    df = du.create_blank_dataframe(project_csv)
    impacted = du.impacted_teams_list(df)
    output_path = ""
    if filter == "":
        for imp in impacted:
            prs = pu.create_blank_presentation(const.FILE_LOCATIONS['pptx_template'])
            pu.create_Impacted_section(df, prs, no_section=True, impacted_team=imp)
            pu.save_exit(prs, "PEA_Project_Report", "_"+imp, output_folder)
    else:
        prs = pu.create_blank_presentation(const.FILE_LOCATIONS['pptx_template'])
        pu.create_Impacted_section(df, prs, no_section=True, impacted_team=filter)
        output_path = pu.save_exit(prs, "PEA_Project_Report", "_"+filter, output_folder)
    return output_path

def allimpacted(project_csv=const.FILE_LOCATIONS['project_csv'], output_folder=const.FILE_LOCATIONS['output_folder'], save=True, prs=None):
    df = du.create_blank_dataframe(project_csv)
    impacted = du.impacted_teams_list(df)
    if prs is None:
        prs = pu.create_blank_presentation(const.FILE_LOCATIONS['pptx_template'])
    output_path = ""
    for imp in impacted:
        pu.create_Impacted_section(df, prs, no_section=True, impacted_team=imp)
    if save:
        output_path = pu.save_exit(prs, "PEA_Project_Report", "_AllImpacts", output_folder)
    return output_path

def objective(project_csv=const.FILE_LOCATIONS['project_csv'], output_folder=const.FILE_LOCATIONS['output_folder'], save=True, prs=None):
    df = du.create_blank_dataframe(project_csv)
    if prs is None:
        prs = pu.create_blank_presentation(const.FILE_LOCATIONS['pptx_template'])
    pu.create_Objective_slides(df, prs)
    output_path = ""
    if save:
        output_path = pu.save_exit(prs, "PEA_Project_Report", "_Objective", output_folder)
    return output_path

def output_all(project_csv=const.FILE_LOCATIONS['project_csv'], output_folder=const.FILE_LOCATIONS['output_folder'], save=True, prs=None):
    df = du.create_blank_dataframe(project_csv)
    if prs is None:
        prs = pu.create_blank_presentation(const.FILE_LOCATIONS['pptx_template'])
    pu.create_AllProjects_slide(df, prs)
    pu.create_ProjectOwner_slides(df, prs)
    pu.create_Objective_slides(df, prs)
    pu.create_title_slide(prs, "Impacted Teams")
    impacted = du.impacted_teams_list(df)
    output_path = ""
    for imp in impacted:
        pu.create_Impacted_section(df, prs, no_section=True, impacted_team=imp)
    pu.create_OnHold_slides(df, prs)
    if save:
        output_path = pu.save_exit(prs, report_type="PEA_Project_Report", folder = output_folder)
    return output_path

def all_docs(project_csv=const.FILE_LOCATIONS['document_csv'], output_folder=const.FILE_LOCATIONS['output_folder'], name_filter='', save=True, prs=None):
    df = du.create_blank_dataframe(project_csv)
    if prs is None:
        prs = pu.create_blank_presentation(const.FILE_LOCATIONS['pptx_template'])
    pu.create_document_release_section(df, prs, name_filter)
    output_path = ""
    if save:
        output_path = pu.save_exit(prs, "PEA_Project_Report", "_DocumentBoard", folder = output_folder)
    return output_path

def doc_changes(project_csv=const.FILE_LOCATIONS['document_csv'], output_folder=const.FILE_LOCATIONS['output_folder'], save=True, prs=None):
    df = du.create_blank_dataframe(project_csv)
    if prs is None:
        prs = pu.create_blank_presentation(const.FILE_LOCATIONS['pptx_template'])
    pu.create_document_changes_section(df, prs)
    output_path = ""
    if save:
        output_path = pu.save_exit(prs, "PEA_Project_Report", "_DocumentChanges", folder = output_folder)
    return output_path

def release_board_slides(project_csv=const.FILE_LOCATIONS['document_csv'], output_folder=const.FILE_LOCATIONS['output_folder'], filter='', save=True, prs=None, internal=False):
    df = du.create_blank_dataframe(project_csv)
    save_tail = "_FullReleaseBoard"
    if filter:
        df = df.loc[df['Release Group'] == filter]   
        save_tail = "_"+filter
    impacted = du.impacted_teams_list(df)

    if internal:
        save_tail = save_tail + "_internal"

    if prs is None:
        prs = pu.create_blank_presentation(const.FILE_LOCATIONS['pptx_template'])

    pu.create_document_release_section(df, prs, filter, internal=internal)
    output_path = ""
    
    if not internal:
        for imp in impacted:
            pu.create_document_Impacted_section(df, prs, no_section=False, impacted_team=imp, group_filter=filter)
    if save:
        output_path = pu.save_exit(prs, "PEA_Document_Release", save_tail, folder = output_folder)

    return output_path

def release_board_slides_multi_filter(project_csv=const.FILE_LOCATIONS['document_csv'], output_folder=const.FILE_LOCATIONS['output_folder'], filter='[]', save=True, prs=None, internal=False):
    df = du.create_blank_dataframe(project_csv)
    save_tail = "_FullReleaseBoard"
    if filter:
        df = df[df['Release Group'].isin(filter)]    
        save_tail = "_" + "And".join(filter)  # Assuming you want to concatenate filter names for the filename
    impacted = du.impacted_teams_list(df)

    if prs is None:
        prs = pu.create_blank_presentation(const.FILE_LOCATIONS['pptx_template'])

    pu.create_document_release_section_multi_filter(df, prs, filter, internal=internal)
    output_path = ""
    
    if not internal:
        for imp in impacted:
            pu.create_document_Impacted_section_multi_filter(df, prs, no_section=False, impacted_team=imp, group_filter=filter)
    if save:
        output_path = pu.save_exit(prs, "PEA_Document_Release", save_tail, folder = output_folder)

    return output_path

def create_folders(folder_names=['output', 'raw', 'templates', 'utilities']):
    for folder_name in folder_names:
        # Construct the path for the folder
        path = os.path.join(os.getcwd(), folder_name)
        
        # Check if the folder already exists to avoid trying to create it again
        if not os.path.exists(path):
            # Create the folder
            os.mkdir(path)
            print(f"Created missing folder: {path}")

def main():
    parser = argparse.ArgumentParser(description="Project Reporting tool developed by the Passive Engineering Team.")
    parser.add_argument("--engineering", action="store_true", help="Create a batch of Presentations, one for each listed Team Member in the constants.py file")
    parser.add_argument("--impact", action="store_true", help="Create individual presentations for each of the unique teams listed in the 'Impacted Teams' column")
    parser.add_argument("--allimpacted", action="store_true", help="Create single presentation to contain impacts against each of the unique teams listed in the 'Impacted Teams' column")
    parser.add_argument("--who", action="store_true", help="Creates a single presentation filtered based on user input to the 'Primary Owner' field - this works on incomplete names")
    parser.add_argument("--debug", action="store_true", help="Used for testing / debug")
    parser.add_argument("--onhold", action="store_true", help="Exports all the on-hold projects to a single presentation")
    parser.add_argument("--objective", action="store_true", help="Exports the Objective view into a single presentation")
    parser.add_argument("--projects", action="store_true", help="Exports the Objective view into a single presentation")
    parser.add_argument("--docs", action="store_true", help="Imports the Document Change Log and outputs the Monthly Release Board slides")
    parser.add_argument("--document_changes", action="store_true", help="Exports the Document Changes view into a single presentation")
    parser.add_argument("--output_all", action="store_true", help="Exports the Standard Output Report in a single presentation")
    parser.add_argument("--release", action="store_true", help="Outputs the ReleaseBoard Impact Report in a single presentation")
    parser.add_argument("--internal", action="store_true", help="Sets a flag for Internal use only - Outputs the ReleaseBoard Impact Report in a single presentation")

    args = parser.parse_args()
    internal = args.internal

    if args.engineering:
        engineering_presentation()
    if args.who:
        who_presentation()
    if args.impact:
        impact_presentation()
    if args.allimpacted:
        allimpacted_presentation()
    if args.objective:
        objective_presentation()
    if args.onhold:
        onhold_presentation()
    if args.projects:
        projects_presentation()
    if args.docs:
        docs_presentation()
    if args.document_changes:
        document_changes_presentation()
    if args.output_all:
        output_all_presentation()
    if args.release:
        release_presentation()

if __name__ == "__main__":
    create_folders()
    main()

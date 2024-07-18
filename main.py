import argparse
import os
import utilities.data_utils as du
import utilities.presentation_utils as pu
import utilities.constants as const

def engineering_presentation(project_csv, output_folder):
    for eng in const.ENGINEERS:
        person_filter(project_csv=project_csv, person=eng, output_folder=output_folder)
    for rtl in const.RTL:
        person_filter(project_csv=project_csv, person=rtl, output_folder=output_folder)

def impact_presentation(project_csv, impact_filter, output_folder):
    if isinstance(impact_filter, list):
        for filter_item in impact_filter:
            impact_slides(project_csv=project_csv, filter=filter_item, output_folder=output_folder)
    else:
        impact_slides(project_csv=project_csv, filter=impact_filter, output_folder=output_folder)

def allimpacted_presentation(project_csv, output_folder):
    allimpacted(project_csv=project_csv, output_folder=output_folder)

def who_presentation(project_csv, name_filter, output_folder):
    person_filter(project_csv=project_csv, person=name_filter, output_folder=output_folder)

def onhold_presentation(project_csv, output_folder):
    df = du.create_blank_dataframe(project_csv)
    prs = pu.create_blank_presentation(const.FILE_LOCATIONS['pptx_template'])
    pu.create_OnHold_slides(df, prs, no_section=True)
    pu.save_exit(prs, "PEA_Project_Report", "_OnHold", output_folder)

def objective_presentation(project_csv, output_folder):
    objective(project_csv=project_csv, output_folder=output_folder)

def projects_presentation(project_csv, output_folder):
    df = du.create_blank_dataframe(project_csv)
    prs = pu.create_blank_presentation(const.FILE_LOCATIONS['pptx_template'])
    pu.create_project_section(df, prs)
    pu.save_exit(prs, "PEA_Project_Report", "_Projects", output_folder)

def docs_presentation(document_csv, date_filter, output_folder):
    all_docs(project_csv=document_csv, name_filter=date_filter, output_folder=output_folder)

def document_changes_presentation(document_csv, output_folder):
    doc_changes(project_csv=document_csv, output_folder=output_folder)

def output_all_presentation(project_csv, output_folder):
    output_all(project_csv=project_csv, output_folder=output_folder)

def release_presentation(document_csv, release_group, internal, output_folder):
    if release_group == "BDUK":
        release_board_slides_multi_filter(project_csv=document_csv, filter=["BDUK - P1", "BDUK - P2", "BDUK - P3", "BDUK - P4"], internal=internal, output_folder=output_folder)
    else:
        release_board_slides(project_csv=document_csv, filter=release_group, internal=internal, output_folder=output_folder)

def person_filter(project_csv, output_folder, person='Matt', save=True, prs=None):
    df = du.create_blank_dataframe(project_csv)
    if prs is None:
        prs = pu.create_blank_presentation(const.FILE_LOCATIONS['pptx_template'])
    pu.create_ProjectOwner_slides(df, prs, person)
    pu.create_Objective_slides(df, prs, person)
    output_path = ""
    if save:
        output_path = pu.save_exit(prs, "PEA_Project_Report", "_"+person, output_folder)
    return output_path

def impact_slides(project_csv, output_folder, filter=""):
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

def allimpacted(project_csv, output_folder, save=True, prs=None):
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

def objective(project_csv, output_folder, save=True, prs=None):
    df = du.create_blank_dataframe(project_csv)
    if prs is None:
        prs = pu.create_blank_presentation(const.FILE_LOCATIONS['pptx_template'])
    pu.create_Objective_slides(df, prs)
    output_path = ""
    if save:
        output_path = pu.save_exit(prs, "PEA_Project_Report", "_Objective", output_folder)
    return output_path

def output_all(project_csv, output_folder, save=True, prs=None):
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

def all_docs(project_csv, output_folder, name_filter='', save=True, prs=None):
    df = du.create_blank_dataframe(project_csv)
    if prs is None:
        prs = pu.create_blank_presentation(const.FILE_LOCATIONS['pptx_template'])
    pu.create_document_release_section(df, prs, name_filter)
    output_path = ""
    if save:
        output_path = pu.save_exit(prs, "PEA_Project_Report", "_DocumentBoard", folder = output_folder)
    return output_path

def doc_changes(project_csv, output_folder, save=True, prs=None):
    df = du.create_blank_dataframe(project_csv)
    if prs is None:
        prs = pu.create_blank_presentation(const.FILE_LOCATIONS['pptx_template'])
    pu.create_document_changes_section(df, prs)
    output_path = ""
    if save:
        output_path = pu.save_exit(prs, "PEA_Project_Report", "_DocumentChanges", folder = output_folder)
    return output_path

def release_board_slides(project_csv, output_folder, filter='', save=True, prs=None, internal=False):
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

def release_board_slides_multi_filter(project_csv, output_folder, filter='[]', save=True, prs=None, internal=False):
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

def run_gui():
    from gui import start_gui
    start_gui()

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
    parser.add_argument("--project_csv", type=str, help="Specify the project CSV file")
    parser.add_argument("--document_csv", type=str, help="Specify the document CSV file")
    parser.add_argument("--output_folder", type=str, help="Specify the output folder")

    args = parser.parse_args()
    internal = args.internal
    project_csv = args.project_csv if args.project_csv else const.FILE_LOCATIONS['project_csv']
    document_csv = args.document_csv if args.document_csv else const.FILE_LOCATIONS['document_csv']
    output_folder = args.output_folder if args.output_folder else const.FILE_LOCATIONS['output_folder']

    if args.gui:
        run_gui()
    if args.engineering:
        engineering_presentation(project_csv=project_csv, output_folder=output_folder)
    if args.who:
        who_presentation(project_csv=project_csv, output_folder=output_folder)
    if args.impact:
        impact_presentation(project_csv=project_csv, impact_filter=None, output_folder=output_folder)
    if args.allimpacted:
        allimpacted_presentation(project_csv=project_csv, output_folder=output_folder)
    if args.objective:
        objective_presentation(project_csv=project_csv, output_folder=output_folder)
    if args.onhold:
        onhold_presentation(project_csv=project_csv, output_folder=output_folder)
    if args.projects:
        projects_presentation(project_csv=project_csv, output_folder=output_folder)
    if args.docs:
        docs_presentation(document_csv=document_csv, date_filter=None, output_folder=output_folder)
    if args.document_changes:
        document_changes_presentation(document_csv=document_csv, output_folder=output_folder)
    if args.output_all:
        output_all_presentation(project_csv=project_csv, output_folder=output_folder)
    if args.release:
        release_presentation(document_csv=document_csv, release_group=None, internal=internal, output_folder=output_folder)

if __name__ == "__main__":
    create_folders()
    main()

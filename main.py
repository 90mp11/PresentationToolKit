import argparse
import os
#import PresentationToolKit.utilities.data_utils as du
#import PresentationToolKit.utilities.presentation_utils as pu
#import PresentationToolKit.utilities.constants as const

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
    parser.add_argument( # --debug
        "--debug",
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
    parser.add_argument( # --document_changes
        "--document_changes",
        action="store_true",
        help="Exports the Document Changes view into a single presentation",
    )
    parser.add_argument( # --output_all
        "--output_all",
        action="store_true",
        help="Exports the Standard Output Report in a single presentation",
    )
    parser.add_argument( # --release
        "--release",
        action="store_true",
        help="Outputs the ReleaseBoard Impact Report in a single presentation",
    )
    parser.add_argument( # --internal
        "--internal",
        action="store_true",
        help="Sets a flag for Internal use only - Outputs the ReleaseBoard Impact Report in a single presentation",
    )

    # Parse the command-line 
    args = parser.parse_args()

    done_test=0
    internal = False

    # Determine which commands to execute based on the command-line arguments
    if args.internal:
        internal = True
    if args.engineering:
        for eng in const.ENGINEERS:
            person_filter(person=eng)
        done_test=1
    if args.who:
        name_filter = input("Name: ")
        person_filter(person=name_filter)
        done_test=1
    if args.impact:
        impact_filter = input("Impacted Area: ")
        impact_slides(filter=impact_filter)
        done_test=1
    if args.allimpacted:
        allimpacted()
        done_test=1
    if args.objective:
        objective()
        done_test=1
    if args.onhold:
        df = du.create_blank_dataframe(const.FILE_LOCATIONS['project_csv'])
        prs = pu.create_blank_presentation(const.FILE_LOCATIONS['pptx_template'])
        pu.create_OnHold_slides(df, prs, no_section=True)
        pu.save_exit(prs, "_OnHold", const.FILE_LOCATIONS['output_folder'])
        done_test=1
    if args.projects:
        df = du.create_blank_dataframe(const.FILE_LOCATIONS['project_csv'])
        prs = pu.create_blank_presentation(const.FILE_LOCATIONS['pptx_template'])
        pu.create_project_section(df, prs)
        pu.save_exit(prs, "_Projects", const.FILE_LOCATIONS['output_folder'])
        done_test=1
    if args.docs:
        name_filter = input("Date: ")
        all_docs(name_filter=name_filter)
        done_test=1
    if args.document_changes:
        doc_changes(const.FILE_LOCATIONS['document_csv'], const.FILE_LOCATIONS['output_folder'])
        done_test=1
    if args.debug:
        prs = pu.create_blank_presentation(const.FILE_LOCATIONS['pptx_template'])
        slide = prs.slides.add_slide(prs.slide_masters[1].slide_layouts[7]) #Doc Release Board Template
        pu.placeholder_identifier(slide)
        done_test=1
    if args.output_all:
        output_all()
        done_test=1
    if args.release:
        release_filter = input("Release Group: ")
        if release_filter == "BDUK":
            release_board_slides_multi_filter(filter=["BDUK - P1", "BDUK - P2", "BDUK - P3", "BDUK - P4"], internal=internal)
        else:
            release_board_slides(filter=release_filter, internal=internal)
        done_test=1

        # TODO - change filename to "Document", change template title to "Document", add filter so only External Review is selected for External Report

    if done_test == 0:
        output_all()
    exit(1)

### PROJECT BOARD FUNCTIONS ###

def person_filter(project_csv=const.FILE_LOCATIONS['project_csv'], output_folder=const.FILE_LOCATIONS['output_folder'], person='Matt', save=True, prs=None):
    df = du.create_blank_dataframe(project_csv)
    if prs is None:
        prs = pu.create_blank_presentation(const.FILE_LOCATIONS['pptx_template'])
    pu.create_ProjectOwner_slides(df, prs, person)
    pu.create_Objective_slides(df, prs, person)
    output_path = ""
    if save:
        output_path = pu.save_exit(prs, "_"+person, output_folder)
    return output_path

def impact_slides(project_csv=const.FILE_LOCATIONS['project_csv'], output_folder=const.FILE_LOCATIONS['output_folder'], filter=""):
    df = du.create_blank_dataframe(project_csv)
    impacted = du.impacted_teams_list(df)
    output_path = ""
    if filter == "":
        for imp in impacted:
            prs = pu.create_blank_presentation(const.FILE_LOCATIONS['pptx_template'])
            pu.create_Impacted_section(df, prs, no_section=True, impacted_team=imp)
            pu.save_exit(prs, "_"+imp, output_folder)
    else:
        prs = pu.create_blank_presentation(const.FILE_LOCATIONS['pptx_template'])
        pu.create_Impacted_section(df, prs, no_section=True, impacted_team=filter)
        output_path = pu.save_exit(prs, "_"+filter, output_folder)
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
        output_path = pu.save_exit(prs, "_AllImpacts", output_folder)
    return output_path

def objective(project_csv=const.FILE_LOCATIONS['project_csv'], output_folder=const.FILE_LOCATIONS['output_folder'], save=True, prs=None):
    df = du.create_blank_dataframe(project_csv)
    if prs is None:
        prs = pu.create_blank_presentation(const.FILE_LOCATIONS['pptx_template'])
    pu.create_Objective_slides(df, prs)
    output_path = ""
    if save:
        output_path = pu.save_exit(prs, "_Objective", output_folder)
    return output_path

def output_all(project_csv=const.FILE_LOCATIONS['project_csv'], output_folder=const.FILE_LOCATIONS['output_folder'], save=True, prs=None):
    df = du.create_blank_dataframe(project_csv)
    if prs is None:
        prs = pu.create_blank_presentation(const.FILE_LOCATIONS['pptx_template'])
    pu.create_ProjectOwner_slides(df, prs)
    pu.create_Objective_slides(df, prs)
    pu.create_title_slide(prs, "Impacted Teams")
    impacted = du.impacted_teams_list(df)
    output_path = ""
    for imp in impacted:
        pu.create_Impacted_section(df, prs, no_section=True, impacted_team=imp)
    pu.create_OnHold_slides(df, prs)
    if save:
        output_path = pu.save_exit(prs, folder = output_folder)
    return output_path

### DOCUMENT FUNCTIONS ###

def all_docs(project_csv=const.FILE_LOCATIONS['document_csv'], output_folder=const.FILE_LOCATIONS['output_folder'], name_filter='', save=True, prs=None):
    df = du.create_blank_dataframe(project_csv)
    if prs is None:
        prs = pu.create_blank_presentation(const.FILE_LOCATIONS['pptx_template'])
    pu.create_document_release_section(df, prs, name_filter)
    output_path = ""
    if save:
        output_path = pu.save_exit(prs, "_DocumentBoard", folder = output_folder)
    return output_path

def doc_changes(project_csv=const.FILE_LOCATIONS['document_csv'], output_folder=const.FILE_LOCATIONS['output_folder'], save=True, prs=None):
    df = du.create_blank_dataframe(project_csv)
    if prs is None:
        prs = pu.create_blank_presentation(const.FILE_LOCATIONS['pptx_template'])
    pu.create_document_changes_section(df, prs)
    output_path = ""
    if save:
        output_path = pu.save_exit(prs, "_DocumentChanges", folder = output_folder)
    return output_path

def release_board_slides(project_csv=const.FILE_LOCATIONS['document_csv'], output_folder=const.FILE_LOCATIONS['output_folder'], filter='', save=True, prs=None, internal=False):
    df = du.create_blank_dataframe(project_csv)
    save_tail = "_DocumentBoard"
    if filter:
        df = df.loc[df['Release Group'] == filter]   
        save_tail = "_"+filter
    impacted = du.impacted_teams_list(df)

    if prs is None:
        prs = pu.create_blank_presentation(const.FILE_LOCATIONS['pptx_template'])

    pu.create_document_release_section(df, prs, filter, internal=internal)
    output_path = ""
    
    if not internal:
        for imp in impacted:
            pu.create_document_Impacted_section(df, prs, no_section=False, impacted_team=imp, group_filter=filter)
    if save:
        output_path = pu.save_exit(prs, save_tail, folder = output_folder)

    return output_path

def release_board_slides_multi_filter(project_csv=const.FILE_LOCATIONS['document_csv'], output_folder=const.FILE_LOCATIONS['output_folder'], filter='[]', save=True, prs=None, internal=False):
    df = du.create_blank_dataframe(project_csv)
    save_tail = "_DocumentBoard"
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
        output_path = pu.save_exit(prs, save_tail, folder = output_folder)

    return output_path

### SETUP FUNCTIONS ###

def create_folders(folder_names=['output', 'raw', 'templates', 'utilities']):
    for folder_name in folder_names:
        # Construct the path for the folder
        path = os.path.join(os.getcwd(), folder_name)
        
        # Check if the folder already exists to avoid trying to create it again
        if not os.path.exists(path):
            # Create the folder
            os.mkdir(path)
            print(f"Created missing folder: {path}")

### NAME==MAIN ###
if __name__ == "__main__":
    create_folders()
    main()

import argparse
import os
import utilities.data_utils as du
import utilities.presentation_utils as pu
import utilities.constants as const
import utilities.builder as bu

def create_folders(folder_names=['output', 'raw', 'templates', 'utilities']):
    for folder_name in folder_names:
        # Construct the path for the folder
        path = os.path.join(os.getcwd(), folder_name)
        
        # Check if the folder already exists to avoid trying to create it again
        if not os.path.exists(path):
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
    parser.add_argument("--gui", action="store_true", help="Launch the graphical user interface")

    args = parser.parse_args()
    internal = args.internal
    project_csv = args.project_csv if args.project_csv else const.FILE_LOCATIONS['project_csv']
    document_csv = args.document_csv if args.document_csv else const.FILE_LOCATIONS['document_csv']
    output_folder = args.output_folder if args.output_folder else const.FILE_LOCATIONS['output_folder']

    trigger = 1

    if args.gui:
        trigger = 0
        run_gui()
    if args.engineering:
        trigger = 0
        bu.engineering_presentation(project_csv=project_csv, output_folder=output_folder)
    if args.who:
        trigger = 0
        bu.who_presentation(project_csv=project_csv, output_folder=output_folder)
    if args.impact:
        trigger = 0
        bu.impact_presentation(project_csv=project_csv, impact_filter=None, output_folder=output_folder)
    if args.allimpacted:
        trigger = 0
        bu.allimpacted_presentation(project_csv=project_csv, output_folder=output_folder)
    if args.objective:
        trigger = 0
        bu.objective_presentation(project_csv=project_csv, output_folder=output_folder)
    if args.onhold:
        trigger = 0
        bu.onhold_presentation(project_csv=project_csv, output_folder=output_folder)
    if args.projects:
        trigger = 0
        bu.projects_presentation(project_csv=project_csv, output_folder=output_folder)
    if args.docs:
        trigger = 0
        bu.docs_presentation(document_csv=document_csv, date_filter=None, output_folder=output_folder)
    if args.document_changes:
        trigger = 0
        bu.document_changes_presentation(document_csv=document_csv, output_folder=output_folder)
    if args.output_all:
        trigger = 0
        bu.output_all_presentation(project_csv=project_csv, output_folder=output_folder)
    if args.release:
        trigger = 0
        bu.release_presentation(document_csv=document_csv, release_group=None, internal=internal, output_folder=output_folder)
    if trigger:
        run_gui()

if __name__ == "__main__":
    main()

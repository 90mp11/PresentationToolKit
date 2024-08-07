import os
import sys
import utilities.data_utils as du
import utilities.presentation_utils as pu
import utilities.constants as const

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
    
    return os.path.join(base_path, relative_path)

def engineering_review_board_presentation(project_csv, output_folder):
    df = du.create_blank_dataframe(project_csv)
    prs = pu.create_blank_presentation(resource_path(const.FILE_LOCATIONS['pptx_template']))
    pu.create_New_slides(df, prs)
    #PROJECT UPDATES SECTION
        #SUMMARISE CURRENT PROJECT PROGRESS
        #HIGHLIGHT KEY MILESTONES
        #DISCUSS DEVIATIONS FROM PROJECT PLANS
    #PROJECT STATUS CHANGES
        #REVIEW AND APPROVE STATUS CHANGES
        #DISCUSS TIMELINE ADJUSTMENTS & BUDGET IMPACTS
    pu.save_exit(prs, "PEA_Project_Review_Call_Report", "", output_folder)

def engineering_presentation(project_csv, output_folder):
    count = 0
    for eng in const.ENGINEERS:
        person_filter(project_csv=project_csv, person=eng, output_folder=output_folder)
        count += 1
    for rtl in const.RTL:
        person_filter(project_csv=project_csv, person=rtl, output_folder=output_folder)
        count += 1
    return count

def impact_presentation(project_csv, impact_filter, output_folder):
    count = 0
    if isinstance(impact_filter, list):
        for filter_item in impact_filter:
            impact_slides(project_csv=project_csv, filter=filter_item, output_folder=output_folder)
            count += 1
    else:
        impact_slides(project_csv=project_csv, filter=impact_filter, output_folder=output_folder)
        count += 1
    return count

def allimpacted_presentation(project_csv, output_folder):
    allimpacted(project_csv=project_csv, output_folder=output_folder)

def who_presentation(project_csv, name_filter, output_folder):
    person_filter(project_csv=project_csv, person=name_filter, output_folder=output_folder)

def onhold_presentation(project_csv, output_folder):
    df = du.create_blank_dataframe(project_csv)
    prs = pu.create_blank_presentation(resource_path(const.FILE_LOCATIONS['pptx_template']))
    pu.create_OnHold_slides(df, prs, no_section=True)
    pu.save_exit(prs, "PEA_Project_Report", "_OnHold", output_folder)

def objective_presentation(project_csv, output_folder):
    objective(project_csv=project_csv, output_folder=output_folder)

def projects_presentation(project_csv, output_folder):
    df = du.create_blank_dataframe(project_csv)
    prs = pu.create_blank_presentation(resource_path(const.FILE_LOCATIONS['pptx_template']))
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
        prs = pu.create_blank_presentation(resource_path(const.FILE_LOCATIONS['pptx_template']))
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
            prs = pu.create_blank_presentation(resource_path(const.FILE_LOCATIONS['pptx_template']))
            pu.create_Impacted_section(df, prs, no_section=True, impacted_team=imp)
            pu.save_exit(prs, "PEA_Project_Report", "_"+imp, output_folder)
    else:
        prs = pu.create_blank_presentation(resource_path(const.FILE_LOCATIONS['pptx_template']))
        pu.create_Impacted_section(df, prs, no_section=True, impacted_team=filter)
        output_path = pu.save_exit(prs, "PEA_Project_Report", "_"+filter, output_folder)
    return output_path

def allimpacted(project_csv, output_folder, save=True, prs=None):
    df = du.create_blank_dataframe(project_csv)
    impacted = du.impacted_teams_list(df)
    if prs is None:
        prs = pu.create_blank_presentation(resource_path(const.FILE_LOCATIONS['pptx_template']))
    output_path = ""
    for imp in impacted:
        pu.create_Impacted_section(df, prs, no_section=True, impacted_team=imp)
    if save:
        output_path = pu.save_exit(prs, "PEA_Project_Report", "_AllImpacts", output_folder)
    return output_path

def objective(project_csv, output_folder, save=True, prs=None):
    df = du.create_blank_dataframe(project_csv)
    if prs is None:
        prs = pu.create_blank_presentation(resource_path(const.FILE_LOCATIONS['pptx_template']))
    pu.create_Objective_slides(df, prs)
    output_path = ""
    if save:
        output_path = pu.save_exit(prs, "PEA_Project_Report", "_Objective", output_folder)
    return output_path

def output_all(project_csv, output_folder, save=True, prs=None):
    df = du.create_blank_dataframe(project_csv)
    if prs is None:
        prs = pu.create_blank_presentation(resource_path(const.FILE_LOCATIONS['pptx_template']))
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
        prs = pu.create_blank_presentation(resource_path(const.FILE_LOCATIONS['pptx_template']))
    pu.create_document_release_section(df, prs, name_filter)
    output_path = ""
    if save:
        output_path = pu.save_exit(prs, "PEA_Project_Report", "_DocumentBoard", folder = output_folder)
    return output_path

def doc_changes(project_csv, output_folder, save=True, prs=None):
    df = du.create_blank_dataframe(project_csv)
    if prs is None:
        prs = pu.create_blank_presentation(resource_path(const.FILE_LOCATIONS['pptx_template']))
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
        prs = pu.create_blank_presentation(resource_path(const.FILE_LOCATIONS['pptx_template']))

    # Create the initial section for New / Updated docs this period
    pu.create_document_release_section(df, prs, filter, internal=internal)
    
    pu.create_document_release_section_commercial_impacts(df, prs, filter, internal=internal)
    
    output_path = ""
    
    # Create the section for docs that Impact specific areas this period
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
        prs = pu.create_blank_presentation(resource_path(const.FILE_LOCATIONS['pptx_template']))

    pu.create_document_release_section_multi_filter(df, prs, filter, internal=internal)
    output_path = ""
    
    if not internal:
        for imp in impacted:
            pu.create_document_Impacted_section_multi_filter(df, prs, no_section=False, impacted_team=imp, group_filter=filter)
    if save:
        output_path = pu.save_exit(prs, "PEA_Document_Release", save_tail, folder = output_folder)

    return output_path
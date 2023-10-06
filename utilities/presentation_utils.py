from pptx import Presentation
from pptx.util import Pt, Cm
from pptx.enum.text import MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from datetime import date, datetime
import pandas as pd
import math
import os

#import PresentationToolKit.utilities.constants as const 
#import PresentationToolKit.utilities.data_utils as du

import utilities.constants as const 
import utilities.data_utils as du

def create_blank_presentation(template: str = './templates/_template.pptx'):
    """
    Create a new PowerPoint presentation object based on a template.

    Args:
    template (str, optional): The path to the PowerPoint template file. Defaults to './templates/_template.pptx'.

    Returns:
    Presentation: A new Presentation object.
    """
    prs = Presentation(template)
    return prs

def generate_filename(modifier: str = "") -> str:
    """
    Generate a filename based on the current date, time, and an optional modifier.

    Args:
    modifier (str, optional): An additional string to append to the filename. Defaults to an empty string.

    Returns:
    str: Generated filename.
    """
    today = date.today()
    current_time = datetime.now().strftime("%H%M")  # Get current hour and minute
    filename = f'PEA_Project_Report_{today.strftime("%y%m%d")}_{current_time}{modifier}.pptx'
    return filename

def save_exit(prs, modifier: str = "", folder: str = "") -> str:
    """
    Save the PowerPoint presentation with a generated filename.

    Args:
    - prs : The presentation object to be saved.
    - modifier (str, optional): An additional string to append to the filename. Defaults to an empty string.
    - folder (str, optional): The folder where the presentation should be saved. Defaults to the current directory.

    Returns:
    str: The name of the saved PowerPoint file.
    """
    filename = generate_filename(modifier)
    save_to_location = os.path.join(folder, filename)
    prs.save(save_to_location)
    return filename

def create_title_slide(prs, title=""):
    """
    Create a title slide and add it to the given presentation.

    Args:
    prs (Presentation): The presentation object to which the slide will be added.
    title (str, optional): The title to set for the slide. Defaults to an empty string.

    Returns:
    Slide: The created title slide.
    """
    SLIDE_MASTER_INDEX = 1
    BLANK_SLIDE_LAYOUT_INDEX = 3

    # Add a new slide with the specified layout
    section_slide = prs.slides.add_slide(prs.slide_masters[SLIDE_MASTER_INDEX].slide_layouts[BLANK_SLIDE_LAYOUT_INDEX])
    
    #Populate Title Placeholder on slide
    set_slide_title(section_slide, title)
        
    return section_slide

def set_slide_title(slide, title_text=""):
    """
    Set the title of a given slide.

    Args:
    slide (Slide): The slide object whose title needs to be set.
    title_text (str, optional): The text to set as the title. Defaults to an empty string.
    """
    # Access the title shape of the slide
    title_shape = slide.shapes.title
    
    # Set the title text
    title_shape.text = title_text

def set_three_col_subtitle(slide, placeholder_indices=const.THREE_COL_PLACEHOLDER_INDICES, subtitle=const.THREE_COL_TITLES):
    """
    Set the subtitles in three columns on a given slide.

    Args:
    slide (Slide): The slide object whose subtitles need to be set.
    placeholder_indices (dict, optional): A dictionary containing the placeholder indices for each column.
                                         Defaults to predefined indices (hardcoded idx values that will be consistent for prs.slide_masters[1].slide_layouts[6] only)
    subtitle (dict, optional): A dictionary containing the title text for each column.
                                         Defaults to predefined values.                                                                                  
    """
    # Load standard column titles from external constant
    subtitle = const.THREE_COL_TITLES
    
    # Set the text in placeholders based on the provided indices
    for col, idx in placeholder_indices.items():
        slide.placeholders[idx].text = subtitle[col]

def set_document_release_subtitle(slide, heading='new', date='08/09/2023', placeholder_indices=const.DOC_RELEASE_PLACEHOLDER_INDICES, subtitle=const.DOC_BOARD_TITLES):
    """
    Set the subtitles for a document release slide.

    Args:
    slide (Slide): The slide object whose subtitles need to be set.
    heading (str, optional): The heading to set for the slide. Defaults to 'new'.
    date (str, optional): The date to set for the slide. Defaults to '08/09/2023'.
    placeholder_indices (dict, optional): A dictionary containing the placeholder indices for heading and date.
                                          Defaults to predefined indices.
    subtitle (dict, optional): A dictionary containing the subtitle headings.
                                          Defaults to predefined indices.
    """
    
    # Set the text in placeholders based on the provided indices
    slide.placeholders[placeholder_indices['heading']].text = subtitle[heading]
    slide.placeholders[placeholder_indices['date']].text = date

# Function to create rounded rectangle shape
def create_rounded_rectangle(slide, left, top, width, height):
    return slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)

# Function to set shape appearance
def set_shape_appearance(shape, fill_color, border_color):
    fill = shape.fill
    fill.solid()
    fill.fore_color.theme_color = fill_color

    line = shape.line
    line.color.theme_color = border_color
    line.color.brightness = -0.50 

    #set shadows to nil
    shadow = shape.shadow
    shadow.inherit = False
    shadow.blur_radius = Cm(0)
    shadow.distance = Cm(0)
    shadow.angle = 0
    shadow.alpha = 0

def set_shape_text_with_bold_first_run(shape, text, font_size, font_color):
    text_box = shape.text_frame
    text_box.text = text
    text_box.vertical_anchor = MSO_ANCHOR.TOP

    first = True
    for paragraph in text_box.paragraphs:
        for run in paragraph.runs:
            run.font.color.theme_color = font_color
            run.font.size = font_size
            if first:
                run.font.bold = True
                first = False

# Refactored version of create_project_button function
def create_project_button(slide, left, top, status="", contents_text="CONTENT", properties=const.PROJECT_BUTTON_CONSTANTS):
    """
    Create a project button on the given slide at the specified position.

    Args:
    slide (Slide): The slide object where the button will be created.
    left (float): The x-coordinate for the button.
    top (float): The y-coordinate for the button.
    status (str, optional): The status of the project. Defaults to an empty string.
    contents_text (str, optional): The text to display on the button. Defaults to "CONTENT".
    properties (dict, optional): A dictionary containing the properties for the button. Defaults to DEFAULT_BUTTON_PROPERTIES.
    """
    # Create a rounded rectangle shape
    rounded_rectangle = create_rounded_rectangle(slide, left, top, properties['rectangle_width'], properties['rectangle_height'])
    
    # Pull through separate status colours & set shape appearance
    if status == "":
        fill_color = properties.get('fill', const.ThemeColors.PINK)
        font_color = properties.get('font_colour', const.ThemeColors.WHITE)
        border_color = properties.get('border', const.ThemeColors.PINK)
    else:
        fill_color = properties.get('fill', const.PROJECT_BUTTON_CONSTANTS['fill'])
        font_color = 0
        border_color = properties.get('border', const.PROJECT_BUTTON_CONSTANTS['border'])

    set_shape_appearance(rounded_rectangle, fill_color, border_color)
    
    # Set text in shape
    font_size = properties.get('font_size', const.PROJECT_BUTTON_CONSTANTS['font_size'])
    font_color = properties.get('font_color', const.PROJECT_BUTTON_CONSTANTS['font_color'])
    set_shape_text_with_bold_first_run(rounded_rectangle, contents_text, font_size, font_color)

def row_calculator(index, col=1):
    """
    Calculate the row index based on the given index and column type.

    Args:
    index (int): The index to be calculated.
    col (int, optional): The column type to divide the index by. Defaults to 1.

    Returns:
    int: The floor value of the division result.
    """
    result = index / col
    row = math.floor(result)
    return row

# Helper function to set contents text based on type_flag
def get_contents_text_by_type(type_flag, row):
    """
    Get contents text for a button based on the type_flag and DataFrame row.

    Args:
    type_flag (str): The type flag to determine the contents text.
    row (pd.Series): A row from the DataFrame containing project details.

    Returns:
    str: The contents text for the button.
    """
    
    # Add project details to the rectangle (based on "type_flag" for specific contents)
    if type_flag == 'ProjectOwner':
        contents_text = f"{row['Title']}\n"
        contents_text += f"Objective: {row['Objective']}\n"
        contents_text += f"Staging: {const.get_staging_text(row['Staging'])}\n"
        contents_text += f"Priority: {const.get_priority_text(row['Priority'])}"
        if row['Status'] == 'Blocked':
            contents_text += f"\nBlocked: {row['Closure Comments']}"
        return contents_text
    
    elif type_flag == 'Objective':
        contents_text = f"{row['Title']}\n"
        contents_text += f"Owner: {row['Primary Owner']}\n"
        contents_text += f"Staging: {const.get_staging_text(row['Staging'])}\n"
        contents_text += f"{const.get_priority_text(row['Priority'])}"
        return contents_text
    
    elif type_flag == 'Impact':
        contents_text = f"{row['Title']}\n"
        contents_text += f"Staging: {const.get_staging_text(row['Staging'])}"
        return contents_text
    
    elif type_flag == 'OnHold':
        contents_text = f"{row['Title']}\n"
        contents_text += f"Owner: {row['Primary Owner']}\n"
        contents_text += f"Staging: {const.get_staging_text(row['Staging'])}\n"
        contents_text += "Project Summary: "
        if not pd.isna(row['Project Summary']):
            contents_text += f"{row['Project Summary']}"
        return contents_text
    
    else:
        return "UNKNOWN INPUT TYPE"
    
def calculate_position(col, index, COLUMN_FORMAT, SLIDE_FORMAT, BUTTON_FORMAT):
    """
    Calculate the position for placing a button based on the row and column.

    Args:
    row (int): The row index.
    col (int): The column index.
    index (int): The current item's index
    BUTTON_FORMAT (dict): The format for the buttons.
    SLIDE_FORMAT (dict): The format for the slide.
    COLUMN_FORMAT (dict): The format for the columns.

    Returns:
    tuple: A tuple containing the left, top positions.
    """
    # Calculate position for the current rectangle
    left = COLUMN_FORMAT['left']  

    row = row_calculator(index, col)

    if 'right' in COLUMN_FORMAT:
        right = COLUMN_FORMAT['right']
    else:
        right = left
    
    if not index % 2 == 0:
        left = right
        
    top = SLIDE_FORMAT['start_top'] + row * (BUTTON_FORMAT['rectangle_height'] + SLIDE_FORMAT['vertical_spacing'])

    return left, top

# Refactored version of populate_column function
def populate_column(df, slide, BUTTON_FORMAT, SLIDE_FORMAT, COLUMN_FORMAT, type_flag='ProjectOwner', col=1):
    """
    Populate a column in the slide with project details from a DataFrame.

    Args:
    df (pd.DataFrame): The DataFrame containing project details.
    slide (Slide): The slide object to be populated.
    BUTTON_FORMAT (dict): The format for the buttons.
    SLIDE_FORMAT (dict): The format for the slide.
    COLUMN_FORMAT (dict): The format for the columns.
    type_flag (str, optional): The type flag to determine the contents text. Defaults to 'ProjectOwner'.
    col (int, optional): The column number. Defaults to 1.
    """
    for index, row in df.iterrows():
        # Calculate position for the button
        left, top = calculate_position(col, index, COLUMN_FORMAT, SLIDE_FORMAT, BUTTON_FORMAT)
       
        # Get contents text for the button
        contents_text = get_contents_text_by_type(type_flag, row)
        
        # Create the project button on the slide
        create_project_button(slide, left, top, status=row['status'], contents_text=contents_text, properties=BUTTON_FORMAT)

def create_body_slide_three_cols(df_1, df_2, df_3, prs, type_flag='ProjectOwner', title_text="", BUTTON_OVERRIDE=""):
    # Function to take contents of df (dataframe) and output onto a 3 column grid using pre-sets from constants.py
    # Set grid parameters
    columns = 3  # Number of columns in the grid

    # Create a new slide
    slide = prs.slides.add_slide(prs.slide_masters[1].slide_layouts[6])  # Blank slide layout
    set_slide_title(slide, title_text)
    set_three_col_subtitle(slide)

    # Set up Constants
    SLIDE_DEF = const.THREE_COL_SLIDE_CONSTANTS

    #Hardcoding in the Column1 Button Dimension - fix in future revision
    COL1_BUTTON_DEF = const.THREE_COL_PROJECT_BUTTON_COL1_CONSTANTS

    if BUTTON_OVERRIDE == "":
        BUTTON_DEF = const.THREE_COL_PROJECT_BUTTON_CONSTANTS
    else:
        BUTTON_DEF = BUTTON_OVERRIDE

    COLUMN_1 = {
        'left': SLIDE_DEF['start_left_col1'],
        'right': SLIDE_DEF['start_left_col2']
    }
    COLUMN_2 = {
        'left': SLIDE_DEF['start_left_col3']
    }
    COLUMN_3 = {
        'left': SLIDE_DEF['start_left_col4']
    }
        # Iterate through each project and add a rounded rectangle (COLUMN 1)
    populate_column(df_1, slide, COL1_BUTTON_DEF, SLIDE_DEF, COLUMN_1, type_flag, col=2)
    populate_column(df_2, slide, BUTTON_DEF, SLIDE_DEF, COLUMN_2, type_flag)
    populate_column(df_3, slide, BUTTON_DEF, SLIDE_DEF, COLUMN_3, type_flag)

def create_body_slide_four_cols(df, prs, type_flag='ProjectOwner', title_text="", BUTTON_OVERRIDE=""):
    # Function to take contents of df (dataframe) and output onto a 4 column grid using pre-sets from constants.py
    # Set grid parameters
    columns = 4  # Number of columns in the grid

    # Create a new slide
    slide = prs.slides.add_slide(prs.slide_masters[1].slide_layouts[5])  # Blank slide layout
    set_slide_title(slide, title_text)

    # Set up Constants
    SLIDE_DEF = const.FOUR_COL_SLIDE_CONSTANTS

    if BUTTON_OVERRIDE == "":
        BUTTON_DEF = const.PROJECT_BUTTON_CONSTANTS
    else:
        BUTTON_DEF = BUTTON_OVERRIDE

        # Iterate through each project and add a rounded rectangle
    for index, (_, project) in enumerate(df.iterrows()):
        # Calculate current row and column
        row = index // columns
        column = index % columns
        status = project['Status']

        # Calculate position for the current rectangle
        left = SLIDE_DEF['start_left'] + column * (BUTTON_DEF['rectangle_width'] + SLIDE_DEF['horizontal_spacing'])
        top = SLIDE_DEF['start_top'] + row * (BUTTON_DEF['rectangle_height'] + SLIDE_DEF['vertical_spacing'])

        # Add project details to the rectangle (based on "type_flag" for specific contents)
        if type_flag == 'ProjectOwner':
            contents_text = f"{project['Title']}\n"
            contents_text += f"Objective: {project['Objective']}\n"
            contents_text += f"Staging: {const.get_staging_text(project['Staging'])}\n"
            contents_text += f"Priority: {const.get_priority_text(project['Priority'])}"
            if project['Status'] == 'Blocked':
                contents_text += f"\nBlocked: {project['Closure Comments']}"

        if type_flag == 'Objective':
            contents_text = f"{project['Title']}\n"
            contents_text += f"Owner: {project['Primary Owner']}\n"
            contents_text += f"Staging: {const.get_staging_text(project['Staging'])}\n"
            contents_text += f"{const.get_priority_text(project['Priority'])}"

        if type_flag == 'Impact':
            contents_text = f"{project['Title']}\n"
            contents_text += f"Owner: {project['Primary Owner']}\n"
            contents_text += f"Staging: {const.get_staging_text(project['Staging'])}\n"
            contents_text += f"{const.get_priority_text(project['Priority'])}"

        if type_flag == 'OnHold':
            contents_text = f"{project['Title']}\n"
            contents_text += f"Owner: {project['Primary Owner']}\n"
            contents_text += f"Staging: {const.get_staging_text(project['Staging'])}\n"
            contents_text += f"Project Summary: {project['Project Summary']}"

        if type_flag == 'Release Forecast':
            contents_text = f"Document: {project['Doc Reference']}\n"
            contents_text += f"Title: {project['Title']}\n"
            contents_text += f"Owner: {project['Primary Owner']}"

        create_project_button(slide, left, top, status, contents_text, OVERRIDE=BUTTON_DEF)

def create_OnHold_slides(df, prs, no_section=False):
    if no_section == False:
        create_title_slide(prs, f'On Hold Projects')
    #Filter the dataframe to only the OnHold projects
    on_hold = du.filter_dataframe_by_status(df, 'On Hold')
    sorted = on_hold.sort_values(by=['Primary Owner'])
    title_text = "On-Hold Projects - " + str(len(on_hold))
    create_body_slide_four_cols(sorted, prs, type_flag='OnHold', title_text=title_text, BUTTON_OVERRIDE=const.ONHOLD_BUTTON_CONSTANTS)

def create_ProjectOwner_slides(df, prs, filter=""):
    
    # Group projects by Project Owner
    grouped = df.groupby('Primary Owner')

    # Iterate through each Project Owner and their projects
    for owner, projects in grouped:
        if not filter == "":
            if filter and filter.lower() not in owner.lower():
                continue  # Skip this owner if the filter doesn't match

        title_text = owner + " - " + str(len(projects))
        sorted_projects = projects.sort_values(by=['Priority', 'Objective'])
        create_body_slide_four_cols(sorted_projects, prs, 'ProjectOwner', title_text)

def create_Objective_slides(df, prs, filter=""):

    create_title_slide(prs, 'By Objective')

    if not filter == "":
        df = df[df['Primary Owner'].str.contains(filter, case=False, na=False)]


    # Group projects by Objective
    grouped = df.groupby('Objective')

    # Iterate through each Objective and their projects
    for objective, projects in grouped:
        
        title_text = objective  + " - " + str(len(projects))
        sorted_projects = projects.sort_values(by=['Priority'])
        create_body_slide_four_cols(sorted_projects, prs, 'Objective', title_text)

def create_Impacted_section(df, prs, no_section=False, impacted_team='Training'):
    if no_section == False:
        create_title_slide(prs, f'Projects Impacting {impacted_team}')

    # Filter projects by Impacted Team
    projects = du.filter_dataframe_by_team(df, impacted_team)
    sorted_projects = projects.sort_values(by=['Priority'])

    title_text = impacted_team  + " - " + str(len(projects))
#    create_body_slide_four_cols(sorted_projects, prs, 'Impact', title_text)

    df_triage = du.filter_dataframe_by_staging(projects, 'Triage')
    df_analysis = du.filter_dataframe_by_staging(projects, 'Analysis')
    df_alpha = du.filter_dataframe_by_staging(projects, 'Alpha Test')
    df_beta = du.filter_dataframe_by_staging(projects, 'Beta Test')
    df_rollout = du.filter_dataframe_by_staging(projects, 'Roll-out')

    col1 = du.combine_dataframe(df_triage, df_analysis)
    col2 = du.combine_dataframe(df_alpha, df_beta)
    col3 = df_rollout

    create_body_slide_three_cols(col1, col2, col3, prs, 'Impact', title_text)

def create_project_section(df, prs, no_section=False):
    if no_section == False:
        create_title_slide(prs, f'Project Details')
    
    #DF Filtering Logic goes here

    for index, (_, project) in enumerate(df.iterrows()):
        create_single_project_slide(project, prs)
    return

def create_single_project_slide(project, prs, title_text=""):

    update = project['Project Updates']
    if pd.isna(update):
       update = " " 
    else:
        update = du.convert_html_to_text_with_newlines(update)

    action = project['Project Actions']
    if pd.isna(action):
       action = " " 
    else:
        action = du.convert_html_to_text_with_newlines(action)

    summary = project['Project Summary']
    if pd.isna(summary):
       summary = " " 

    slide = prs.slides.add_slide(prs.slide_masters[1].slide_layouts[7])  # Blank slide layout
    
    set_slide_title(slide, project['Title'])

    #Slide subtitle
    slide.placeholders[10].text = f"Project Summary: {summary}"

    #Column Subtitles
    slide.placeholders[26].text = "Project Updates" #Left
    slide.placeholders[28].text = "Project Actions" #Right

    #Column Text
    slide.placeholders[24].text = f"{update}" #Left
    slide.placeholders[27].text = f"{action}" #Right

    return

def create_document_release_section(df, prs, filter=''):
    
    if filter:
        df = df.loc[df['Release Forecast'] == filter]

    grouped = df.groupby(df['Release Forecast'].fillna('None'))

    for date, documents in grouped:
        title_text = 'Technical Releases'

        # Filtering based on 'Doc Reference'
        new_docs = documents[documents['Doc Reference'].str.endswith('NEW')]
        update_docs = documents[~documents['Doc Reference'].str.endswith('NEW')]

        sorted_new_docs = new_docs.sort_values(by=['Doc Reference'])
        sorted_update_docs = update_docs.sort_values(by=['Doc Reference'])

        create_document_release_slide(sorted_new_docs, prs, date, title_text, const.DOCUMENT_BUTTON_CONSTANTS, 'new')
        create_document_release_slide(sorted_update_docs, prs, date, title_text, const.DOCUMENT_BUTTON_CONSTANTS, 'update')
    return

def create_document_changes_section(df, prs, filter=''):
    
    if filter:
        df = df.loc[df['Release Forecast'] == filter]

    grouped = df.groupby(df['Doc Reference'].fillna('None'))

    for doc, changes in grouped:
        title_text = df['DocReference']

        sorted_changes = changes.sort_values(by=['Release Forecast'].fillna('None'))

        create_document_release_slide(sorted_changes, prs, doc, title_text, const.DOCUMENT_BUTTON_CONSTANTS, 'new')
    return

def create_document_release_slide(df, prs, date='08/09/2023', title_text="", BUTTON_OVERRIDE="", type_flag='new'):
    # Function to take contents of df (dataframe) and output onto a 2 column grid using pre-sets from constants.py for the Document Release Board

    columns = 2

    # Create a new slide
    slide = prs.slides.add_slide(prs.slide_masters[1].slide_layouts[7])  # Blank slide layout
    set_slide_title(slide, title_text)
    set_document_release_subtitle(slide, type_flag, date)

    # Set up Constants
    SLIDE_DEF = const.DOC_RELEASE_SLIDE_CONSTANTS

    if BUTTON_OVERRIDE == "":
        BUTTON_DEF = const.DOCUMENT_BUTTON_CONSTANTS
    else:
        BUTTON_DEF = BUTTON_OVERRIDE

        # Iterate through each project and add a rounded rectangle
    for index, (_, project) in enumerate(df.iterrows()):
        # Calculate current row and column
        row = index // columns
        column = index % columns
        status = ""

        # Calculate position for the current rectangle
        left = SLIDE_DEF['start_left'] + column * (BUTTON_DEF['rectangle_width'] + SLIDE_DEF['horizontal_spacing'])
        top = SLIDE_DEF['start_top'] + row * (BUTTON_DEF['rectangle_height'] + SLIDE_DEF['vertical_spacing'])

        # Add project details to the rectangle (based on "type_flag" for specific contents)
        if type_flag == 'new':
            contents_text = f"Document: {project['Doc Reference']}\n"
            contents_text += f"Title: {project['Title']}\n"
            contents_text += f"Changes: "

        if type_flag == 'update':
            contents_text = f"Document: {project['Doc Reference']}\n"
            contents_text += f"Title: {project['Title']}\n"
            contents_text += f"Changes: "

        create_project_button(slide, left, top, status, contents_text, OVERRIDE=BUTTON_DEF)

    return

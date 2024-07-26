from pptx.util import Pt, Cm
from pptx.enum.dml import MSO_THEME_COLOR

ENGINEERS = ['Andy Oxford', 'Chris Kelly', 'Luke Phillips', 'Matthew Harbord', 'Neil Griffin', 'Gordon Pyrah', 'Tom Wright']
RTL = ['Irina Gonik', 'Peter Little', 'Craig McNamara']

# Theme Colors
class ThemeColors:
    PINK = MSO_THEME_COLOR.ACCENT_5
    GREEN = MSO_THEME_COLOR.ACCENT_4
    BLUE = MSO_THEME_COLOR.ACCENT_3
    PURPLE = MSO_THEME_COLOR.ACCENT_6
    ORANGE = MSO_THEME_COLOR.ACCENT_2
    TEAL = MSO_THEME_COLOR.ACCENT_1
    BLACK = MSO_THEME_COLOR.TEXT_1
    WHITE = MSO_THEME_COLOR.BACKGROUND_1

STATUS_COLOUR = {
    'Open': ThemeColors.PINK,
    'On Hold': ThemeColors.ORANGE,
    'New': ThemeColors.PURPLE,
    'Blocked': ThemeColors.ORANGE,
}

# File Locations
FILE_LOCATIONS = {
    'project_csv': './raw/PROJECT.csv',
    'document_csv': './raw/DOC.csv',
    'contact_csv': './raw/CONTACT.csv',
    'concession_csv': './raw/CONCESSION.csv',
    'pptx_template': './templates/_template.pptx',
    'output_folder': './output/',
}
    
# Text Representations for Staging
STAGING_TEXT_REPRESENTATION = {
    'Triage': "▰▱▱▱▱",
    'Analysis': "▰▰▱▱▱",
    'Alpha Test': "▰▰▰▱▱",
    'Beta Test': "▰▰▰▰▱",
    'Roll-out': "▰▰▰▰▰",
}

# Text Representations for Priority
PRIORITY_TEXT_REPRESENTATION = {
    'P1': "P1 🔥",
    'P1 🔥': "P1 🔥",
    'P2': "P2 🚨",
    'P3': "P3 ⭐",
    'P4': "P4 🐢",
    'P5': "P5 🐌",
}

# Project Button Constants
PROJECT_BUTTON_CONSTANTS = {
    'rectangle_width': Cm(7.8),
    'rectangle_height': Cm(2.8),
    'font_size': Pt(11),
    'fill': ThemeColors.PINK,
    'font_colour': ThemeColors.WHITE,
    'border': ThemeColors.PINK,
    'fill_brightness': 0,
    'status_colors': {
        'Open': ThemeColors.PINK,
        'On Hold': ThemeColors.ORANGE,
        'New': ThemeColors.PURPLE,
        'Blocked': ThemeColors.ORANGE
    }
}

ALL_PROJECTS_BUTTON_CONSTANTS = {
    'rectangle_width': Cm(7.8),
    'rectangle_height': Cm(0.8),
    'font_size': Pt(11),
    'fill': ThemeColors.PINK,
    'font_colour': ThemeColors.WHITE,
    'border': ThemeColors.PINK,
    'fill_brightness': 0,
    'status_colors': {
        'Open': ThemeColors.PINK,
        'On Hold': ThemeColors.ORANGE,
        'New': ThemeColors.PURPLE,
        'Blocked': ThemeColors.ORANGE
    }
}

# Document Button Constants
DOCUMENT_BUTTON_CONSTANTS = {
    'rectangle_width': Cm(15),
    'rectangle_height': Cm(2.3),
    'font_size': Pt(11),
    'fill': ThemeColors.WHITE,
    'font_colour': ThemeColors.BLACK,
    'border': ThemeColors.PINK,
    'fill_brightness': 0.8,
    'status_colors': {
        'New': ThemeColors.GREEN,
        'In Progress': ThemeColors.ORANGE,
        'Internal Review': ThemeColors.BLUE,
        'External Review': ThemeColors.BLUE,
        'Ready to Release': ThemeColors.PINK,
        'Internal Draft Review': ThemeColors.BLUE, # new internal review status
        'PDF Final Review': ThemeColors.BLUE, # new external review status
    }
}

# Project Button Constants
ONHOLD_BUTTON_CONSTANTS = {
    'rectangle_width': Cm(7.8),
    'rectangle_height': Cm(4.4),
    'font_size': Pt(11),
    'fill': ThemeColors.ORANGE,
    'font_colour': ThemeColors.WHITE,
    'border': ThemeColors.ORANGE,
    'fill_brightness': 0,
    'status_colors': {
        'Open': ThemeColors.PINK,
        'On Hold': ThemeColors.ORANGE,
        'New': ThemeColors.PURPLE,
        'Blocked': ThemeColors.ORANGE,
    }
}

FOUR_COL_SLIDE_CONSTANTS = {
    'start_left': Cm(0.65),
    'start_top': Cm(2),
    'horizontal_spacing': Cm(0.2),
    'vertical_spacing': Cm(0.2),
}

# Project Button Constants
THREE_COL_PROJECT_BUTTON_CONSTANTS = {
    'rectangle_width': Cm(8.6),
    'rectangle_height': Cm(2),
    'font_size': Pt(10),
    'fill': ThemeColors.ORANGE,
    'font_colour': ThemeColors.WHITE,
    'border': ThemeColors.ORANGE,
    'fill_brightness': 0,
    'status_colors': {
        'Open': ThemeColors.PINK,
        'On Hold': ThemeColors.ORANGE,
        'New': ThemeColors.PURPLE,
        'Blocked': ThemeColors.ORANGE,
    }
}

THREE_COL_PROJECT_BUTTON_COL1_CONSTANTS = {
    'rectangle_width': Cm(6.27),
    'rectangle_height': Cm(2),
    'font_size': Pt(10),
    'fill': ThemeColors.PINK,
    'font_colour': ThemeColors.WHITE,
    'border': ThemeColors.PINK,
    'fill_brightness': 0,
    'status_colors': {
        'Open': ThemeColors.PINK,
        'On Hold': ThemeColors.ORANGE,
        'New': ThemeColors.PURPLE,
        'Blocked': ThemeColors.ORANGE,
    }
}

THREE_COL_SLIDE_CONSTANTS = {
    'start_left_col1': Cm(0.65),
    'start_left_col2': Cm(7.29),
    'start_left_col3': Cm(14.82),
    'start_left_col4': Cm(24.33),
    'start_top': Cm(3.39),
    'vertical_spacing': Cm(0.2)
}

DOC_RELEASE_SLIDE_CONSTANTS = {
    'start_left': Cm(0.65),
    'start_top': Cm(3.87),
    'vertical_spacing': Cm(0.2),
    'horizontal_spacing': Cm(0.2),
}

THREE_COL_TITLES = {
    'col1': 'Assessing:',
    'col2': 'Testing:',
    'col3': 'Preparing Release:'
}

DOC_BOARD_TITLES = {
    'main': 'Technical Releases',
    'new': 'New Documents',
    'update': 'Updated Documents',
    'changes': ' ',
    'release_impact': 'Changes to Review:',
    'urgency_quarterly': 'Quarterly Releases',
    'urgency_monthly': 'Monthly Releases',
}

IMPACT_SYMBOL_MAPPING = {
    "BDUK": "A",
    "New Product": "B",
    "Change of Method": "C",
    "Change to Time Taken": "D",
    "Change to Materials": "E",
}

# Functions
def get_staging_text(staging):
    return STAGING_TEXT_REPRESENTATION.get(staging, "-----")

def get_priority_text(priority):
    return PRIORITY_TEXT_REPRESENTATION.get(priority, "unknown")
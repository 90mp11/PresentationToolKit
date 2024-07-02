# Project Reporting Tool

## Description

The Project Reporting Tool is developed by the Passive Engineering Team to automate the generation of various types of presentations and reports based on project data. For specific help, please contact Matt Proctor.

## Usage

The script can be run from the command line with various options to generate different types of presentations. Below are the available command-line arguments and their descriptions.

## Command-Line Arguments

- `--engineering`: Create a batch of presentations, one for each listed team member in the `constants.py` file.
- `--impact`: Create individual presentations for each of the unique teams listed in the 'Impacted Teams' column.
- `--allimpacted`: Create a single presentation to contain impacts against each of the unique teams listed in the 'Impacted Teams' column.
- `--who`: Creates a single presentation filtered based on user input to the 'Primary Owner' field. This works on incomplete names.
- `--debug`: Used for testing/debugging purposes.
- `--onhold`: Exports all the on-hold projects to a single presentation.
- `--objective`: Exports the Objective view into a single presentation.
- `--projects`: Exports the Project view into a single presentation.
- `--docs`: Imports the Document Change Log and outputs the Monthly Release Board slides.
- `--document_changes`: Exports the Document Changes view into a single presentation.
- `--output_all`: Exports the Standard Output Report in a single presentation.
- `--release`: Outputs the ReleaseBoard Impact Report in a single presentation.
- `--internal`: Sets a flag for internal use only, outputs the ReleaseBoard Impact Report in a single presentation.

## Functions

### main()

This is the main function that parses command-line arguments and triggers the appropriate functions based on the provided arguments.

### person_filter()

Generates a presentation filtered by the given person's name.

**Parameters:**
- `project_csv`: Path to the project CSV file.
- `output_folder`: Path to the output folder.
- `person`: Name of the person to filter by.
- `save`: Boolean indicating whether to save the presentation.
- `prs`: Existing presentation object (if any).

### impact_slides()

Generates presentations for impacted teams.

**Parameters:**
- `project_csv`: Path to the project CSV file.
- `output_folder`: Path to the output folder.
- `filter`: Team to filter by.

### allimpacted()

Generates a single presentation for all impacted teams.

**Parameters:**
- `project_csv`: Path to the project CSV file.
- `output_folder`: Path to the output folder.
- `save`: Boolean indicating whether to save the presentation.
- `prs`: Existing presentation object (if any).

### objective()

Generates a presentation based on project objectives.

**Parameters:**
- `project_csv`: Path to the project CSV file.
- `output_folder`: Path to the output folder.
- `save`: Boolean indicating whether to save the presentation.
- `prs`: Existing presentation object (if any).

### output_all()

Generates a comprehensive report containing all projects.

**Parameters:**
- `project_csv`: Path to the project CSV file.
- `output_folder`: Path to the output folder.
- `save`: Boolean indicating whether to save the presentation.
- `prs`: Existing presentation object (if any).

### all_docs()

Generates a document release section based on a date filter.

**Parameters:**
- `project_csv`: Path to the document CSV file.
- `output_folder`: Path to the output folder.
- `name_filter`: Date filter for document releases.
- `save`: Boolean indicating whether to save the presentation.
- `prs`: Existing presentation object (if any).

### doc_changes()

Generates a presentation based on document changes.

**Parameters:**
- `project_csv`: Path to the document CSV file.
- `output_folder`: Path to the output folder.
- `save`: Boolean indicating whether to save the presentation.
- `prs`: Existing presentation object (if any).

### release_board_slides()

Generates a Release Board Impact Report.

**Parameters:**
- `project_csv`: Path to the document CSV file.
- `output_folder`: Path to the output folder.
- `filter`: Filter for the release group.
- `save`: Boolean indicating whether to save the presentation.
- `prs`: Existing presentation object (if any).
- `internal`: Boolean indicating whether the report is for internal use only.

### release_board_slides_multi_filter()

Generates a Release Board Impact Report with multiple filters.

**Parameters:**
- `project_csv`: Path to the document CSV file.
- `output_folder`: Path to the output folder.
- `filter`: List of filters for the release groups.
- `save`: Boolean indicating whether to save the presentation.
- `prs`: Existing presentation object (if any).
- `internal`: Boolean indicating whether the report is for internal use only.

### create_folders()

Creates the necessary folders for the project.

**Parameters:**
- `folder_names`: List of folder names to create.

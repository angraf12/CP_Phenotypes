Cerebral Palsy Gait Phenotype Classification Tool

Overview

The Cerebral Palsy Gait Phenotype Classification Tool is a Python-based application that leverages Tkinter to provide a user-friendly GUI for:
-Selecting an Excel file with gait analysis data.
-Choosing encounter dates for analysis.
-Classifying gait phenotypes using predefined decision trees.
-Exporting the classification results as an Excel file or PDF report.

Features
-Intuitive GUI: Easily load data, select dates, and run analyses.
-Data Validation: Ensures that the provided Excel file contains all the necessary columns.
-Customizable Date Selection: Allows you to select specific encounter dates to analyze.
-Automated Classification: Uses decision trees to determine sagittal and transversal gait deviations.
-Export Options: Save your results either as an Excel file for further analysis or as a formatted PDF report.

Prerequisites
-Python 3.x
-Required libraries:
  -pandas
  -tkinter (usually comes with Python)
  -numpy
  -Pillow (PIL)
  -reportlab
  -textwrap (standard library)

Installing Dependencies
Install the required packages using pip:
pip install pandas numpy Pillow reportlab
(Note: If tkinter is not installed on your system, please refer to your operating system’s documentation for installation instructions.)

Usage Instructions
1. Prepare Your Excel File
-Ensure your Excel file (.xlsx) contains the following columns (order is not important):
  -MotionParams_Side (e.g., "Left" or "Right")
  -EncounterAge
  -Date
  -MRN
  -AnkleDorsiPlantarPeakDFStance
  -KneeFlexionMeanStance
  -KneeFlexionTotalRange
  -HipFlexionMin
  -KneeFlexionMin
  -FootProgressionMeanStance
  -KneeRotationMeanStance
  -HipRotationMean
  -PelvicRotationMean

2. Launch the Application
  -Run the Python script (for example, python cp_gait_classification.py).

3. Load Your Data
-Click on the Select Excel file button.
-Browse to and select your prepared Excel file.
-The application will verify that all required columns are present.

4. Select Encounter Dates
-If prompted, choose the encounter dates you wish to analyze. You can select specific dates or choose "All" to process the entire dataset.

5. Run the Classification
-Click on the CP Gait Classification button.
-The tool will process the data and display the classification results in the application’s text area.

6. Export Your Results
-To save your results, click the Export Excel File button to create an Excel output.
-Alternatively, click the Export Report button to generate a PDF report.

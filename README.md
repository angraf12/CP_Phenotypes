# CP_Phenotypes
Automated Identification of Clinically Meaningful Biomechanical Phenotypes in Cerebral Palsy Using Gait Data

Instructions for Using the Cerebral Palsy Gait Phenotype Classification Tool:
Welcome to the Cerebral Palsy Gait Phenotype Classification tool. This tool requires an Excel file with specific column headers to function correctly however the order of the columns does not matter. Please ensure your Excel file contains the following columns:

MotionParams_Side: The side of the motion parameters (e.g., Left or Right).
EncounterAge: The age of the patient at the time of the encounter (e.g., 12).
Date: The date of the encounter (e.g., 2024-06-13).
MRN: The Medical Record Number of the patient (e.g., 123456).
AnkleDorsiPlantarPeakDFStance: The peak dorsiflexion angle of the ankle during stance (e.g., 15.2).
KneeFlexionMeanStance: The mean knee flexion angle during stance (e.g., 20.5).
KneeFlexionTotalRange: The total range of knee flexion in degrees (e.g., 30.4).
HipFlexionMin: The minimum hip flexion angle (e.g., 10.3).
KneeFlexionMin: The minimum knee flexion angle (e.g., 5.1).
FootProgressionMeanStance: The mean foot progression angle during stance (e.g., 7.8).
KneeRotationMeanStance: The mean rotation of the knee during stance in degrees (e.g., 12.0).
HipRotationMean: The mean rotation of the hip in degrees (e.g., 9.4).
PelvicRotationMean: The mean rotation of the pelvis in degrees (e.g., 4.2).

Steps to Use the Tool:
1.)Prepare Your Excel File:
  Ensure your Excel file (.xlsx) contains the columns listed above.
  Each column should be labeled exactly as specified to ensure compatibility with the tool.
2.)Launch the Application:
  Open the application by running the script.
3.)Select the Excel File:
  Click on the "Select Excel file" button.
  Browse and select your prepared Excel file.
4.)Run the Classification:
  Once the file is selected, the tool will check if all required columns are present.
  If any columns are missing, an error message will appear specifying the missing columns.
  If all required columns are present, you can proceed to run the classification by clicking the       "CP Gait Classification" button.
5.)Export Results:
  After the classification process, you can export the results.
  Use the "Export Excel file" button to save the results in an Excel file.
  Use the "Export Report" button to save the results in a PDF report

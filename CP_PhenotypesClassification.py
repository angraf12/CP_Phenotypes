"""
Cerebral Palsy Gait Phenotype Classification GUI Application
----------------------------------------------------------------
This application uses Tkinter to provide a graphical interface for:
- Selecting an Excel file containing gait analysis data.
- Choosing encounter dates from the data.
- Classifying gait phenotypes based on predefined decision trees.
- Exporting the classification results as an Excel file or PDF report.

Usage:
1. Run the script.
2. Click the "Select Excel file" button to load your data.
3. Follow on-screen prompts to select encounter dates.
4. Click on "CP Gait Classification" to generate results.
5. Use the export buttons to save the results as Excel or PDF.

Requirements:
- Python 3.x
- Required libraries: os, sys, pandas, tkinter, textwrap, numpy, PIL, reportlab
"""

# Import necessary libraries
import os 
import sys
import pandas as pd 
import tkinter as tk
import textwrap
import numpy as np
import PIL

# Import specific components from tkinter and reportlab
from tkinter import messagebox, filedialog, StringVar, END, Scrollbar, HORIZONTAL
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from tkinter import *  # Note: Importing everything from tkinter (use with caution)

# Print the version of the PIL module to the console (for debugging purposes)
print(PIL.__version__)

# ---------------------- Setup Main Application Window ----------------------
root = Tk()  # Create the main window
root.title("Cerebral Palsy Gait Phenotype Classification")
root.geometry("1400x900+280+60")
root.resizable(False, False)

# ----------------------------- Global Variables -----------------------------
folder_location = StringVar()  # Holds the path string for file saving
selected_file_path = None      # Will store the path of the selected Excel file

# --------------------------- Function Definitions ---------------------------

def select_dates(df):
    """
    Opens a pop-up window for the user to select encounter dates from the DataFrame.
    
    Parameters:
        df (pandas.DataFrame): The loaded data containing a 'Date' column.
        
    Returns:
        list: The list of selected dates.
    """
    popup1 = tk.Toplevel()
    popup1.title("Select Dates")
    
    # Get unique dates from the DataFrame and calculate window height based on count
    unique_dates = list(df['Date'].apply(lambda x: x.date()).unique())
    count_dates = len(unique_dates)
    popup1.geometry("360x" + str(((count_dates+1)*20)+150) + "+850+450")
    
    # Global list to store selected dates
    global selected_dates 
    selected_dates = []
    
    # Title label for the pop-up window
    title_label = tk.Label(popup1, text="Select the Encounter Date(s) \nfor which the recommendations are required")
    title_label.grid(row=0, column=0, sticky=NSEW)
    
    # "All" option checkbutton
    var_all = BooleanVar()
    cb_all = Checkbutton(popup1, text="All", variable=var_all, onvalue=True, offvalue=False)
    cb_all.grid(row=1, column=0, sticky=W)
    
    # Create a checkbutton for each unique date
    for i, date in enumerate(unique_dates):
        var = BooleanVar()
        cb = Checkbutton(popup1, text=date, variable=var, onvalue=True, offvalue=False)
        cb.grid(row=i+2, column=0, sticky=W)
        
        # Toggle function to add/remove dates from selected_dates list
        def toggle_date(date=date, var=var):
            if var.get():
                if date not in selected_dates:
                    selected_dates.append(date)
            else:
                if date in selected_dates:
                    selected_dates.remove(date)
        cb.config(command=toggle_date)
    
    def save_dates():
        """
        Saves the selected dates and closes the pop-up window.
        If "All" is selected, all dates are used.
        """
        global selected_dates
        if var_all.get():
            selected_dates = unique_dates
        popup1.destroy()
    
    # "Continue" button to save the selection and close the pop-up
    button_save = Button(popup1, text="Continue", padx=20, pady=10, command=save_dates)
    button_save.grid(row=2, column=1)
    
    # Make sure the pop-up window stays on top until closed
    popup1.grab_set()
    popup1.wait_window()
    
    return selected_dates

def browse_file():
    """
    Opens a file dialog for the user to select an Excel file.
    Checks if the file has the required columns and updates the GUI labels accordingly.
    """
    global error_label, ins_label, file_directory, selected_file_path
    selected_file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if selected_file_path == '':
        messagebox.showerror("Error", "Please choose a file or type an MRN number")
        return
    file_directory = os.path.dirname(selected_file_path)  # Get the directory of the selected file
    
    # Remove previous error/instruction labels if they exist
    if error_label:
        error_label.destroy()
    if ins_label:
        ins_label.destroy()
    
    # Load the Excel file into a DataFrame
    df = pd.read_excel(selected_file_path)
    
    # Check if the file contains all necessary columns
    required_columns = [
        "MotionParams_Side", "EncounterAge", "Date", "MRN",
        "AnkleDorsiPlantarPeakDFStance", "KneeFlexionMeanStance", "KneeFlexionTotalRange",
        "HipFlexionMin", "KneeFlexionMin", "FootProgressionMeanStance",
        "KneeRotationMeanStance", "HipRotationMean", "PelvicRotationMean"
    ]
    
    if all(col in df.columns for col in required_columns):
        error_label = Label(root, text="This file contains data appropriate columns for the CP Gait Classification", fg="green")
        error_label.place(x=75, y=570)
        ins_label = Label(root, text="Click on the CP Gait Classification Button:", fg="blue")
        ins_label.place(x=90, y=620)
        buttonGaitDT.config(state='normal')
    else:
        error_label = Label(root, text="This file does not contain the appropriate data", fg="red")
        error_label.place(x=75, y=570)
        buttonGaitDT.config(state='disabled')
        if ins_label:
            ins_label.destroy()

def save_Excel(knee_recommend, foot_recommend, sagittal_col, trans_col, selected_dates2):
    """
    Exports the classification results to an Excel file.
    
    Parameters:
        knee_recommend (list): Recommendations for knee.
        foot_recommend (list): Recommendations for foot.
        sagittal_col (list): Sagittal plane classifications.
        trans_col (list): Transversal plane classifications.
        selected_dates2 (list): The dates that were selected for filtering the data.
    """
    # Get the folder path entered by the user (if any)
    fold_name = folder_location.get().strip()

    # Determine the file path based on the folder provided or prompt the user.
    if fold_name:
        # Ensure the folder exists; if not, create it.
        if not os.path.exists(fold_name):
            try:
                os.makedirs(fold_name)
            except Exception as e:
                messagebox.showerror("Folder Creation Error",
                                     f"Failed to create folder '{fold_name}'.\n{e}")
                return  # Exit if the folder cannot be created
        # Use a generic file name.
        file_name = "DT_query.xlsx"
        file_path = os.path.join(fold_name, file_name)
    else:
        # Prompt the user to choose a save location if no folder is provided.
        file_path = filedialog.asksaveasfilename(
            title="Save Excel File",
            initialdir=file_directory,
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )

    # Verify that there are selected dates to filter the data.
    if not selected_dates2:
        messagebox.showwarning("Export Error", 
                               "There are no recommendations/classifications to export.\n"
                               "Try again by running the desired decision trees.")
        return

    # Load the Excel file from the determined file path.
    df = pd.read_excel(file_path)
    df = df[df['Date'].isin(selected_dates2)]  # Filter DataFrame by selected dates
    
    # Convert date columns to proper date format.
    df['Date'] = pd.to_datetime(df['Date']).dt.date
    df['DOB'] = pd.to_datetime(df['DOB']).dt.date

    # Check for length mismatches before adding recommendations.
    if (not knee_recommend) and (not foot_recommend) and (not sagittal_col) and (not trans_col):
        messagebox.showwarning("Export Error", 
                               "There are no recommendations/classifications to export.\n"
                               "Try again by running the desired decision trees.")
        return
    else:
        if knee_recommend and len(knee_recommend) != len(df):
            messagebox.showwarning("Export Error", 
                                   "Length mismatch for knee recommendations.\n"
                                   "Please click on 'New patient' to start again.")
            knee_recommend = None
        elif foot_recommend and len(foot_recommend) != len(df):
            messagebox.showwarning("Export Error", 
                                   "Length mismatch for foot recommendations.\n"
                                   "Please click on 'New patient' to start again.")
            foot_recommend = None
        elif sagittal_col and len(sagittal_col) != len(df):
            messagebox.showwarning("Export Error", 
                                   "Length mismatch for sagittal classifications.\n"
                                   "Please click on 'New patient' to start again.")
            sagittal_col = None
        elif trans_col and len(trans_col) != len(df):
            messagebox.showwarning("Export Error", 
                                   "Length mismatch for transversal classifications.\n"
                                   "Please click on 'New patient' to start again.")
            trans_col = None
        else:
            # Append missing values if necessary.
            if knee_recommend:
                if len(knee_recommend) < len(df):
                    knee_recommend += [np.nan] * (len(df) - len(knee_recommend))
                df.loc[:, 'Freeman Miller Knee Recommendations'] = knee_recommend

            if foot_recommend:
                if len(foot_recommend) < len(df):
                    foot_recommend += [np.nan] * (len(df) - len(foot_recommend))
                df.loc[:, 'Freeman Miller Foot Recommendations'] = foot_recommend

            if sagittal_col:
                if len(sagittal_col) < len(df):
                    sagittal_col += [np.nan] * (len(df) - len(sagittal_col))
                df.loc[:, 'Gait Deviation Sagittal Plane'] = sagittal_col

            if trans_col:
                if len(trans_col) < len(df):
                    trans_col += [np.nan] * (len(df) - len(trans_col))
                df.loc[:, 'Gait Deviation Transversal Plane'] = trans_col
            
            # Finally, ask the user where to save the exported file.
            initial_dir = file_directory
            export_file_path = filedialog.asksaveasfilename(
                initialdir=initial_dir, 
                defaultextension='.xlsx'
            )
            
            df.to_excel(export_file_path, index=False)
            messagebox.showinfo("Export", "File exported")

def new_patient():
    """
    Resets the GUI to allow the entry of a new patient's data.
    Clears previous recommendations and destroys existing widgets.
    """
    global error_label, ins_label, button_export1, button_export2, button_criteria, diagnosis, diagnosis_label, knee_recommend, foot_recommend, sagittal_col, trans_col
    recommendation_text.delete('1.0', END)  # Clear recommendation text area
    folder_location.set("")  # Reset folder location
    
    # Destroy previous widgets if they exist
    if error_label:
        error_label.destroy()
    if ins_label:
        ins_label.destroy()
    if diagnosis_label:
        diagnosis_label.destroy()
    if diagnosis:
        diagnosis.destroy()
    if button_export1:
        button_export1.destroy()
    if button_export2:
        button_export2.destroy()
    if button_criteria:
        button_criteria.destroy()
    
    # Clear recommendation lists
    knee_recommend.clear()
    foot_recommend.clear()
    sagittal_col.clear()
    trans_col.clear()

def GaitDev():
    """
    Processes the loaded Excel file to classify gait phenotypes.
    It:
      - Reads the data.
      - Allows the user to select encounter dates.
      - Classifies each encounter based on sagittal and transversal parameters.
      - Displays the results in a text widget.
      - Provides options to export results as Excel or PDF.
    """
    global error_label, ins_label, button_export1, button_export2, button_criteria, sagittal_col, trans_col, instructions_label, selected_file_path
    # Destroy previous labels/buttons if they exist
    if error_label:
        error_label.destroy()
    if ins_label:
        ins_label.destroy()
    if button_export1:
        button_export1.destroy()
    if button_export2:
        button_export2.destroy()
    if button_criteria:
        button_criteria.destroy()
    if instructions_label:
        instructions_label.destroy()
    
    recommendation_str = ""  # Initialize the recommendation string
    
    # Ensure a file has been selected
    if not selected_file_path:
        messagebox.showerror("Missing data", "No MRN typed or file selected.\nPlease try again.")
        return
    
    # Load data from the selected Excel file
    df = pd.read_excel(selected_file_path)
    
    # Determine if all data should be processed or if date selection is needed
    if 'complete_query' in selected_file_path:
        num_rows = len(df.index)
        filtered_rows = range(num_rows)
    else:
        # Ask user to select dates and filter DataFrame accordingly
        selected_dates2 = select_dates(df)
        df['Date'] = pd.to_datetime(df['Date']).dt.date
        filtered_rows = df[df['Date'].isin(selected_dates2)].index
    
    # Update title label with classification header
    title_label.config(text="CP Gait Classification", fg="black", font=('Open Sans - Bold', 16))
    
    # Process each row in the filtered data to classify gait
    for row_index in filtered_rows:
        # Extract required values from the DataFrame
        KneeFlexMin = df.loc[row_index, 'KneeFlexionMin'] 
        AnkleDorsi = df.loc[row_index, 'AnkleDorsiPlantarPeakDFStance'] 
        KneeFlexMean = df.loc[row_index, 'KneeFlexionMeanStance'] 
        KneeRange = df.loc[row_index, 'KneeFlexionTotalRange']
        HipFlexMin = df.loc[row_index, 'HipFlexionMin']
        FootProgress = df.loc[row_index, 'FootProgressionMeanStance']
        KneeRot = df.loc[row_index, 'KneeRotationMeanStance']
        HipRot = df.loc[row_index, 'HipRotationMean']
        PelvicRot = df.loc[row_index, 'PelvicRotationMean']
        Side = df.loc[row_index,'MotionParams_Side'] 
        EncounterAge = df.loc[row_index, 'EncounterAge']
        Trial = df.loc[row_index, 'GcdFile']  # Trial number from the 'GcdFile' column
        Date = df.loc[row_index, 'Date']
        MRN = int(df.loc[row_index, 'MRN'])
        e_date = str(Date).split()
        Date_E = str(e_date[0])  # Extract date (ignoring time if present)
        
        # Determine which side (left/right) based on data
        if 'R' in str(Side).upper():
            text_side = 'Right side'
        elif 'L' in str(Side).upper():
            text_side = 'Left side'
        else:
            text_side = None
            continue  # Skip row if side is not defined
        
        # ------------------ Sagittal Plane Classification ------------------
        if KneeFlexMin <= 0:
            if AnkleDorsi <= 4.4:
                if HipFlexMin <= 6.9:
                    sagittal = 'Knee Recurvatum, Ankle Plantar Flexed and Hip Not Flexed'
                else:
                    sagittal = 'Knee Recurvatum, Ankle Plantar Flexed and Hip Flexed'
            elif (AnkleDorsi <= 20.3) and (AnkleDorsi > 4.4):
                if HipFlexMin <= 6.9:
                    sagittal = 'Knee Recurvatum, Ankle Normal Flexed and Hip Not Flexed'
                else:
                    sagittal = 'Knee Recurvatum, Ankle Normal Flexed and Hip Flexed'
            elif AnkleDorsi > 20.3:
                if HipFlexMin <= 6.9:
                    sagittal = 'Knee Recurvatum, Ankle Dorsi Flexed and Hip Not Flexed'
                else:
                    sagittal = 'Knee Recurvatum, Ankle Dorsi Flexed and Hip Flexed'
            else:
                sagittal = 'No Classification'
        elif KneeFlexMin > 0:
            if AnkleDorsi > 20.3:
                if (KneeFlexMean <= 23.1) and (HipFlexMin > 6.9):
                    sagittal = 'Ankle with Excess Dorsiflexion and Hip Flexed'
                elif (KneeFlexMean > 23.1) and (HipFlexMin > 6.9):
                    sagittal = 'Ankle with Excess Dorsiflexion, Crouch Hip and Knee Flexed'
                elif (KneeFlexMean <= 23.1) and (HipFlexMin <= 6.9):
                    sagittal = 'Ankle with Excess Dorsiflexion and Knee Not Flexed'
                elif (KneeFlexMean > 23.1) and (HipFlexMin <= 6.9):
                    sagittal = 'Ankle with Excess Dorsiflexion and Knee Flexed'
                else:
                    sagittal = 'No Classification'
            elif AnkleDorsi <= 0:
                if (KneeFlexMean > 23.1) and (HipFlexMin <= 6.9):
                    sagittal = 'Ankle Equinus and Excess Knee Flexion'
                elif (KneeFlexMean <= 23.1) and (HipFlexMin > 6.9):
                    sagittal = 'Ankle Equinus and Excess Hip Flexion'
                elif (KneeFlexMean <= 23.1) and (HipFlexMin <= 6.9):
                    sagittal = 'Ankle Equinus, No Excess Flexion of Knee or Hip'
                elif (KneeFlexMean > 23.1) and (HipFlexMin > 6.9):
                    sagittal = 'Ankle Equinus, Excess Knee & Hip Flexion'
                else:
                    sagittal = 'No Classification'
            elif (AnkleDorsi > 0) and (AnkleDorsi < 20.3):
                if (AnkleDorsi > 0) and (AnkleDorsi <= 4.4):
                    if (KneeFlexMean <= 23.1) and (HipFlexMin > 6.9):
                        sagittal = 'Apparent Ankle Equinus and Hip Flexed'
                    elif (KneeFlexMean > 23.1) and (HipFlexMin > 6.9):
                        sagittal = 'Apparent Ankle Equinus & Hip and Knee Flexed'
                    elif (KneeFlexMean > 23.1) and (HipFlexMin <= 6.9):
                        sagittal = 'Apparent Ankle Equinus and Knee Flexed'
                    elif (KneeFlexMean <= 23.1) and (HipFlexMin <= 6.9):
                        sagittal = 'Apparent Ankle Equinus & No Excess Hip or Knee Flexion'
                    else:
                        sagittal = 'No Classification'
                elif (AnkleDorsi > 4.4) and (AnkleDorsi < 20.3):
                    if (KneeFlexMean <= 23.1 and KneeFlexMean > 4.4) and (HipFlexMin > 6.9):
                        sagittal = 'Normal Ankle, Exclusive Hip Flexion'
                    elif (KneeFlexMean > 23.1) and (HipFlexMin > 6.9):
                        sagittal = 'Normal Ankle, Knee & Hip Flexion'
                    elif (KneeFlexMean > 23.1) and (HipFlexMin <= 6.9):
                        sagittal = 'Normal Ankle, Exclusive Knee Flexion'
                    elif (KneeFlexMean <= 23.1 and KneeFlexMean > 4.4) and (HipFlexMin <= 6.9) and (KneeRange > 43.3):
                        sagittal = 'Normal Ankle, Normal Ranges for Knee and Hip'
                    else:
                        sagittal = 'No Classification'
                else:
                    if (KneeRange <= 43.3) and (KneeFlexMean <= 23.1) and (HipFlexMin <= 6.9):
                        sagittal = 'Exclusive Stiff Knee'
                    else:
                        sagittal = 'No Classification'
            else:
                sagittal = 'No Classification'
        else:
            sagittal = 'No Classification'
        
        # ---------------- Transversal Plane Classification ----------------
        if FootProgress > 6.3:
            if (KneeRot > 17.4) and (HipRot > 19) and (PelvicRot > 7.8):
                transversal = 'Foot Adductus Level 3, knee, hip and pelvis internal'
            elif (KneeRot <= 17.4) and (HipRot <= 19) and (PelvicRot <= 7.8):
                transversal = 'Foot Adductus Level 0, Knee, hip and pelvis not internal'
            elif ((KneeRot > 17.4) and (HipRot > 19)):
                transversal = 'Foot Adductus Level 2, with internal knee and hip'
            elif ((HipRot > 19) and (PelvicRot > 7.8)):
                transversal = 'Foot Adductus Level 2, with internal hip and pelvis'
            elif ((KneeRot > 17.4) and (PelvicRot > 7.8)):
                transversal = 'Foot Adductus Level 2, with internal knee and pelvis'
            else:
                if (KneeRot > 17.4):
                    transversal = 'Foot Adductus Level 1 with internal knee'
                elif (HipRot > 19):
                    transversal = 'Foot Adductus Level 1 with internal hip'
                elif (PelvicRot > 7.8):
                    transversal = 'Foot Adductus Level 1 with internal pelvis'
                else:
                    transversal = 'No Classification'
        elif FootProgress <= -21.5:
            if (KneeRot <= -15.8) and (HipRot <= -17.4) and (PelvicRot <= -7.3):
                transversal = 'Foot Abductus Level 3, knee, hip and pelvis external'
            elif (KneeRot > -15.8) and (HipRot > -17.4) and (PelvicRot > -7.3):
                transversal = 'Foot Abductus Level 0, Knee, hip and pelvis not external'
            elif ((KneeRot <= -15.8) and (HipRot <= -17.4)):
                transversal = 'Foot Abductus Level 2, with external knee and hip'
            elif ((HipRot <= -17.4) and (PelvicRot <= -7.3)):
                transversal = 'Foot Abductus Level 2, with external hip and pelvis'
            elif ((KneeRot <= -15.8) and (PelvicRot <= -7.3)):
                transversal = 'Foot Abductus Level 2, with external knee and pelvis'
            else:
                if (KneeRot <= -15.8):
                    transversal = 'Foot Abductus Level 1, with external knee'
                elif (HipRot <= -17.4):
                    transversal = 'Foot Abductus Level 1, with external hip'
                elif (PelvicRot <= -7.3):
                    transversal = 'Foot Abductus Level 1, with external pelvis'
                else:
                    transversal = 'No Classification'
        elif (FootProgress > -21.5) and (FootProgress <= 6.3):
            if ((KneeRot <= 17.4) and (KneeRot > -15.8)) and ((HipRot <= 19) and (HipRot > -17.4)) and ((PelvicRot <= 7.8) and (PelvicRot > -7.3)):
                transversal = 'True Neutral'
            else:
                if KneeRot > 17.4:
                    if HipRot > 19:
                        if PelvicRot > 7.8:
                            trans1 = ' (internal knee, hip and pelvic rotation)'
                        elif PelvicRot <= -7.3:
                            trans1 = ' (internal knee and hip rotation, and external pelvic rotation)'
                        else:
                            trans1 = ' (internal knee and hip rotation)'
                    elif HipRot <= -17.4:
                        trans1 = ' (internal knee and external hip rotation)'
                    elif PelvicRot > 7.8:
                        trans1 = ' (internal knee and pelvic rotation)'
                    elif PelvicRot <= -7.3:
                        trans1 = ' (internal knee and external pelvic rotation)'
                    else:
                        trans1 = ' (internal knee rotation)'
                elif KneeRot <= -15.8:
                    if HipRot > 19:
                        if PelvicRot > 7.8:
                            trans1 = ' (external knee, and internal hip and pelvic rotation)'
                        elif PelvicRot <= -7.3:
                            trans1 = ' (external knee and pelvic rotation and internal hip rotation)'
                        else:
                            trans1 = ' (external knee and internal hip rotation)'
                    elif HipRot <= -17.4:
                        trans1 = ' (external knee and hip rotation)'
                    elif PelvicRot > 7.8:
                        trans1 = ' (external knee and internal pelvic rotation)'
                    elif PelvicRot <= -7.3:
                        trans1 = ' (external knee and external pelvic rotation)'
                    else:
                        trans1 = ' (external knee rotation)'
                elif HipRot > 19:
                    if PelvicRot > 7.8:
                        trans1 = ' (internal hip and pelvic rotation)'
                    elif PelvicRot <= -7.3:
                        trans1 = ' (internal hip and external pelvic rotation)'
                    else:
                        trans1 = ' (internal hip rotation)'
                elif HipRot <= -17.4:
                    if PelvicRot > 7.8:
                        trans1 = ' (external hip and internal pelvic rotation)'
                    elif PelvicRot <= -7.3:
                        trans1 = ' (external hip and pelvic rotation)'
                    else:
                        trans1 = ' (external hip rotation)'
                elif PelvicRot > 7.8:
                    trans1 = ' (internal pelvic rotation)'
                elif PelvicRot <= -7.3:
                    trans1 = ' (external pelvic rotation)'
                else:
                    trans1 = ''
                
                sum_C = KneeRot + HipRot + PelvicRot
                if (sum_C > -10) and (sum_C < 10):
                    transversal = 'Compensated Neutral' + trans1
                elif (sum_C > 10) or (sum_C < -10):
                    transversal = 'Neutral Other' + trans1
                else:
                    transversal = 'No Classification'
        else:
            transversal = 'No Classification'
        
        # Append the classification result to the recommendation string
        recommendation_str += (
            "Patient " + str(MRN) + ", " + text_side +
            ", Encounter Date: " + str(Date_E) +
            ", Trial: " + str(Trial) +
            "\n  Sagittal: " + str(sagittal) +
            "\n  Transverse: " + str(transversal) + "\n\n"
        )
        sagittal_col.append(sagittal)
        trans_col.append(transversal)
    
    # Display the classification results in the text widget
    lines = recommendation_str.split("\n\n")
    recommendation_text.delete('1.0', END)
    for line in lines:
        recommendation_text.insert(END, line + "\n\n", ('tag1', 'color_black'))
        recommendation_text.tag_configure('tag1', font=('Open Sans', 12))
        recommendation_text.tag_configure('color_black', foreground='black')
    
    def export_Excel(sagittal_col, trans_col, df):
        """
        Exports the current classification results to an Excel file.
        """
        df['Gait Deviation Sagittal Plane'] = sagittal_col
        df['Gait Deviation Transverse Plane'] = trans_col
        export_file_path = filedialog.asksaveasfilename(defaultextension='.xlsx')
        df.to_excel(export_file_path, index=False)
    
    def export_pdf(recommendation_str, MRN, selected_dates2):
        """
        Exports the classification report to a PDF file.
        """
        messagebox.showinfo("Save report", "Select the folder to export the report\n\nThe report is in PDF format")
        folder_pathpdf = filedialog.askdirectory()
        if not folder_pathpdf:
            return
        # Create a string of years from the selected dates for the filename
        dates = [str(date.year) for date in selected_dates2]
        joined_dates = '_'.join(dates)
        filename = f"{folder_pathpdf}/{MRN}_GaitClassif_{joined_dates}_report.pdf"
        
        c = canvas.Canvas(filename, pagesize=letter)
        max_width = 115  # Controls the horizontal text wrap width
        page_height = c._pagesize[1] - 5
        recommendation_lines = recommendation_str.split("\n\n")
        
        y = page_height - 30
        add_extra_text = True  # Header flag
        
        # Write the report text into the PDF
        for i, rec in enumerate(recommendation_lines):
            if add_extra_text and i == 0:
                y -= 5
                c.setFont("Helvetica-Bold", 16)
                c.drawString(230, y, "Gait Classifications")
                y -= 40
                add_extra_text = False
            
            c.setFont("Helvetica", 10)
            c.setFillColorRGB(0, 0, 0)  # Black text
            wrapped_recommendation = textwrap.wrap(rec, width=max_width)
            for line in wrapped_recommendation:
                c.drawString(50, y, line)
                y -= 12
            
            # Add a new page if there is not enough space left
            if y < 40:
                c.showPage()
                y = page_height - 30
            y -= 15
        c.save()
        messagebox.showinfo("Export", "Report exported")
    
    # Create export buttons on the main window
    button_export2 = Button(root, text="Export\nExcel\nFile", font=("Open Sans", 12, 'bold'), bg="#BFBFBF", width=13, height=4,
                            command=lambda: export_Excel(sagittal_col, trans_col, df))
    button_export2.place(x=30, y=400)
    
    button_export1 = Button(root, text="Export\nReport", font=("Open Sans", 12, 'bold'), bg="#BFBFBF", width=13, height=4,
                            command=lambda: export_pdf(recommendation_str, MRN, selected_dates2))
    button_export1.place(x=185, y=400)

# ------------------------ Main Application Execution ------------------------

# Handle resource paths when bundled as an executable
if getattr(sys, 'frozen', False):
    script_dir = sys._MEIPASS
else:
    script_dir = os.path.dirname(os.path.abspath(__file__))

# Display version information
Label_Version = Label(root, text="Version 1.1\n2/2025", fg="#940434")
Label_Version.config(font=("Open Sans", 8))
Label_Version.place(x=1300, y=10)

# Main title label for the application
Label2 = Label(root, text="Cerebral Palsy Management\nAutomated Gait Phenotype Classification ", fg="#940434")
Label2.config(font=("Open Sans", 22))
Label2.place(x=375, y=50)

# Button to select the Excel file
ButtonB = Button(root, width=25, text="Select Excel file", font=("Open Sans", 12, 'bold'),
                 padx=20, pady=5, bg="#A7C4C2", command=browse_file)
ButtonB.place(x=30, y=230)

# Initialize global label variables for later use
error_label = None
ins_label = None
button_export1 = None
button_export2 = None
button_criteria = None
diagnosis = None
diagnosis_label = None
selected_dates2 = None
sagittal_col = []
trans_col = []
instructions_label = None 

# Title label to display dynamic information during classification
title_label = Label(root)
title_label.place(x=520, y=285)

# Button to trigger the CP Gait Classification process
buttonGaitDT = Button(root, width=25, text="CP Gait Classification", font=("Open Sans", 12, 'bold'),
                      padx=20, pady=15, bg="#FFF2CC", command=GaitDev)
buttonGaitDT.place(x=30, y=290)

# ------------------- Setup Text Widget for Displaying Results -------------------
frame = Frame(root)
frame.place(x=490, y=230)

# Vertical scrollbar for the text widget
vscrollbar = Scrollbar(frame)
vscrollbar.pack(side=RIGHT, fill=Y)

# Horizontal scrollbar for the text widget
hscrollbar = Scrollbar(frame, orient=HORIZONTAL)
hscrollbar.pack(side=BOTTOM, fill=X)

# Text widget to display recommendations and classifications
recommendation_text = Text(frame, height=25, width=105, wrap=NONE,
                            yscrollcommand=vscrollbar.set, xscrollcommand=hscrollbar.set, spacing3=5)
recommendation_text.pack(side=LEFT, fill=Y)

# Configure scrollbars to control the text widget
vscrollbar.config(command=recommendation_text.yview)
hscrollbar.config(command=recommendation_text.xview)

# Instructions in the large text box
recommendation_str = (
    "1. Click on 'Select Excel file' to load your Excel file containing the required data.\n"
    "2. Ensure the file includes all necessary columns (e.g., Date, MRN, MotionParams_Side, etc.).\n"
    "3. Click the 'CP Gait Classification' button to run the analysis.\n"
    "4. Review the classification results displayed in this area.\n"
    "5. Use the export buttons to save the results as an Excel file or PDF report.\n"
)
recommendation_text.insert(END, recommendation_str, ('tag1', 'color_blue', 'space1'))
recommendation_text.tag_configure('tag1', font=('Open Sans', 12))
recommendation_text.tag_configure('color_blue', foreground='black')
recommendation_text.tag_configure('space1', spacing3=15)

# ---------------------------- Start the GUI Loop ----------------------------
root.mainloop()  # Start the event loop for the application

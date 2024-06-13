# Importing necessary libraries
import os 
import sys
import pandas as pd 
import tkinter as tk
import math
import textwrap
import pyodbc
import re
import numpy as np
import struct
import PIL

from tkinter import messagebox, filedialog, StringVar, END, Scrollbar, HORIZONTAL
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from tkinter import *
from PIL import Image, ImageTk
print(PIL.__version__)

root = Tk()
root.title("Cerebral Palsy Gait Phenotype Classification")
root.geometry("1400x900+280+60")
root.resizable(False, False)

'''##################################################################################'''
# Definition of Variables
folder_location = StringVar()  # Create a StringVar object to hold the path string
selected_file_path = None  # Global variable to store the selected file path
'''##################################################################################'''

# Definition of Functions #########################################################

def select_dates(df):
    '''# Select the Encounter dates that you would like to show the recommendations for'''
    popup1 = tk.Toplevel()
    popup1.title("Select Dates")

    unique_dates = list(df['Date'].apply(lambda x: x.date()).unique())
    count_dates = len(unique_dates)

    popup1.geometry("360x" + str(((count_dates+1)*20)+150) + "+850+450")
    # Create a list to store the selected dates
    global selected_dates 
    selected_dates = []

    # Add a title to the window
    title_label = tk.Label(popup1, text="Select the Encounter Date(s) \nfor which the recommendations are required")
    title_label.grid(row=0, column=0, sticky=NSEW)

    # Create a Checkbutton for the "All" option
    var_all = BooleanVar()
    cb_all = Checkbutton(popup1, text="All", variable=var_all, onvalue=True, offvalue=False)
    cb_all.grid(row=1, column=0, sticky=W)

    # Create a Checkbutton for each unique date
    for i, date in enumerate(unique_dates):
        var = BooleanVar()
        cb = Checkbutton(popup1, text=date, variable=var, onvalue=True, offvalue=False)
        cb.grid(row=i+2, column=0, sticky=W)

        # Bind a function to the Checkbutton to add/remove the date from the selected_dates list
        def toggle_date(date=date, var=var):
            if var.get():
                if date not in selected_dates:
                    selected_dates.append(date)
            else:
                if date in selected_dates:
                    selected_dates.remove(date)

        cb.config(command=toggle_date)

    def save_dates():
        '''# Function to save selected dates and close the pop-up window'''
        global selected_dates
        if var_all.get():
            # If "All" is selected, use all dates
            selected_dates = unique_dates
        popup1.destroy()

    # Create a "Save" button to save selected dates and close the pop-up window
    button_save = Button(popup1, text="Continue", padx=20, pady=10, command=save_dates)
    button_save.grid(row=2, column=1)

    # Display the pop-up window and wait for it to be destroyed
    popup1.grab_set()
    popup1.wait_window()

    # Return the user's choice
    return selected_dates

def browse_file():
    global error_label, ins_label, file_directory, selected_file_path
    selected_file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if selected_file_path == '':
        messagebox.showerror("Error", "Please choose a file or type an MRN number")
        return
    file_directory = os.path.dirname(selected_file_path)  # Capture the directory of the chosen file
    if error_label:
        error_label.destroy()
    if ins_label:
        ins_label.destroy()

    df = pd.read_excel(selected_file_path)
       
    # Existing checks and labels
    if all(col in df.columns for col in ["MotionParams_Side", "EncounterAge", "Date", "MRN", \
        "AnkleDorsiPlantarPeakDFStance", "KneeFlexionMeanStance", "KneeFlexionTotalRange", "HipFlexionMin", "KneeFlexionMin", \
            "FootProgressionMeanStance", "KneeRotationMeanStance", "HipRotationMean", "PelvicRotationMean"]):
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
    fold_name = folder_location.get()
    if fold_name:
        if fold_name.isdigit():
            file_name = str(fold_name) + "_DT_query.xlsx"
            folder_path = os.path.join(folder_location, str(fold_name))
            file_path = os.path.join(folder_path, file_name)
        else:
            file_path= fold_name
            file_path=f"\\\\CHI-FS-APP03\\Motion Analysis Center\\Gait Lab\\_Patients\\{mrn_input}\\{mrn_input}_DT_query.xlsx"
            #file_path=f"G:\\Gait Lab\\_Patients\\{mrn_input}\\{mrn_input}_DT_query.xlsx"

        if not selected_dates2:
            messagebox.showwarning("Export Error", "There are no recommendations/classifications to export.\nTry again by running the desired decision trees.")
        else:
            df = pd.read_excel(file_path)  # Read the Excel file
            df = df[df['Date'].isin(selected_dates2)]  # Subset the DataFrame based on selected dates

            df['Date'] = pd.to_datetime(df['Date']).dt.date
            df['DOB'] = pd.to_datetime(df['DOB']).dt.date

            if (not knee_recommend) and (not foot_recommend) and (not sagittal_col) and (not trans_col):
                messagebox.showwarning("Export Error", "There are no recommendations/classifications to export. Try again by running the desired decision trees.")
            else:
                if knee_recommend and len(knee_recommend) != len(df):
                    messagebox.showwarning("Export Error", "Length mismatch: The dates selected for the knee recommendations do not match the ones selected for the other"+\
                        " decision tree or the gait deviation classifications.\nPlease click on 'New patient' to start again.")
                    knee_recommend = None
                elif foot_recommend and len(foot_recommend) != len(df):
                    messagebox.showwarning("Export Error", "Length mismatch: The dates selected for the foot recommendations do not match the ones selected for the other"+\
                        " decision tree or the gait deviation classifications.\nPlease click on 'New patient' to start again.")
                    foot_recommend = None
                elif sagittal_col and len(sagittal_col) != len(df):
                    messagebox.showwarning("Export Error", "Length mismatch: The dates selected for the gait deviations do not match the ones selected for the other"+\
                        " decision trees.\nPlease click on 'New patient' to start again.")
                    sagittal_col = None
                elif trans_col and len(trans_col) != len(df):
                    messagebox.showwarning("Export Error", "Length mismatch: The dates selected for the gait deviations do not match the ones selected for the other"+\
                        " decision trees.\nPlease click on 'New patient' to start again.")
                    trans_col = None
                else:
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
                
                    initial_dir = file_directory  # Use the directory where the file was chosen from
                    export_file_path = filedialog.asksaveasfilename(initialdir=initial_dir, defaultextension='.xlsx')

                    df.to_excel(export_file_path, index=False)
                    messagebox.showinfo("Export", "File exported")
    else:
        messagebox.showwarning("Export Error", "There is no data to export. \n First select the patient's data and try again.")

def new_patient():
    '''# Function to delete the previous data and look for a new patient'''
    global error_label, ins_label, button_export1, button_export2, button_criteria, diagnosis, diagnosis_label, knee_recommend, foot_recommend, sagittal_col, trans_col
    recommendation_text.delete('1.0', END)
    folder_location.set("")

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

    knee_recommend.clear()
    foot_recommend.clear()
    sagittal_col.clear()
    trans_col.clear()

'''Gait Classification - Phenotypes'''
def GaitDev():
    global error_label, ins_label, button_export1, button_export2, button_criteria, sagittal_col, trans_col, instructions_label, selected_file_path
    # Destroy previous labels if they exist
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

    # Start of the function
    recommendation_str = ""

    if not selected_file_path:
        messagebox.showerror("Missing data", "There was not MRN typed or any browsed file\nPlease try again")
        return

    df = pd.read_excel(selected_file_path)  # Read the excel file

    if 'complete_query' in selected_file_path:
        num_rows = len(df.index)
        filtered_rows = range(num_rows)
    else:
        # Select dates
        selected_dates2 = select_dates(df)
        df['Date'] = pd.to_datetime(df['Date']).dt.date
        filtered_rows = df[df['Date'].isin(selected_dates2)].index

    title_label.config(text="CP Gait Classification", fg="black", font=('Open Sans - Bold', 16))

    ######
    # Establish cutoff values (they can be an input from the user)
    for row_index in filtered_rows:
        # Global variables for this part
        KneeFlexMin = df.loc[row_index, 'KneeFlexionMin'] 
        AnkleDorsi = df.loc[row_index, 'AnkleDorsiPlantarPeakDFStance'] 
        KneeFlexMean = df.loc[row_index, 'KneeFlexionMeanStance'] 
        KneeRange = df.loc[row_index, 'KneeFlexionTotalRange']
        HipFlexMin = df.loc[row_index, 'HipFlexionMin']
        FootProgress = df.loc[row_index, 'FootProgressionMeanStance']
        KneeRot= df.loc[row_index, 'KneeRotationMeanStance']
        HipRot = df.loc[row_index, 'HipRotationMean']
        PelvicRot = df.loc[row_index, 'PelvicRotationMean']
        Side = df.loc[row_index,'MotionParams_Side'] 
        EncounterAge = df.loc[row_index, 'EncounterAge']
        Trial = df.loc[row_index, 'GcdFile']  # Trial number that is in the GcdFile column
        Date = df.loc[row_index, 'Date']
        MRN = int(df.loc[row_index, 'MRN'])
        e_date = str(Date).split()
        if len(e_date) > 1:
            t_date = None
            Date_E = str(e_date[0]) 
        else:
            Date_E = str(e_date[0]) 

        # Variables for right side
        if 'R' in str(Side).upper():
            text_side = 'Right side'

        # Variables for left side
        elif 'L' in str(Side).upper():
            text_side = 'Left side'
        
        # If there's no L or R it won't go to the decision tree
        else: 
            text_side = None
            continue

        # Sagittal PLANE 
        if KneeFlexMin <= 0:
            if AnkleDorsi <= 4.4:
                if HipFlexMin <= 6.9:
                    sagittal = 'Knee Recurvatum, Ankle Plantar Flexed and Hip Not Flexed'
                else:  # HipFlexMin > 6.9:
                    sagittal = 'Knee Recurvatum, Ankle Plantar Flexed and Hip Flexed'
            elif (AnkleDorsi <= 20.3) and (AnkleDorsi > 4.4):
                if HipFlexMin <= 6.9:
                    sagittal = 'Knee Recurvatum, Ankle Normal Flexed and Hip Not Flexed'
                else:  # HipFlexMin > 6.9:
                    sagittal = 'Knee Recurvatum, Ankle Normal Flexed and Hip Flexed'
            elif AnkleDorsi > 20.3:
                if HipFlexMin <= 6.9:
                    sagittal = 'Knee Recurvatum, Ankle Dorsi Flexed and Hip Not Flexed'
                else:  # HipFlexMin > 6.9:
                    sagittal = 'Knee Recurvatum, Ankle Dorsi Flexed and Hip Flexed'
            else:
                sagittal = 'No Classification'  # No classification
        
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
                        sagittal = 'No Classification'  # No classification
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
                        sagittal = 'No Classification'  # No classification
                else:
                    if (KneeRange <= 43.3) and (KneeFlexMean <= 23.1) and (HipFlexMin <= 6.9):
                        sagittal = 'Exclusive Stiff Knee'
                    else:
                        sagittal = 'No Classification'
            else:
                sagittal = 'No Classification'
        else:
            sagittal = 'No Classification'

        # TRANSVERSAL PLANE
        # Internal
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
            
        # External
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

        # Neutral 
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
        
        recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E)  + ", Trial: " + str(Trial) + "\n  Sagittal: " + str(sagittal) + \
            "\n  Transverse: " + str(transversal) + "\n\n"
        sagittal_col.append(sagittal)
        trans_col.append(transversal)

    lines = recommendation_str.split("\n\n")
    recommendation_text.delete('1.0', END)
    # Use a loop to add 1 recommendation and 1 explanation at a time
    for i in range(len(lines)):
        recommendation_text.insert(END, lines[i] + "\n\n", ('tag1', 'color_black'))
        recommendation_text.tag_configure('tag1', font=('Open Sans', 12))
        recommendation_text.tag_configure('color_black', foreground='black')
    
    def export_Excel(sagittal_col, trans_col, df):
        df['Gait Deviation Sagittal Plane'] = sagittal_col
        df['Gait Deviation Transverse Plane'] = trans_col
        export_file_path = filedialog.asksaveasfilename(defaultextension='.xlsx')
        df.to_excel(export_file_path, index=False)

    def export_pdf(recommendation_str, MRN, selected_dates2):
        messagebox.showinfo("Save report", "Select the folder to export the report\n\nThe report is in pdf format")
        folder_pathpdf = filedialog.askdirectory()
        if not folder_pathpdf:
            return
        dates = [str(date.year) for date in selected_dates2]

        # Join dates with underscore
        joined_dates = '_'.join(dates)

        filename=f"{folder_pathpdf}/" + str(MRN) + "_GaitClassif_"+ joined_dates +"_report.pdf"

        c = canvas.Canvas(filename, pagesize=letter)
        max_width = 115  # This controls the horizontal size in which the text can be displayed
        page_height = c._pagesize[1] - 5
        recommendation_lines = recommendation_str.split("\n\n")

        y = page_height - 30
        add_extra_text = True  # Verify that there is enough space left on the same page for the full recommendation or explanation
        for i in range(len(recommendation_lines)):
            if add_extra_text and i == 0:
                y -= 5  # Space to add the header
                c.setFont("Helvetica-Bold", 16)
                c.drawString(230, y, "Gait Classifications")
                # c.drawString(180, y-20, "Knee Treatment Decision Tree")
                y -= 40
                add_extra_text = False

            c.setFont("Helvetica", 10)
            c.setFillColorRGB(0, 0, 0)  # Black color
            wrapped_recommendation = textwrap.wrap(recommendation_lines[i], width=max_width)
            for line in wrapped_recommendation:
                c.drawString(50, y, line)
                y -= 12
        
            # Check the space left on the page to avoid cropped information
            if y < 40:
                c.showPage()
                y = page_height - 30
            y -= 15
        c.save()
        messagebox.showinfo("Export", "Report exported")
    
    # Button to export Excel file
    button_export2 = Button(root, text="Export\nExcel\nFile", font=("Open Sans", 12, 'bold'), bg="#BFBFBF", width=13, height=4, command=lambda: export_Excel(sagittal_col, trans_col, df))
    button_export2.place(x=30, y=400)

    # Button to export the recommendations in a pdf file
    button_export1 = Button(root, text="Export\nReport", font=("Open Sans", 12, 'bold'), bg="#BFBFBF", width=13, height=4, command=lambda: export_pdf(recommendation_str, MRN, selected_dates2))
    button_export1.place(x=185, y=400)

#####################################################################################

# To access the resources (pdfs and images) when having the .exe
if getattr(sys, 'frozen', False):
    script_dir = sys._MEIPASS
else:
    script_dir = os.path.dirname(os.path.abspath(__file__))

# Version Number
Label_Version = Label(root, text="Version 1.0\n7/2024", fg="#940434")
Label_Version.config(font=("Open Sans", 8))
Label_Version.place(x=1300, y=10)

# Simple title label for the decision Tree
Label2 = Label(root, text="Cerebral Palsy Management\nAutomated Gait Phenotype Classification ", fg="#940434")
Label2.config(font=("Open Sans", 22))
Label2.place(x=375, y=50)

# Select button
ButtonB = Button(root, width=25, text="Select Excel file", font=("Open Sans", 12, 'bold'), padx=20, pady=5, bg="#A7C4C2", command=browse_file)
ButtonB.place(x=30, y=230)

# When pressing 'enter' the search button triggers automatically
# folder_name.bind("<Return>", on_enter)

# Define the labels but start them in zero to use them in multiple functions
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
title_label = None
title_label = Label(root)  
title_label.place(x=520, y=285)

# Button for the Gait Deviation Classification
buttonGaitDT = Button(root, width=25, text="CP Gait Classification", font=("Open Sans", 12, 'bold'), padx=20, pady=15, bg="#FFF2CC", command=GaitDev)
buttonGaitDT.place(x=30, y=290)

# Create a Frame to hold the Text widgets and scrollbar
frame = Frame(root)
frame.place(x=490, y=230)

# Create a scrollbar for the Text widgets
vscrollbar = Scrollbar(frame)
vscrollbar.pack(side=RIGHT, fill=Y)

# Create a horizontal scrollbar for the Text widget
hscrollbar = Scrollbar(frame, orient=HORIZONTAL)
hscrollbar.pack(side=BOTTOM, fill=X)

# Create a Text widget to hold the recommendation text
recommendation_text = Text(frame, height=25, width=105, wrap=NONE, yscrollcommand=vscrollbar.set, xscrollcommand=hscrollbar.set, spacing3=5)
recommendation_text.pack(side=LEFT, fill=Y)

# Set the scrollbar to control the Text widget
vscrollbar.config(command=recommendation_text.yview)

# Set the horizontal scrollbar to control the Text widget
hscrollbar.config(command=recommendation_text.xview)

# Instructions in the large text box
recommendation_str = ""
recommendation_str += " 1. Add Instructions Here...\n"
recommendation_text.insert(END, recommendation_str, ('tag1', 'color_blue', 'space1'))
recommendation_text.tag_configure('tag1', font=('Open Sans', 12))
recommendation_text.tag_configure('color_blue', foreground='black')
recommendation_text.tag_configure('space1', spacing3=15)

'''''last line follows this'''
root.mainloop()  # Start the event loop

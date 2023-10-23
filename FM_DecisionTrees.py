'''
Program by Ana Flores 
Freeman Miller Knee and Foot Decision Trees
Gait Deviation Classification
'''

from tkinter import *
import os 
import sys
import pandas as pd 
import tkinter as tk
from PIL import Image, ImageTk
from tkinter import messagebox, filedialog
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import math
import textwrap
import pyodbc
import re
import numpy as np

root = Tk()
root.title("Cerebral Palsy Management Decision Trees")
root.geometry("1400x900+280+60")

root.resizable(False, False)

# to access the resources (pdfs and images) when having the .exe
if getattr(sys, 'frozen', False):
    script_dir = sys._MEIPASS
else:
    script_dir = os.path.dirname(os.path.abspath(__file__))

# First instruction to enter MRN number
Label1 = Label(root, text="Enter the patient MRN: ", font=("Open Sans", 12))
Label1.place(x=75, y=170)

message1 = Label(root, text="Or type the patient's Last AND First name")
message1.place(x=85, y=220)

Label2 = Label(root, text="Last Name: ", font=("Open Sans", 12))
Label2.place(x=75, y=250)

Label3 = Label(root, text="First Name: ", font=("Open Sans", 12))
Label3.place(x=75, y=290)

Label4 = Label(root, text="Or select the Excel file: ", font=("Open Sans", 12))
Label4.place(x=100, y=370) #280

# Shriners logo
logo_path = os.path.join(script_dir, 'resources', 'logo.png')
logo = Image.open(logo_path)
logo1=logo.resize((321,140), Image.LANCZOS) # the original size of the image is too big
logo2 = ImageTk.PhotoImage(logo1)
logo_label = tk.Label(image=logo2) # logo2 is the image resized, this one goes to the label
logo_label.place(x=30, y=5)

# simple title label for the decision Tree
Label2 = Label(root, text="Cerebral Palsy Management\nFreeman Miller Decision Tree Recommendations", fg="#940434")
Label2.config(font=("Open Sans", 22))
Label2.place(x=580, y=50)

# Receive the MRN number that is also the folder name
folder_name = Entry(root, width=15, borderwidth=5, background="#DDEBF7", font=("Open Sans", 12))
folder_name.place(x=250, y=170)

first_name = Entry(root, width=19, borderwidth=5, background="#DDEBF7", font=("Open Sans", 12))
first_name.place(x=180, y=250)
last_name = Entry(root, width=19, borderwidth=5, background="#DDEBF7", font=("Open Sans", 12))
last_name.place(x=180, y=290)

# define the labels but start them in zero to use them in multiple functions
error_label = None
ins_label = None
button_export1 = None
button_export2 = None
button_criteria = None
diagnosis = None
diagnosis_label = None
selected_dates2 = None

knee_recommend = []
foot_recommend = []
sagittal_col = []
trans_col = []

# Label for the title text
title_label = Label(root)
title_label.place(x=520, y=245)

instructions_label = Label(root, text="Instructions", fg="black", font=('Open Sans - Bold', 16))
instructions_label.place(x=850, y=245)

#folder_location = "C:\\Users\\AGraf\\Desktop\\python_app"
folder_location="\\\\CHI-FS-APP03\\Motion Analysis Center\\Gait Lab\\_Patients"
#folder_location = "G:\\Gait Lab\\_Patients"

# Create a Frame to hold the Text widgets and scrollbar
frame = Frame(root)
frame.place(x=490, y=280)

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

recommendation_str = ""
recommendation_str += "1. Type a Patient's Medical Record Number or both First and Last Name.\n"
recommendation_str += "2. Click on 'Search' to create the patient's query.\n"
recommendation_str += "3. If there is an existing query created, the button 'Browse' can also be clicked to select that file.\n"
recommendation_str += "4. Click on the corresponding buttons for the Knee or Foot Decision Tree, or on the Gait Deviation Classification.\n"
recommendation_str += "5. Select the dates of the encounter and the recommendations will be shown.\n"
recommendation_str += "6. For the Decision Trees, the button 'Recommendation Criteria' shows the details for each recommendation made.\n"
recommendation_str += "7. There is the option to export reports, all of them in pdf format\n"
recommendation_str += "8. To export in an Excel Worksheet there are two options:\n" 
recommendation_str += "     a. After clicking one of the 3 possible options (Knee, Foot or Gait), next to the 'Recommendation Criteria' there is the\n" 
recommendation_str += "        'Export to excel' button. This will export to excel the recommendations as a new column, ONLY for that decision tree. \n" 
recommendation_str += "     b. The green 'Export Excel File' button at the left bottom exports all the recommendations. This has to be clicked after\n" 
recommendation_str += "        clicking the decision trees or the gait button. To export the data correctly, make sure the selected date is the same\n"
recommendation_str += "        in all the cases. \n"
recommendation_str += " * The 'run complete query' button is for a project (it contains data of patients with CP with encounter dates after 2005)\n"

recommendation_text.insert(END, recommendation_str, ('tag1', 'color_blue', 'space1'))
recommendation_text.tag_configure('tag1', font=('Open Sans', 12))
recommendation_text.tag_configure('color_blue', foreground='black')
recommendation_text.tag_configure('space1', spacing3=15)

# function to activate the search button by clicking enter 
def on_enter(event):
    buttonS.invoke()
    buttonS.focus_set()

# when pressiong 'enter' the search button triggers automatically
folder_name.bind("<Return>", on_enter)

#select the dates to show the recommendations
def select_dates(df):
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

    # Function to save selected dates and close the pop-up window
    def save_dates():
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

# function to find the excel file
def findExcel():
    global error_label, ins_label, diagnosis, diagnosis_label
    
    # Destroy previous labels if they exist
    if error_label:
        error_label.destroy()
    if ins_label:
        ins_label.destroy()
    if diagnosis:
        diagnosis.destroy()
    if diagnosis_label:
        diagnosis_label.destroy()

    conn = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb)};DBQ=\\CHI-FS-APP03\Motion Analysis Center\Gait Lab\_MAL Database\PatientsAndServicesDatabase.mdb;')
    cursor = conn.cursor()

    mrn_input = folder_name.get()
    last = last_name.get()
    first = first_name.get()
    
    if mrn_input:
        sql_query="""
        SELECT Patients.MRN, Patients.LastName, Patients.FirstName, Patients.PrimaryDiagnosis, 
            Patients.PatternOfInvolvement, Patients.Side AS Patients_Side, Patients.DOB, Encounters.Date, 
            DateDiff('yyyy',[Patients].[DOB],Date()) AS AgeNow, 
            DateDiff('yyyy',[Patients].[DOB],[Encounters].[Date]) AS EncounterAge, 
            MotionParams.GcdFile, PhysicalExams.GmfcsNew, 
            PhysicalExams.LeftForefootWtBearing, PhysicalExams.RightForefootWtBearing, PhysicalExams.LeftMidfootWtBearing, 
            PhysicalExams.RightMidfootWtBearing, PhysicalExams.LeftHindFootWtBearing, PhysicalExams.RightHindFootWtBearing, 
            PhysicalExamsMore.LeftHalluxWtBearing, PhysicalExamsMore.RightHalluxWtBearing, FootMilwaukeeAnglesEtc.LeftPhiHind, 
            FootMilwaukeeAnglesEtc.RightPhiHind, PhysicalExams.LeftFootSupple, PhysicalExams.RightFootSupple, 
            PhysicalExams.LeftOrthoticTolerated, PhysicalExams.RightOrthoticTolerated, PhysicalExams.LeftToneGastroc, 
            PhysicalExams.RightToneGastroc, PhysicalExams.LeftTonePeroneals, PhysicalExams.RightTonePeroneals, 
            PhysicalExams.LeftAnkleRomInver, PhysicalExams.RightAnkleRomInver, PhysicalExams.LeftAnkleRomEver, 
            PhysicalExams.RightAnkleRomEver, PhysicalExams.LeftThighFootAngle, PhysicalExams.RightThighFootAngle, 
            MotionParams.Side AS MotionParams_Side, MotionParams.AnkleDorsiPlantarMeanStanceDF, MotionParams.KneeFlexionAtIC, 
            PhysicalExams.LeftForefootOffWtBear, PhysicalExams.RightForefootOffWtBear, PhysicalExams.LeftMidfootOffWtBear, 
            PhysicalExams.RightMidfootOffWtBear, PhysicalExams.LeftHindFootOffWtBear, PhysicalExams.RightHindFootOffWtBear, 
            PhysicalExams.LeftFootOther, PhysicalExams.RightFootOther, PhysicalExams.LeftTonePeronealsClonus, 
            PhysicalExams.RightTonePeronealsClonus, PhysicalExams.LeftKneeRomPopAngle, PhysicalExams.RightKneeRomPopAngle, 
            PhysicalExams.LeftKneeRomPopAngleR1, PhysicalExams.RightKneeRomPopAngleR1, PhysicalExams.LeftKneeRomExt, 
            PhysicalExams.RightKneeRomExt, MotionParams.KneeFlexionMin, MotionParams.HipFlexionMin, 
            MotionParams.AnkleDorsiPlantarPeakDFStance, MotionParams.KneeFlexionMeanStance, 
            MotionParams.KneeFlexionMinIpsiStance, MotionParams.KneeFlexionTotalRange, 
            MotionParams.FootProgressionMeanStance, MotionParams.KneeRotationMeanStance, MotionParams.HipRotationMean, 
            MotionParams.PelvicRotationMean
        FROM (Patients 
            LEFT JOIN MotionParams ON Patients.MRN = MotionParams.MRN) 
            INNER JOIN (((Encounters 
                            INNER JOIN FootMilwaukeeAnglesEtc ON (Encounters.TestNumber = FootMilwaukeeAnglesEtc.TestNumber) 
                                                            AND (Encounters.Date = FootMilwaukeeAnglesEtc.Date) 
                                                            AND (Encounters.MRN = FootMilwaukeeAnglesEtc.MRN)) 
                            INNER JOIN PhysicalExams ON (Encounters.TestNumber = PhysicalExams.TestNumber) 
                                                    AND (Encounters.Date = PhysicalExams.Date) 
                                                    AND (Encounters.MRN = PhysicalExams.MRN)) 
                            INNER JOIN PhysicalExamsMore ON (Encounters.TestNumber = PhysicalExamsMore.TestNumber) 
                                                        AND (Encounters.Date = PhysicalExamsMore.Date) 
                                                        AND (Encounters.MRN = PhysicalExamsMore.MRN)) 
                        ON Patients.MRN = Encounters.MRN
        WHERE Patients.MRN = ?
        """
        # Execute the query and fetch the results into a Pandas DataFrame
        df_q = pd.read_sql(sql_query, conn, params=(mrn_input))
            # For the two columns with dates, remove the timestamp
        df_q['Date'] = pd.to_datetime(df_q['Date']).dt.date
        df_q['DOB'] = pd.to_datetime(df_q['DOB']).dt.date

        diagnosis_label = Label(root, text="Primary Diagnosis:", font=("Open Sans", 12))
        diagnosis_label.place(x=590, y = 150)

        diagnosis = Entry(root, width=15, borderwidth=5, background="#CDCDCD", font=("Open Sans", 12))
        diagnosis.place(x=590, y=170)

        diagnosis_input = df_q.loc[1, 'PrimaryDiagnosis']
        diagnosis.insert(0, diagnosis_input)
        diagnosis.config(state="disable")

        last_input2 = df_q.loc[1, 'LastName']
        first_name.insert(0, last_input2)

        first_input2 = df_q.loc[1, 'FirstName']
        last_name.insert(0, first_input2)

        # Close the cursor and connection
        cursor.close()
        conn.close()

        # Path to save the query exported
        file_path=f"\\\\CHI-FS-APP03\\Motion Analysis Center\\Gait Lab\\_Patients\\{mrn_input}\\{mrn_input}_DT_query.xlsx"
        #file_path=f"G:\\Gait Lab\\_Patients\\{mrn_input}\\{mrn_input}_DT_query.xlsx"

        # Export the DataFrame to an Excel file
        df_q.to_excel(file_path, index=False)

        error_label = Label(root, text="The query is now created", fg="green")
        error_label.place(x=75, y=330)
        ins_label = Label(root, text="Click on the Decision Tree you want\nor to get the Get Deviation Classification:", fg="blue")
        ins_label.place(x=90, y=440)

    elif last and first:
        sql_query="""
        SELECT Patients.MRN, Patients.LastName, Patients.FirstName, Patients.PrimaryDiagnosis, 
            Patients.PatternOfInvolvement, Patients.Side AS Patients_Side, Patients.DOB, Encounters.Date, 
            DateDiff('yyyy',[Patients].[DOB],Date()) AS AgeNow, 
            DateDiff('yyyy',[Patients].[DOB],[Encounters].[Date]) AS EncounterAge, 
            MotionParams.GcdFile, PhysicalExams.GmfcsNew, 
            PhysicalExams.LeftForefootWtBearing, PhysicalExams.RightForefootWtBearing, PhysicalExams.LeftMidfootWtBearing, 
            PhysicalExams.RightMidfootWtBearing, PhysicalExams.LeftHindFootWtBearing, PhysicalExams.RightHindFootWtBearing, 
            PhysicalExamsMore.LeftHalluxWtBearing, PhysicalExamsMore.RightHalluxWtBearing, FootMilwaukeeAnglesEtc.LeftPhiHind, 
            FootMilwaukeeAnglesEtc.RightPhiHind, PhysicalExams.LeftFootSupple, PhysicalExams.RightFootSupple, 
            PhysicalExams.LeftOrthoticTolerated, PhysicalExams.RightOrthoticTolerated, PhysicalExams.LeftToneGastroc, 
            PhysicalExams.RightToneGastroc, PhysicalExams.LeftTonePeroneals, PhysicalExams.RightTonePeroneals, 
            PhysicalExams.LeftAnkleRomInver, PhysicalExams.RightAnkleRomInver, PhysicalExams.LeftAnkleRomEver, 
            PhysicalExams.RightAnkleRomEver, PhysicalExams.LeftThighFootAngle, PhysicalExams.RightThighFootAngle, 
            MotionParams.Side AS MotionParams_Side, MotionParams.AnkleDorsiPlantarMeanStanceDF, MotionParams.KneeFlexionAtIC, 
            PhysicalExams.LeftForefootOffWtBear, PhysicalExams.RightForefootOffWtBear, PhysicalExams.LeftMidfootOffWtBear, 
            PhysicalExams.RightMidfootOffWtBear, PhysicalExams.LeftHindFootOffWtBear, PhysicalExams.RightHindFootOffWtBear, 
            PhysicalExams.LeftFootOther, PhysicalExams.RightFootOther, PhysicalExams.LeftTonePeronealsClonus, 
            PhysicalExams.RightTonePeronealsClonus, PhysicalExams.LeftKneeRomPopAngle, PhysicalExams.RightKneeRomPopAngle, 
            PhysicalExams.LeftKneeRomPopAngleR1, PhysicalExams.RightKneeRomPopAngleR1, PhysicalExams.LeftKneeRomExt, 
            PhysicalExams.RightKneeRomExt, MotionParams.KneeFlexionMin, MotionParams.HipFlexionMin, 
            MotionParams.AnkleDorsiPlantarPeakDFStance, MotionParams.KneeFlexionMeanStance, 
            MotionParams.KneeFlexionMinIpsiStance, MotionParams.KneeFlexionTotalRange, 
            MotionParams.FootProgressionMeanStance, MotionParams.KneeRotationMeanStance, MotionParams.HipRotationMean, 
            MotionParams.PelvicRotationMean
        FROM (Patients 
            LEFT JOIN MotionParams ON Patients.MRN = MotionParams.MRN) 
            INNER JOIN (((Encounters 
                            INNER JOIN FootMilwaukeeAnglesEtc ON (Encounters.TestNumber = FootMilwaukeeAnglesEtc.TestNumber) 
                                                            AND (Encounters.Date = FootMilwaukeeAnglesEtc.Date) 
                                                            AND (Encounters.MRN = FootMilwaukeeAnglesEtc.MRN)) 
                            INNER JOIN PhysicalExams ON (Encounters.TestNumber = PhysicalExams.TestNumber) 
                                                    AND (Encounters.Date = PhysicalExams.Date) 
                                                    AND (Encounters.MRN = PhysicalExams.MRN)) 
                            INNER JOIN PhysicalExamsMore ON (Encounters.TestNumber = PhysicalExamsMore.TestNumber) 
                                                        AND (Encounters.Date = PhysicalExamsMore.Date) 
                                                        AND (Encounters.MRN = PhysicalExamsMore.MRN)) 
                        ON Patients.MRN = Encounters.MRN
        WHERE Patients.LastName = ? AND Patients.FirstName = ?
        """
        # Execute the query and fetch the results into a Pandas DataFrame
        df_q = pd.read_sql(sql_query, conn, params=(first, last))

        diagnosis_label = Label(root, text="Primary Diagnosis:", font=("Open Sans", 12))
        diagnosis_label.place(x=590, y = 150)

        diagnosis = Entry(root, width=25, borderwidth=5, background="#CDCDCD", font=("Open Sans", 12))
        diagnosis.place(x=590, y=180)
        diagnosis_input = df_q.loc[1, 'PrimaryDiagnosis']
        diagnosis.insert(0, diagnosis_input)
        diagnosis.config(state="disable")

        mrn_input2 = df_q.loc[1, 'MRN']
        folder_name.insert(0, mrn_input2)
        
        # For the two columns with dates, remove the timestamp
        df_q['Date'] = pd.to_datetime(df_q['Date']).dt.date
        df_q['DOB'] = pd.to_datetime(df_q['DOB']).dt.date

        # Close the cursor and connection
        cursor.close()
        conn.close()

        # Path to save the query exported
        # file_path=f"G:\\Gait Lab\\_Patients\\{mrn_input}\\{mrn_input}_DT_query.xlsx"
        file_path=f"\\\\CHI-FS-APP03\\Motion Analysis Center\\Gait Lab\\_Patients\\{mrn_input}\\{mrn_input}_DT_query.xlsx"
        # Export the DataFrame to an Excel file
        df_q.to_excel(file_path, index=False)

        error_label = Label(root, text="The query is now created", fg="green")
        error_label.place(x=75, y=330)
        ins_label = Label(root, text="Click on the Decision Tree you want\nor to get the Get Deviation Classification:", fg="blue")
        ins_label.place(x=90, y=440)

def browse_file():
    global error_label, ins_label
    files_name = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if files_name == '':
        messagebox.showerror("Error", "Please choose a file or type an MRN number")
        return
    folder_name.delete(0, tk.END)
    folder_name.insert(0, files_name)
    if error_label:
        error_label.destroy()
    if ins_label:
        ins_label.destroy()

    df = pd.read_excel(files_name)
    if all(col in df.columns for col in ["PatternOfInvolvement", "GmfcsNew", "MotionParams_Side", "EncounterAge", "GcdFile", "Date", "MRN", \
        "RightOrthoticTolerated","RightFootSupple", "RightForefootWtBearing", "RightMidfootWtBearing","RightHindFootWtBearing", "RightHalluxWtBearing", \
            "RightForefootOffWtBear", "RightMidfootOffWtBear", "RightHindFootOffWtBear","RightPhiHind", "RightToneGastroc", "RightTonePeroneals", \
                "RightTonePeronealsClonus", "RightThighFootAngle"]):
        if all(col in df.columns for col in ["AnkleDorsiPlantarPeakDFStance", "KneeFlexionMeanStance", "KneeFlexionTotalRange", "HipFlexionMin", \
            "FootProgressionMeanStance", "KneeRotationMeanStance", "HipRotationMean", "PelvicRotationMean"]):
            if all(col in df.columns for col in ["KneeFlexionAtIC", "KneeFlexionMin", "HipFlexionMin", "MotionParams_Side", "EncounterAge", \
                "AnkleDorsiPlantarMeanStanceDF", "GcdFile", "Date", "MRN", "RightKneeRomPopAngle", "RightKneeRomExt", "RightThighFootAngle", \
                    "RightForefootWtBearing", "RightMidfootWtBearing", "RightHindFootWtBearing"]):
                error_label = Label(root, text="This file contains data for the foot and knee DT\nand to get the Gait Deviation Classification", fg="green")
                error_label.place(x=75, y=320)
                ins_label = Label(root, text="Click on the Knee or Foot DT\nor Gait Deviation:", fg="blue")
                ins_label.place(x=90, y=440)
                buttonFootDT.config(state='normal')
                buttonGaitDT.config(state='normal')
                buttonKneeDT.config(state='normal')
            else:
                error_label = Label(root, text="This file contains data for the foot DT\nand to get the Gait Deviation Classification", fg="green")
                error_label.place(x=75, y=320)
                ins_label = Label(root, text="Click on the Foot DT\nor Gait Deviation:", fg="blue")
                ins_label.place(x=90, y=440)
                buttonFootDT.config(state='normal')
                buttonGaitDT.config(state='normal')
                buttonKneeDT.config(state='disabled')
        else:
            if all(col in df.columns for col in ["KneeFlexionAtIC", "KneeFlexionMin", "HipFlexionMin", "MotionParams_Side", "EncounterAge", \
                "AnkleDorsiPlantarMeanStanceDF", "GcdFile", "Date", "MRN", "RightKneeRomPopAngle", "RightKneeRomExt", "RightThighFootAngle", "RightForefootWtBearing"]):
                error_label = Label(root, text="This file contains data for the foot and knee DT\nthe Gait Deviation needs another file", fg="green")
                error_label.place(x=75, y=320)
                ins_label = Label(root, text="Click on the Knee or Foot DT:", fg="blue")
                ins_label.place(x=90, y=440)
                buttonGaitDT.config(state='disabled')
                buttonFootDT.config(state='normal')
                buttonKneeDT.config(state='normal')
            else:
                error_label = Label(root, text="This file contains data for the foot DT", fg="green")
                error_label.place(x=75, y=320)
                ins_label = Label(root, text="Click on the Foot DT:", fg="blue")
                ins_label.place(x=90, y=440)
    elif all(col in df.columns for col in ["KneeFlexionAtIC", "KneeFlexionMin", "HipFlexionMin", "MotionParams_Side", "EncounterAge", "AnkleDorsiPlantarMeanStanceDF",\
         "GcdFile", "Date", "MRN", "RightKneeRomPopAngle", "RightKneeRomExt", "RightThighFootAngle", "RightForefootWtBearing", "RightMidfootWtBearing", \
            "RightHindFootWtBearing"]):
        error_label = Label(root, text="This file contains data for the knee DT", fg="green")
        error_label.place(x=75, y=320)
        ins_label = Label(root, text="Click on the Knee DT:", fg="blue")
        ins_label.place(x=90, y=440)
        buttonFootDT.config(state='disabled')
        buttonGaitDT.config(state='disabled')
        buttonKneeDT.config(state='normal')
    else:
        error_label = Label(root, text="This file does not contain data to run any decision tree", fg="red")
        error_label.place(x=75, y=320)
        buttonFootDT.config(state='disabled')
        buttonGaitDT.config(state='disabled')
        buttonKneeDT.config(state='disabled')
        if ins_label:
            ins_label.destroy()

# Following the Knee Flexion Deformity decision diagram 
def KneeFlexDefor():
    global error_label, ins_label, button_export1, button_export2, button_criteria, knee_recommend, selected_dates2, instructions_label
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

    #start of the function
    recommendation_str = ""
    explanation_str = ""

    fold_name = folder_name.get()
    if fold_name == "":
        messagebox.showerror("Missing data", "There was not MRN typed or any browsed file\nPlease try again")
        return
    elif fold_name.isdigit():
        file_name = str(fold_name) + "_DT_query.xlsx"
        folder_path = os.path.join(folder_location, str(fold_name))
        file_path = os.path.join(folder_path, file_name)
    else:
        file_path= fold_name

    df = pd.read_excel(file_path) # read the excel file

    if 'complete_query' in fold_name:
        num_rows = len(df.index)
        filtered_rows = range(num_rows)
    else:
        # Select dates
        selected_dates2 = select_dates(df)
        df['Date'] = pd.to_datetime(df['Date']).dt.date
        filtered_rows = df[df['Date'].isin(selected_dates2)].index

    title_label.config(text="Knee Recommendations", fg="black", font=('Open Sans - Bold', 16))

    for row_index in filtered_rows:
        # global variables for this part
        FCKneeFlex = df.loc[row_index, 'KneeFlexionAtIC'] # equivalent to 'foot contact knee flexion'
        MidsKneeFlex = df.loc[row_index, 'KneeFlexionMin'] # equivalent to 'midstance knee flexion'
        HipFlexion = df.loc[row_index, 'HipFlexionMin'] # equivalent to 'hip flexion'
        Side = df.loc[row_index,'MotionParams_Side'] # to determine which variables to condiser for the DT
        EncounterAge = df.loc[row_index, 'EncounterAge']
        Ankle = df.loc[row_index, 'AnkleDorsiPlantarMeanStanceDF']
        Trial = df.loc[row_index, 'GcdFile'] # trial number that is in the GcdFile column
        Date = df.loc[row_index, 'Date']
        MRN = int(df.loc[row_index, 'MRN'])
        e_date = str(Date).split()
        if len(e_date) > 1:
            t_date = None
            Date_E = str(e_date[0]) 
        else:
            Date_E = str(e_date[0]) 

        # variables for right side
        if 'R' in str(Side).upper():
            pop1 = df.loc[row_index, 'RightKneeRomPopAngle']
            if pd.isna(pop1):
                Pop_angle = 1
            else:
                match = re.search(r'-?\d+(\.\d+)?', str(pop1))  # Extract the first sequence of digits (with optional decimal part) from the cell value
                if match:
                    Pop_angle = float(match.group())  # Convert the extracted value to a float
                else:
                    Pop_angle = 1
            
            
            knee1 = df.loc[row_index, 'RightKneeRomExt']
            if pd.isna(knee1):
                KneeFlexContr = 1
            else:
                match = re.search(r'[-+]?\d+(\.\d+)?', str(knee1))
                if match:
                    KneeFlexContr = float(match.group())
                else:
                    KneeFlexContr = 1


            ThighFoot = df.loc[row_index, 'RightThighFootAngle']
            if pd.isna(ThighFoot):
                number_thigh = -1
                ext_int_thigh = 'not'
            else:
                value_thigh_foot = str(ThighFoot).split()

                if len(value_thigh_foot) == 1:
                    match = re.search(r'(\d+)(\D+)?', value_thigh_foot[0])
                    if match:
                        try:
                            number_thigh = int(match.group(1))
                            ext_int_thigh = match.group(2) if match.group(2) else 'not'
                        except ValueError:
                            number_thigh = -1
                            ext_int_thigh = 'not'
                    else:
                        try:
                            number_thigh = int(value_thigh_foot[0])
                            ext_int_thigh = 'not'
                        except ValueError:
                            number_thigh = -1
                            ext_int_thigh = 'not'
                elif len(value_thigh_foot) == 2:
                    number_thigh = int(value_thigh_foot[0])
                    ext_int_thigh = value_thigh_foot[1]
                else:
                    number_thigh = -1
                    ext_int_thigh = 'not'

            forefoot = df.loc[row_index, 'RightForefootWtBearing']
            midfoot = df.loc[row_index, 'RightMidfootWtBearing']
            hindfoot = df.loc[row_index, 'RightHindFootWtBearing']
            text_side = 'Right side'

        # variables for left side
        elif 'L' in str(Side).upper():
            pop1 = df.loc[row_index, 'LeftKneeRomPopAngle']
            if pd.isna(pop1):
                Pop_angle = 1
            else:
                match = re.search(r'-?\d+(\.\d+)?', str(pop1))  # Extract the first sequence of digits (with optional decimal part) from the cell value
                if match:
                    Pop_angle = float(match.group())  # Convert the extracted value to a float
                else:
                    Pop_angle = 1


            knee1 = df.loc[row_index, 'LeftKneeRomExt']
            if pd.isna(knee1):
                KneeFlexContr = 1
            else:
                match = re.search(r'[-+]?\d+(\.\d+)?', str(knee1))
                if match:
                    KneeFlexContr = float(match.group())
                else:
                    KneeFlexContr = 1

            ThighFoot = df.loc[row_index, 'LeftThighFootAngle']
            if pd.isna(ThighFoot):
                number_thigh = -1
                ext_int_thigh = 'not'
            else:
                value_thigh_foot = str(ThighFoot).split()

                if len(value_thigh_foot) == 1:
                    match = re.search(r'(\d+)(\D+)?', value_thigh_foot[0])
                    if match:
                        try:
                            number_thigh = int(match.group(1))
                            ext_int_thigh = match.group(2) if match.group(2) else 'not'
                        except ValueError:
                            number_thigh = -1
                            ext_int_thigh = 'not'
                    else:
                        try:
                            number_thigh = int(value_thigh_foot[0])
                            ext_int_thigh = 'not'
                        except ValueError:
                            number_thigh = -1
                            ext_int_thigh = 'not'
                elif len(value_thigh_foot) == 2:
                    number_thigh = int(value_thigh_foot[0])
                    ext_int_thigh = value_thigh_foot[1]
                else:
                    number_thigh = -1
                    ext_int_thigh = 'not'

            forefoot = df.loc[row_index, 'LeftForefootWtBearing']
            midfoot = df.loc[row_index, 'LeftMidfootWtBearing']
            hindfoot = df.loc[row_index, 'LeftHindFootWtBearing']
            text_side = 'Left side'
        
        # if there's no L or R it wont go to the decision tree
        else: 
            recommendation_str += "The file does not specify L or R side.\n"
            Pop_angle = None
            KneeFlexContr = None
            ThighFoot = None
            value_thigh_foot = None
            text_side = None
            continue

        forefoot_10 = ['abd', 'pronation', 'valgus']
        midfoot_10 = ['low med arch', 'midfoot break', 'planovalgus', 'planus', 'valgus', 'pronation', 'rockr bttm']
        hindfoot_10 = ['valgus']

        # First Age Group < 5
        if EncounterAge < 5:
            # Branch and recommendation 1
            if Pop_angle < 60:
                if Ankle <= 2:
                    recommend1 = 'Do botulinum injection & knee splinting, may repeat 4 - 5 times if still getting a positive effect.'
                    recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) + ", Trial: " + str(Trial) + ", Recommendation: " + recommend1 + "\n"
                    explanation_str += "The criteria for this recommendation is having a popliteal angle less than 60 degrees\nIn this trial the patient has a popliteal angle of " + str(Pop_angle) + " degrees.\n\n"
                    knee_recommend.append(recommend1)
                else:
                    recommend1 = 'No recommendation.'
                    recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) + ", Trial: " + str(Trial) + ", " + recommend1 + "\n"
                    explanation_str += "The criteria for having a recommendation in patients under 5 years old is:\n - Having a popliteal angle less than 60 degrees \n - Having a popliteal angle greater than 60 degrees and a knee flexion contracture greater than 10 degrees. \nIn this trial the patient has:\n - Popliteal angle of " + str(Pop_angle) + " degrees\n - Knee flexion contracture of " + str(KneeFlexContr) + " degrees \n\n"
                    knee_recommend.append(recommend1)

            # Branch and recommendation 2
            elif Pop_angle >= 60:
                if KneeFlexContr <= -10:
                    recommend1 = 'Do hamstring lengthening.'
                    recommendation_str+= "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) + ", Trial: " + str(Trial) + ", Recommendation: " + recommend1 + "\n"
                    explanation_str += "The criteria for this recommendation is having a popliteal angle greater than 60 degrees and more than 10 degrees of knee flexion contracture.\nIn this trial the patient has:\n - Popliteal angle of " + str(Pop_angle) + " degrees\n - Knee flexion contracture of " + str(KneeFlexContr) + " degrees.\n\n"
                    knee_recommend.append(recommend1)

            # No matches -> no recommendation 
                else: 
                    recommend1 = 'No recommendation.'
                    recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) + ", Trial: " + str(Trial) + ", " + recommend1 + "\n"
                    explanation_str += "The criteria for having a recommendation in patients under 5 years old is:\n - Having a popliteal angle less than 60 degrees \n - Having a popliteal angle greater than 60 degrees and a knee flexion contracture greater than 10 degrees. \nIn this trial the patient has:\n - Popliteal angle of " + str(Pop_angle) + " degrees\n - Knee flexion contracture of " + str(KneeFlexContr) + " degrees \n\n"
                    knee_recommend.append(recommend1)
        
            # Second Age Group 5 - 10
        if (EncounterAge >= 5) and (EncounterAge <= 10):
            FCKnee = round(FCKneeFlex, 2)
            MidsKnee = round(MidsKneeFlex, 2)
                # Branch and recommendation 3
            if (KneeFlexContr < -20) and (KneeFlexContr >= -40):
                recommend1 = 'Do posterior knee capsulotomy & hamstring lengthening'
                recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) + ", Trial: " + str(Trial) + ", Recommendation: " + recommend1 +"\n"
                explanation_str += "The criteria for this recommendation is having a knee flexion contracture between 20 and 40 degrees\nIn this trial the patient has a knee flexion contracture of " + str(KneeFlexContr) + " degrees.\n\n"
                knee_recommend.append(recommend1)

            # Branch and recommendation 4
            elif (KneeFlexContr >= -20 and KneeFlexContr <= 0) and (FCKneeFlex >= 25) and (Pop_angle >= 50) and (MidsKneeFlex >= 25):
                recommend1 = 'Do hamstring lengthening.'
                recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) + ", Trial: " + str(Trial) + ", Recommendation: " + recommend1 +"\n"
                explanation_str += "The criteria for this recommendation is having a foot contact knee flexion greater than 25 degrees, popliteal angle greater than 50 degrees, fixed knee flexion contracture less than 20 degrees and midstance knee flexion greater than 25 degrees.\nIn this trial the patient has:\n - Foot contact knee flexion of " + str(FCKnee) + " degrees.\n - Popliteal angle of " + str(Pop_angle) +" degrees\n - Knee flexion contracture of " + str(KneeFlexContr) +" degrees\n - Midstance knee flexion of " + str(MidsKnee) +" degrees.\n\n"
                knee_recommend.append(recommend1)

           # No matches -> no recommendation 
            else: 
                recommend1 = 'No recommendation.'
                recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) + ", Trial: " + str(Trial) + ", " + recommend1 + "\n"
                explanation_str += "The criteria for having a recommendation in patients between 5 and 10 years old is:\n - Having a knee flexion contracture between 20 and 40 degrees\n - Having foot contact knee flexion greater than 25 degrees, popliteal angle greater than 50 degrees, knee flexion contracture less than 20 degrees and midstance knee flexion greater than 25 degrees.\nIn this trial the patients has:\n - Foot contact knee flexion of " + str(FCKnee) + " degrees.\n - Popliteal angle of " + str(Pop_angle) +" degrees\n - Knee flexion contracture of " + str(KneeFlexContr) +" degrees\n - Midstance knee flexion of " + str(MidsKnee) +" degrees.\n\n"
                knee_recommend.append(recommend1)

        # Third Age Group > 10
        if EncounterAge > 10:
            FCKnee = round(FCKneeFlex, 2)
            MidsKnee = round(MidsKneeFlex, 2)
            # Branch and recommendation 8
            if MidsKneeFlex < 0:
                recommend1 = 'Do very careful correction of ankle equinus.'
                recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) + ", Trial: " + str(Trial) + ", Recommendation: " \
                    + recommend1 + "\n"
                explanation_str += "The criteria for this recommendation is having a midstance knee klexion less than 0 degrees\nIn this trial the patient has:"+\
                    "\n - Midstance knee flexion of " + str(MidsKnee) +" degrees.\n\n"
                knee_recommend.append(recommend1)
            else:
                if HipFlexion <= -2:
                    if (FCKneeFlex >= 25) and (Pop_angle >= 50):
                        if KneeFlexContr >= -10:
                            recommend1 = 'Do hamstring lengthening and correct crouch causes.'
                            recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) + ", Trial: " + str(Trial) + \
                                ", Recommendation: "+ recommend1 + "\n"
                            explanation_str += "The criteria for this recommendation is having a foot contact knee flexion greater than 25 degrees, popliteal "+\
                                "angle greater than 50 degrees and a knee flexion contracture less than 10 degrees.\nIn this trial the patient has:\n - Foot contact "+\
                                    "knee flexion of " + str(FCKnee) + " degrees.\n - Popliteal angle of " + str(Pop_angle) +" degrees\n - Knee flexion contracture of "+\
                                         str(KneeFlexContr) +" degrees.\n\n"
                            knee_recommend.append(recommend1)
                        elif KneeFlexContr >= -30 and KneeFlexContr <= -10:
                            recommend1 = 'Do posterior knee capsulotomy & correct all other elements of crouch.'
                            recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) + ", Trial: " + str(Trial) +\
                                 ", Recommendation: " + recommend1 +  "\n"
                            explanation_str += "The criteria for this recommendation is having crouched gait with knee flexion contracture greater than 10 degrees "+\
                                "but less than 30 degrees\nIn this trial the patient has:\n - Knee flexion contracture of " + str(KneeFlexContr) +" degrees.\n\n"
                            knee_recommend.append(recommend1)
                        elif KneeFlexContr < -30:
                            recommend1 = 'Do knee extension osteotomy & correct all other crouch causes.'
                            recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) + ", Trial: " + str(Trial) + \
                                ", Recommendation: " + recommend1 + "\n"
                            explanation_str += "The criteria for this recommendation is having crouched gait with knee flexion contracture greater than 30 degrees"+\
                                "\nIn this trial the patient has:\n - Knee flexion contracture of " + str(KneeFlexContr) +" degrees.\n\n"
                            knee_recommend.append(recommend1)
                        else: 
                            recommend1 = 'No recommendation.'
                            recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) + ", Trial: " + str(Trial) + ", " +\
                                 recommend1 + "\n"
                            explanation_str += "The criteria for having a recommendation in patients between 5 and 10 years old is:\n - Having a foot contact knee "+\
                                "greater than 25 degrees, popliteal angle greater than 50 degrees and a knee flexion contracture less than 10 degrees.\n - Having "+\
                                "crouched gait with knee flexion contracture greater than 10 degrees but less than 30 degrees\n - Having crouched gait with knee "+\
                                "flexion contracture greater than 30 degrees.\n - Having a midstance knee klexion less than 0 degrees\n - Having increased hip flexion "\
                                "(> -2 degrees) with planovalgus, ankle equinus and external tibial torsion (>20 degrees).\nIn this trial the patient has:\n - Foot "+\
                                "contact knee flexion of " + str(FCKnee) + " degrees.\n - Popliteal angle of " + str(Pop_angle) +" degrees\n - Knee flexion contracture of " +\
                                str(KneeFlexContr) +" degrees\n - " + str(ext_int_thigh) + " tibial torsion of " + str(number_thigh) + " degrees\n\n"
                            knee_recommend.append(recommend1)
                    else:
                        recommend1 = 'No recommendation.'
                        recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) + ", Trial: " + str(Trial) + ", " + recommend1 + "\n"
                        explanation_str += "The criteria for having a recommendation in patients between 5 and 10 years old is:\n - Having a foot contact knee greater than "+\
                            "25 degrees, popliteal angle greater than 50 degrees and a knee flexion contracture less than 10 degrees.\n - Having crouched gait with knee "+\
                            "flexion contracture greater than 10 degrees but less than 30 degrees\n - Having crouched gait with knee flexion contracture greater than 30 "+\
                            "degrees.\n - Having a midstance knee klexion less than 0 degrees\n - Having increased hip flexion (> -2 degrees) with planovalgus, ankle "+\
                            "equinus and external tibial torsion (>20 degrees).\nIn this trial the patient has:\n - Foot contact knee flexion of " + str(FCKnee) + " degrees."+\
                            "\n - Popliteal angle of " + str(Pop_angle) +" degrees\n - Knee flexion contracture of " + str(KneeFlexContr) +" degrees\n - " + str(ext_int_thigh)+\
                            " tibial torsion of "+ str(number_thigh) + " degrees\n\n"
                        knee_recommend.append(recommend1)
                else:
                    if (any(word in str(forefoot) for word in forefoot_10) or any(word in str(midfoot) for word in midfoot_10) or any(word in str(hindfoot) for word \
                        in hindfoot_10)) and (Ankle <= 2) and ((number_thigh >= 20) and ('E' in ext_int_thigh.upper())):
                        recommend1 = 'Do hip extension osteotomy'
                        if (FCKneeFlex >= 25) and (Pop_angle >= 50):
                            if KneeFlexContr >= -10:
                                recommend1 = recommend1 + ' AND Do hamstring lengthening and correct crouch causes.'
                                recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  ", Trial: " + str(Trial) + \
                                ",  Recommendation: " + recommend1 + "\n"
                                explanation_str += "The criteria for the hip recommendation is having increased hip flexion (> -2 degrees) with planovalgus, ankle "+\
                                    "equinus and external tibial torsion (>20 degrees).\nIn this trial the patient has:\n - Hip flexion of " + str(HipFlexion) +\
                                    "degrees\n - Tibial torsion of " + str(number_thigh) + "degrees\n The criteria for the knee recommendation is having a foot contact "+\
                                    "knee flexion greater than 25 degrees, popliteal angle greater than 50 degrees and a knee flexion contracture less than 10 degrees."+\
                                    "\nIn this trial the patient has:\n - Foot contact knee flexion of " + str(FCKnee) + " degrees.\n - Popliteal angle of " +\
                                    str(Pop_angle) +" degrees\n - Knee flexion contracture of " + str(KneeFlexContr) +" degrees.\n\n"
                                knee_recommend.append(recommend1)
                            elif KneeFlexContr >= -30 and KneeFlexContr <= -10:
                                recommend1 = recommend1 + ' AND Do posterior knee capsulotomy & correct all other elements of crouch.'
                                recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  ", Trial: " + str(Trial) +\
                                    ",  Recommendation: " + recommend1 + "\n"
                                explanation_str += "The criteria for the hip recommendation is having increased hip flexion (> -2 degrees) with planovalgus, ankle "+\
                                    "equinus and external tibial torsion (>20 degrees).\nIn this trial the patient has:\n - Hip flexion of " + str(HipFlexion) +\
                                    "degrees\n - Tibial torsion of " + str(number_thigh) + "degrees\n The criteria for the knee recommendation is having crouched gait "+\
                                    "with knee flexion contracture greater than 10 degrees but less than 30 degrees\nIn this trial the patient has:\n - Knee flexion "+\
                                    "contracture of " + str(KneeFlexContr) +" degrees.\n\n"
                                knee_recommend.append(recommend1)
                            elif KneeFlexContr < -30:
                                recommend1 = recommend1 + ' AND Do knee extension osteotomy & correct all other crouch causes.'
                                recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  ", Trial: " + str(Trial) +\
                                    ",  Recommendation: " + recommend1 + "\n"
                                explanation_str += "The criteria for the hip recommendation is having increased hip flexion (> -2 degrees) with planovalgus, ankle "+\
                                    "equinus and external tibial torsion (>20 degrees).\nIn this trial the patient has:\n - Hip flexion of " + str(HipFlexion) +\
                                    "degrees\n - Tibial torsion of " + str(number_thigh) + "degrees\n The criteria for the knee recommendation is having crouched gait "+\
                                    "with knee flexion contracture greater than 30 degrees\nIn this trial the patient has:\n - Knee flexion contracture of " + \
                                    str(KneeFlexContr) +" degrees.\n\n"
                                knee_recommend.append(recommend1)
                            else:
                                recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  ", Trial: " + str(Trial) + \
                                    ",  Recommendation: " + recommend1 + "\n"
                                explanation_str += "The criteria for this recommendation is having increased hip flexion (> -2 degrees) with planovalgus, ankle "+\
                                    "equinus and external tibial torsion (>20 degrees).\nIn this trial the patient has:\n - Hip flexion of " + str(HipFlexion) + \
                                    "degrees\n - Tibial torsion of " + str(number_thigh) + "degrees\n\n"
                                knee_recommend.append(recommend1)
                        else:
                            recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  ", Trial: " + str(Trial) + \
                                ",  Recommendation: " + recommend1 + "\n"
                            explanation_str += "The criteria for this recommendation is having increased hip flexion (> -2 degrees) with planovalgus, ankle equinus "+\
                                "and external tibial torsion (>20 degrees).\nIn this trial the patient has:\n - Hip flexion of " + str(HipFlexion) + "degrees\n - Tibial "+\
                                "torsion of " + str(number_thigh) + "degrees\n\n"
                            knee_recommend.append(recommend1)
                    else:
                        if (FCKneeFlex >= 25) and (Pop_angle >= 50):
                            if KneeFlexContr >= -10:
                                recommend1 = 'No hip recommendation, Knee recommendation: Do hamstring lengthening and correct crouch causes.'
                                recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) + ", Trial: " + str(Trial) + ", " +\
                                    recommend1 + "\n"
                                explanation_str += "The criteria for this recommendation is having a foot contact knee flexion greater than 25 degrees, popliteal angle "+\
                                    "greater than 50 degrees and a knee flexion contracture less than 10 degrees.\nIn this trial the patient has:\n - Foot contact knee "+\
                                    "flexion of " + str(FCKnee) + " degrees.\n - Popliteal angle of " + str(Pop_angle) +" degrees\n - Knee flexion contracture of " +\
                                    str(KneeFlexContr) +" degrees.\n\n"
                                knee_recommend.append(recommend1)
                            elif KneeFlexContr >= -30 and KneeFlexContr <= -10:
                                recommend1 = 'No hip recommendation, Knee recommendation: Do posterior knee capsulotomy & correct all other elements of crouch.'
                                recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) + ", Trial: " + str(Trial) + ", " +\
                                    recommend1 + "\n"
                                explanation_str += "The criteria for this recommendation is having crouched gait with knee flexion contracture greater than 10 degrees "+\
                                    "but less than 30 degrees\nIn this trial the patient has:\n - Knee flexion contracture of " + str(KneeFlexContr) +" degrees.\n\n"
                                knee_recommend.append(recommend1)
                            elif KneeFlexContr < -30:
                                recommend1 = 'No hip recommendation, Knee recommendation: Do knee extension osteotomy & correct all other crouch causes.'
                                recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) + ", Trial: " + str(Trial) + ", " +\
                                    recommend1 + "\n"
                                explanation_str += "The criteria for this recommendation is having crouched gait with knee flexion contracture greater than 30 degrees"+\
                                    "\nIn this trial the patient has:\n - Knee flexion contracture of " + str(KneeFlexContr) +" degrees.\n\n"
                                knee_recommend.append(recommend1)
                            else: 
                                recommend1 = 'No hip and no knee recommendation.'
                                recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) + ", Trial: " + str(Trial) + ", " +\
                                    recommend1 + "\n"
                                explanation_str += "The criteria for having a recommendation in patients between 5 and 10 years old is:\n - Having a foot contact knee "+\
                                    "greater than 25 degrees, popliteal angle greater than 50 degrees and a knee flexion contracture less than 10 degrees.\n - Having "+\
                                    "crouched gait with knee flexion contracture greater than 10 degrees but less than 30 degrees\n - Having crouched gait with knee "+\
                                    "flexion contracture greater than 30 degrees.\n - Having a midstance knee klexion less than 0 degrees\n - Having increased hip flexion "+\
                                    "(> -2 degrees) with planovalgus, ankle equinus and external tibial torsion (>20 degrees).\nIn this trial the patient has:\n - Foot "+\
                                    "contact knee flexion of " + str(FCKnee) + " degrees.\n - Popliteal angle of " + str(Pop_angle) +" degrees\n - Knee flexion contracture "+\
                                    "of " + str(KneeFlexContr) +" degrees\n - " + str(ext_int_thigh) + " tibial torsion of " + str(number_thigh) + " degrees\n\n"
                                knee_recommend.append(recommend1)
                        else:
                            recommend1 = 'No hip recommendation, No knee recommendation'
                            recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) + ", Trial: " + str(Trial) + ", " +\
                                recommend1 + "\n"
                            explanation_str += "The criteria for having a recommendation in patients between 5 and 10 years old is:\n - Having a foot contact knee greater "+\
                                "than 25 degrees, popliteal angle greater than 50 degrees and a knee flexion contracture less than 10 degrees.\n - Having crouched gait with "+\
                                "knee flexion contracture greater than 10 degrees but less than 30 degrees\n - Having crouched gait with knee flexion contracture greater "+\
                                "than 30 degrees.\n - Having a midstance knee klexion less than 0 degrees\n - Having increased hip flexion (> -2 degrees) with planovalgus, "+\
                                "ankle equinus and external tibial torsion (>20 degrees).\nIn this trial the patient has:\n - Foot contact knee flexion of " + str(FCKnee) +\
                                " degrees.\n - Popliteal angle of " + str(Pop_angle) +" degrees\n - Knee flexion contracture of " + str(KneeFlexContr) +" degrees\n - "\
                                + str(ext_int_thigh) + " tibial torsion of " + str(number_thigh) + " degrees\n\n"
                            knee_recommend.append(recommend1)
        #else:
            #error_label.config(text="The Excel file for this patient has not been created.")

    recommendation_text.delete('1.0', END)
    # add format to the recommendation text displayed
    recommendation_text.insert(END, recommendation_str, ('tag1', 'color_blue', 'space1'))
    recommendation_text.tag_configure('tag1', font=('Open Sans', 12))
    recommendation_text.tag_configure('color_blue', foreground='black')
    recommendation_text.tag_configure('space1', spacing3=15)

    def export_Excel(knee_recommend, df):
        df['Freeman Miller Knee Recommendations'] = knee_recommend

        export_file_path = filedialog.asksaveasfilename(defaultextension='.xlsx')
        df.to_excel(export_file_path, index=False)

    # function to export the recommendation report to a pdf file
    def export_pdf_full(recommendation_str, explanation_str, MRN, selected_dates2):
        messagebox.showinfo("Save report", "Select the folder to export the report\n\nThe report is in pdf format")
        folder_pathpdf = filedialog.askdirectory()
        if not folder_pathpdf:
            return
        dates = [str(date.year) for date in selected_dates2]
        # join dates with underscore
        joined_dates = '_'.join(dates)
        filename=f"{folder_pathpdf}/" + str(MRN) + "_KneeRec_"+ joined_dates +"_full_report.pdf"

        c=canvas.Canvas(filename, pagesize=letter)
        max_width = 95
        page_height = c._pagesize[1] - 5

        recommendation_lines = recommendation_str.split("\n")
        explanation_lines = explanation_str.split("\n\n")

        y = page_height - 30
        add_extra_text = True
        # add 1 recommendation at a time
        for i in range(len(recommendation_lines)):
            if add_extra_text and i == 0:
                y -= 5 # space to add the header
                c.setFont("Helvetica-Bold", 16)
                c.drawString(135, y, "Recommendations from the Freeman Miller")
                c.drawString(180, y-20, "Knee Treatment Decision Tree")
                y -= 60
                add_extra_text = False

            c.setFont("Helvetica", 12)
            c.setFillColorRGB(0, 0, 1) # blue color
            wrapped_recommendation = textwrap.wrap(recommendation_lines[i], width = max_width)
           
            num_lines_rec = len(wrapped_recommendation)
            if y - num_lines_rec*20 < 35:
                
                c.showPage()
                y = page_height - 30
                ##
            c.setFont("Helvetica", 12)
            c.setFillColorRGB(0, 0, 1) # blue color
            for line in wrapped_recommendation:
                c.drawString(50, y, line)
                y -= 20
            
            # check the space left in the page to avoid cropped information
            exp_lines = explanation_lines[i].split('\n')
            num_lines = sum(len(textwrap.wrap(line, width=max_width+5)) for line in exp_lines)
            if y - num_lines*20 < 35:
                c.showPage()
                y = page_height - 30
            
            # show every explanation for each recommendation
            c.setFont("Helvetica", 10)
            for line in exp_lines:
                wrapped_explanation = textwrap.wrap(line, width=(max_width+15))
                for exp_line in wrapped_explanation:
                    c.setFillColorRGB(0, 0, 0) # black color
                    c.drawString(70, y, exp_line.strip())
                    y -= 20
            y -= 20

        c.save()
        messagebox.showinfo("Export", "Report exported")

    # function to export the recommendation report to a pdf file
    def export_pdf(recommendation_str, MRN, selected_dates2):
        messagebox.showinfo("Save report", "Select the folder to export the report\n\nThe report is in pdf format")
        folder_pathpdf = filedialog.askdirectory() # maybe change this to also let the user choose the name?
        if not folder_pathpdf:
            return
        dates = [str(date.year) for date in selected_dates2]

        # join dates with underscore
        joined_dates = '_'.join(dates)

        filename=f"{folder_pathpdf}/" + str(MRN) + "_KneeRec_"+ joined_dates +"_report.pdf"

        c=canvas.Canvas(filename, pagesize=letter)
        max_width = 115 # this control the horizontal size in which the text can be displayed
        page_height = c._pagesize[1] - 5
        recommendation_lines = recommendation_str.split("\n")

        y = page_height - 30
        add_extra_text = True # verify that there is enough space left in the same page for the full recommendation or explanation
        for i in range(len(recommendation_lines)):
            if add_extra_text and i == 0:
                y -= 5 # space to add the header
                c.setFont("Helvetica-Bold", 16)
                c.drawString(135, y, "Recommendations from the Freeman Miller")
                c.drawString(180, y-20, "Knee Treatment Decision Tree")
                y -= 60
                add_extra_text = False

            c.setFont("Helvetica", 10)
            c.setFillColorRGB(0, 0, 0) # black color
            wrapped_recommendation = textwrap.wrap(recommendation_lines[i], width = max_width)
            for line in wrapped_recommendation:
                c.drawString(50, y, line)
                y -= 12
            
            # check the space left in the page to avoid cropped information
            if y < 40:
                c.showPage()
                y = page_height - 30
            y -= 15
            
        c.save()
        messagebox.showinfo("Export", "Report exported")

    # button to export the recommendations and criteria in a pdf file
    button_export1=Button(root, text="Export\nreport\n(with criteria)", font=("Open Sans", 10), bg="#BFBFBF", padx=10, pady=8, command= lambda: \
        export_pdf_full(recommendation_str, explanation_str, MRN, selected_dates2))
    button_export1.place(x=1230, y=190)

    # button to export Excel file
    button_export2=Button(root, text="Export\nExcel\nfile", font=("Open Sans", 10), bg="#BFBFBF", padx=15, pady=8, command= lambda: export_Excel(knee_recommend, df))
    button_export2.place(x=1100, y=190)

    # button for the pop up window with the criteria
    button_criteria=Button(root, text="Recommendation\nCriteria", font=("Open Sans", 12), bg="#D9E1F2", padx=10, pady=8, command=lambda: \
        show_popup(recommendation_str, explanation_str))
    button_criteria.place(x=850, y=195)

    # pop up window that shows the criteria used for each recommendation
    def show_popup(recommendation_str, explanation_str):
        # create a child window
        popup_window = Toplevel(root)
        popup_window.geometry("950x575+700+350")
        popup_window.title("Recommendation Criteria")

        popup_window.resizable(False,False)
        # create a Frame to hold the Text widgets and scrollbar
        frame = Frame(popup_window)
        frame.place(x=10, y=10)

        # create a scrollbar for the Text widgets
        vscrollbar = Scrollbar(frame)
        vscrollbar.pack(side=RIGHT, fill=Y)

        # create a horizontal scrollbar for the Text widget
        hscrollbar = Scrollbar(frame, orient=HORIZONTAL)
        hscrollbar.pack(side=BOTTOM, fill=X)

        # create a Text widget to hold the recommendation text
        recommendation_text = Text(frame, height=25, width=115, wrap=NONE, yscrollcommand=vscrollbar.set, xscrollcommand=hscrollbar.set, spacing3=5)
        recommendation_text.pack(side=LEFT, fill=Y)

        # cet the scrollbar to control the Text widget
        vscrollbar.config(command=recommendation_text.yview)

        # cet the horizontal scrollbar to control the Text widget
        hscrollbar.config(command=recommendation_text.xview)

        # add the recommendation and explanation text to the new window
        recommendation_text.delete('1.0', END)

        lines = recommendation_str.split("\n")

        #use a loop to add 1 recommendation and 1 explanation at a time
        for i in range(len(lines)):
            recommendation_text.insert(END, lines[i] + "\n", ('tag1', 'color_blue'))
            recommendation_text.tag_configure('tag1', font=('Open Sans', 12))
            recommendation_text.tag_configure('color_blue', foreground='blue')
            if i<len(explanation_str.split("\n\n")):
                recommendation_text.insert(END, explanation_str.split("\n\n")[i] + "\n\n", ('tag2', 'color_black'))
                recommendation_text.tag_configure('tag2', font=('Open Sans', 10))
                recommendation_text.tag_configure('color_black', foreground='black')

# Following the Foot Decision Tree
def Foot():
    global error_label, ins_label, button_export1, button_export2, button_criteria, selected_dates2, instructions_label
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

    #start of the function
    recommendation_str = ""
    explanation_str = ""

    fold_name = folder_name.get()
    if fold_name == "":
        messagebox.showerror("Missing data", "There was not MRN typed or any browsed file\nPlease try again")
        return
    elif fold_name.isdigit():
        file_name = str(fold_name) + "_DT_query.xlsx"
        folder_path = os.path.join(folder_location, str(fold_name))
        file_path = os.path.join(folder_path, file_name)
    else:
        file_path= fold_name

    df = pd.read_excel(file_path) # read the excel file

    if 'complete_query' in fold_name:
        num_rows = len(df.index)
        filtered_rows = range(num_rows)
    else:
        # Select dates
        selected_dates2 = select_dates(df)
        df['Date'] = pd.to_datetime(df['Date']).dt.date
        filtered_rows = df[df['Date'].isin(selected_dates2)].index

    title_label.config(text="Foot Recommendations", fg="black", font=('Open Sans - Bold', 16))

    def get_user_R_L():
        global var1
        global var2
        # Create the pop-up window for the right side
        popup1 = tk.Toplevel()
        popup1.title("Grade of deformity")
        popup1.geometry("300x300+850+450")

        # Create the label and radio buttons
        label1 = tk.Label(popup1, text="For the Left side\nSelect the grade of the deformity:")
        label1.pack(padx=10, pady=10)

        var1 = tk.StringVar(value="No deformity")
        no_def_button1 = tk.Radiobutton(popup1, text="No deformity", variable=var1, value="No deformity")
        no_def_button1.pack(padx=10, pady=5)

        mild_button1 = tk.Radiobutton(popup1, text="Mild", variable=var1, value="Mild")
        mild_button1.pack(padx=10, pady=5)

        moderate_button1 = tk.Radiobutton(popup1, text="Moderate", variable=var1, value="Moderate")
        moderate_button1.pack(padx=10, pady=5)

        severe_button1 = tk.Radiobutton(popup1, text="Severe", variable=var1, value="Severe")
        severe_button1.pack(padx=10, pady=5)

        # Create the OK button and set focus to it
        ok_button1 = tk.Button(popup1, text="Save", padx=20, pady=10, command=popup1.destroy)
        ok_button1.pack(padx=10, pady=10)
        ok_button1.focus_set()

        # Bind the Enter key to the OK button
        popup1.bind("<Return>", lambda event: ok_button1.invoke())

        # Set focus to the default button
        if default_choice1 == "No deformity":
            no_def_button1.focus_set()
        elif default_choice1 == "Mild":
            mild_button1.focus_set()
        elif default_choice1 == "Moderate":
            moderate_button1.focus_set()
        elif default_choice1 == "Severe":
            severe_button1.focus_set()

        # Create the pop-up window
        popup2 = tk.Toplevel()
        popup2.title("Grade of deformity")
        popup2.geometry("300x300+1200+450")

        # Create the label and radio buttons
        label2 = tk.Label(popup2, text="For the Right side\nSelect the grade of the deformity:")
        label2.pack(padx=10, pady=10)

        var2 = tk.StringVar(value="No deformity")
        no_def_button2 = tk.Radiobutton(popup2, text="No deformity", variable=var2, value="No deformity")
        no_def_button2.pack(padx=10, pady=5)

        mild_button2 = tk.Radiobutton(popup2, text="Mild", variable=var2, value="Mild")
        mild_button2.pack(padx=10, pady=5)

        moderate_button2 = tk.Radiobutton(popup2, text="Moderate", variable=var2, value="Moderate")
        moderate_button2.pack(padx=10, pady=5)

        severe_button2 = tk.Radiobutton(popup2, text="Severe", variable=var2, value="Severe")
        severe_button2.pack(padx=10, pady=5)

        # Create the OK button and set focus to it
        ok_button2 = tk.Button(popup2, text="Save", padx=20, pady=10, command=popup2.destroy)
        ok_button2.pack(padx=10, pady=10)
        ok_button2.focus_set()

        # Bind the Enter key to the OK button
        popup2.bind("<Return>", lambda event: ok_button2.invoke())

        # Set focus to the default button
        if default_choice2 == "No deformity":
            no_def_button2.focus_set()
        elif default_choice2 == "Mild":
            mild_button2.focus_set()
        elif default_choice2 == "Moderate":
            moderate_button2.focus_set()
        elif default_choice2 == "Severe":
            severe_button2.focus_set()

        # Display the pop-up window and wait for it to be destroyed
        popup1.grab_set()
        popup1.wait_window()
        popup2.grab_set()
        popup2.wait_window()

        # Return the user's choice
        return var1.get(), var2.get()
    
    default_choice1 = "No deformity"
    default_choice2 = "No deformity"
    severity_L, severity_R = get_user_R_L()

    for row_index in filtered_rows:
        # global variables for this part
        Side = df.loc[row_index,'MotionParams_Side']
        EncounterAge = df.loc[row_index, 'EncounterAge']
        Trial = df.loc[row_index, 'GcdFile']
        pattern = df.loc[row_index, 'PatternOfInvolvement']
        GMFCS_level = df.loc[row_index, 'GmfcsNew']
        Date = df.loc[row_index, 'Date']
        MRN = int(df.loc[row_index, 'MRN'])
        e_date = str(Date).split()
        if len(e_date) > 1:
            t_date = None
            Date_E = str(e_date[0]) 
        else:
            Date_E = str(e_date[0])

        # variables for right side
        if str(Side).upper() == "R" or str(Side).lower() == "right":
            severity = severity_R
            AFO_tolerated = df.loc[row_index, 'RightOrthoticTolerated']
            supple_foot = df.loc[row_index, 'RightFootSupple']
            forefoot = df.loc[row_index, 'RightForefootWtBearing']
            midfoot = df.loc[row_index, 'RightMidfootWtBearing']
            hindfoot = df.loc[row_index, 'RightHindFootWtBearing']
            hallux = df.loc[row_index, 'RightHalluxWtBearing']
            forefootOff = df.loc[row_index, 'RightForefootOffWtBear']
            midfootOff = df.loc[row_index, 'RightMidfootOffWtBear']
            hindfootOff = df.loc[row_index, 'RightHindFootOffWtBear']
            phi_hind = df.loc[row_index, 'RightPhiHind']
            tone_gastroc1 = df.loc[row_index, 'RightToneGastroc']
            if isinstance(tone_gastroc1, (int, float)):
                tone_gastroc = tone_gastroc1
            else:
                tone_gastroc = -1
            peroneals = df.loc[row_index, 'RightTonePeroneals']
            per_clonus = df.loc[row_index, 'RightTonePeronealsClonus']

            ThighFoot = df.loc[row_index, 'RightThighFootAngle']
            if pd.isna(ThighFoot):
                number_thigh = -1
                ext_int_thigh = 'not'
            else:
                value_thigh_foot = str(ThighFoot).split()

                if len(value_thigh_foot) == 1:
                    match = re.search(r'(\d+)(\D+)?', value_thigh_foot[0])
                    if match:
                        try:
                            number_thigh = int(match.group(1))
                            ext_int_thigh = match.group(2) if match.group(2) else 'not'
                        except ValueError:
                            number_thigh = -1
                            ext_int_thigh = 'not'
                    else:
                        try:
                            number_thigh = int(value_thigh_foot[0])
                            ext_int_thigh = 'not'
                        except ValueError:
                            number_thigh = -1
                            ext_int_thigh = 'not'
                elif len(value_thigh_foot) == 2:
                    number_thigh = int(value_thigh_foot[0])
                    ext_int_thigh = value_thigh_foot[1]
                else:
                    number_thigh = -1
                    ext_int_thigh = 'not'

            text_side = 'Right side'
            
        # variables for left side
        elif str(Side).upper() == "L" or str(Side).lower() == "left":
            severity = severity_L
            AFO_tolerated = df.loc[row_index, 'LeftOrthoticTolerated']
            supple_foot = df.loc[row_index, 'LeftFootSupple']
            forefoot = df.loc[row_index, 'LeftForefootWtBearing']
            midfoot = df.loc[row_index, 'LeftMidfootWtBearing']
            hindfoot = df.loc[row_index, 'LeftHindFootWtBearing']
            hallux = df.loc[row_index, 'LeftHalluxWtBearing']
            forefootOff = df.loc[row_index, 'LeftForefootOffWtBear']
            midfootOff = df.loc[row_index, 'LeftMidfootOffWtBear']
            hindfootOff = df.loc[row_index, 'LeftHindFootOffWtBear']
            phi_hind = df.loc[row_index, 'LeftPhiHind']
            tone_gastroc1 = df.loc[row_index, 'LeftToneGastroc']
            if isinstance(tone_gastroc1, (int, float)):
                tone_gastroc = tone_gastroc1
            else:
                tone_gastroc = -1
            peroneals = df.loc[row_index, 'LeftTonePeroneals']
            per_clonus = df.loc[row_index, 'LeftTonePeronealsClonus']

            ThighFoot = df.loc[row_index, 'LeftThighFootAngle']
            if pd.isna(ThighFoot):
                number_thigh = -1
                ext_int_thigh = 'not'
            else:
                value_thigh_foot = str(ThighFoot).split()

                if len(value_thigh_foot) == 1:
                    match = re.search(r'(\d+)(\D+)?', value_thigh_foot[0])
                    if match:
                        try:
                            number_thigh = int(match.group(1))
                            ext_int_thigh = match.group(2) if match.group(2) else 'not'
                        except ValueError:
                            number_thigh = -1
                            ext_int_thigh = 'not'
                    else:
                        try:
                            number_thigh = int(value_thigh_foot[0])
                            ext_int_thigh = 'not'
                        except ValueError:
                            number_thigh = -1
                            ext_int_thigh = 'not'
                elif len(value_thigh_foot) == 2:
                    number_thigh = int(value_thigh_foot[0])
                    ext_int_thigh = value_thigh_foot[1]
                else:
                    number_thigh = -1
                    ext_int_thigh = 'not'
            text_side = 'Left side'
        
        # if there's no L or R it wont go to the decision tree
        else: 
            recommendation_str += "The file does not specify L or R side.\n"
            Pop_angle = None
            KneeFlexContr = None
            ThighFoot = None
            value_thigh_foot = None
            planovalgus = None
            ankle_equinus = None
            text_side = None
            continue
                    
        # add multiple keywords to variables that are needed in some conditions
        first_check_fore = ['abd', 'add', 'pronation', 'supination', 'valgus', 'varus', 'equinus', 'PF', 'PF 1st ray', 'PF 5th ray']
        first_check_mid = ['cavus','low med arch', 'high med arch', 'midfoot break', 'planovalgus', 'planus', 'valgus', 'pronation', 'rockr bttm', 'varus']
        first_check_hind = ['valgus', 'varus', 'calcaneus', 'equinus', 'mild varus', 'mild valgus']
        valgus_hind = ['valgus', 'mild valgus']
        valgus_fore = ['abd', 'mild abd', 'pronation', 'mild pronation', 'valgus', 'mild valgus']
        valgus_mid = ['low med arch', 'midfoot break', 'planovalgus', 'planus', 'mild planus', 'valgus', 'mild valgus', 'pronation', 'rockr bttm']
        varus_hind = ['varus', 'mild varus'] # first check +10 years
        varus_fore = ['add', 'mild add', 'mild supination', 'mild varus', 'supination', 'varus'] # first check +10 years
        varus_mid = ['cavus', 'varus', 'supination'] # first check +10 years
        mid_cavus = ['cavus', 'mild cavus']
        pos_hallux = ['PF 1st ray']
        pos_forefoot = ['abd', 'pronation', 'valgus']
        pos_midfoot = ['low med arch', 'midfoot break', 'planovalgus', 'planus', 'valgus', 'pronation', 'rockr bttm']
        pos_hindfoot = ['valgus']
        fore_hallux_pos = ['supination', 'PF 1st ray']
        # for +10 years
        valgus_hind2 = ['valgus', 'mild valgus', 'equinus', 'mild equinus'] 
        mild_fore = ['mild abd', 'mild equinus', 'mild pronation', 'mild valgus']
        mild_mid = ['mild planus', 'mild valgus']
        mild_hind = ['mild equinus', 'mild valgus']
        moderate_fore = ['abd', 'equinus', 'pronation', 'valgus']
        moderate_mid = ['equinus', 'low med arch', 'midfoot break', 'planovalgus', 'planus', 'valgus', 'pronation', 'rockr bttm']
        moderate_hind = ['valgus', 'equinus']
        severe_fore = ['abd', 'pronation', 'valgus', 'PF 1st ray']
        severe_hallux = ['HalVal']
        severe_mid = ['low med arch', 'midfoot break', 'planovalgus', 'planus', 'valgus', 'pronation', 'rockr bttm']
        severe_plano = ['supination', 'PF 1st ray']
        peroneals_clonus = ['+', '++', 'catch', '1 beat', '2 beats']
    

        # evaluate deformities 
        if (any(all(word in str(forefoot) for word in phrase.split()) for phrase in first_check_fore)) or (any(all(word in str(midfoot) for word in phrase.split()) \
            for phrase in first_check_mid)) or (any(all(word in str(hindfoot) for word in phrase.split()) for phrase in first_check_hind)):

            #First Age Group < 5
            if EncounterAge < 5:
                recommend1 = 'Use AFO'
                recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  ", Trial:" + str(Trial) + ", Recommendation: " \
                    + recommend1 +"\n"
                explanation_str += "The criteria for this recommendation is having varus or valgus foot deformities and being under 5 years old.\n In this trial "+\
                    "the patient's age is: " + str(EncounterAge) + " years old and matches the criteria\n\n"
                foot_recommend.append(recommend1)

            #Second Age Group 5 - 10
            if (EncounterAge >= 5) and (EncounterAge <= 10):
                # first decision: fixed deformity? (foot supple)
                if ('Y' in str(supple_foot).upper()) and ('Y' in str(AFO_tolerated).upper()):
                    recommend1 = 'Continue AFO as needed for function.'
                    recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  ", Trial:" + str(Trial) +\
                        ", Recommendation: " + recommend1 + "\n"
                    explanation_str += "The criteria for this recommendation is having a supple foot and AFO toleration\n\n"
                    foot_recommend.append(recommend1)

                elif (str(supple_foot).upper() == "N") or (str(supple_foot).upper() == "NO") or (str(AFO_tolerated).upper() == "N") or (str(AFO_tolerated).upper() == "NO"):
                    if (str(supple_foot).upper() == "N") or (str(supple_foot).upper() == "NO"):
                        if (str(AFO_tolerated).upper() == "N") or (str(AFO_tolerated).upper() == "NO"):
                            supple_tolerate = 'Fixed deformity and does not tolerate AFO'
                        else: 
                            supple_tolerate = 'Fixed deformity'
                    else:
                        supple_tolerate = 'No AFO toleration'

                    # check varus or valgus in hindfoot
                    if any(word in str(hindfoot) for word in varus_hind): #check for varus
                        # varus & hemiplegia 
                        if str(pattern).lower() == "hemiplegia":
                            #tibialis anterior branch 
                            def Ask1(Trial, text_side):
                                result = messagebox.askyesno("Varus and hemiplegia", "For trial " + str(Trial) + ", " + str(text_side) + ", Is the tibialis anterior "+\
                                    "muscle out of phase or constant on?")
                                return result
                            tibialis_anterior = Ask1(Trial, text_side)
                            if tibialis_anterior== True:
                                recommend1 = 'Do split transfer tibialis anterior.'
                                recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  ", Trial:" + str(Trial) +\
                                    ", Recommendation: " + recommend1 + "\n"
                                explanation_str += "The criteria for this recommedation is having:\n-Fixed deformity or not tolerating AFO\n-Varus\n-Hemiplegia"+\
                                    "\n-Tibialis anterior out of phase or constant on\n In this case the patient matches all the criteria\n\n"
                                foot_recommend.append(recommend1)
                            #tibialis posterior branch
                            else:
                                def Ask2(Trial, text_side):
                                    result = messagebox.askyesno("Varus and hemiplegia", "For trial " + str(Trial) + ", " + str(text_side) + ", Is the tibialis posterior "+\
                                        "out of phase or constant on?")
                                    return result
                                tibialis_posterior = Ask2(Trial, text_side)
                                if tibialis_posterior == True:
                                    recommend1 = 'Do split transfer tibialis posterior.'
                                    recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  ", Trial:" + str(Trial) +\
                                        ", Recommendation: " + recommend1 + "\n"
                                    explanation_str += "The criteria for this recommedation is having:\n-Fixed deformity or not tolerating AFO\n-Varus\n-Hemiplegia"+\
                                        "\n-Tibialis posterior out of phase or constant on\n In this case the patient matches all the criteria\n\n"
                                    foot_recommend.append(recommend1)
                                # no recommendation
                                else:
                                    recommend1 = 'No recommendation'
                                    recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  ", Trial:" + str(Trial) + \
                                        ", " + recommend1 + "\n"
                                    explanation_str += "This patient is between 5 and 10 years old, has a fixed deformity or does not tolerate AFO, has varus, has "+\
                                        "hemiplegia\nBut neither the tibials anterior or tibials posterior is out of phase or constant on.\nWith this criteria there is "+\
                                        "not an available recommendation using this Decision Tree\n\n"
                                    foot_recommend.append(recommend1)

                        # varus & diplegia
                        elif str(pattern).lower() == "diplegia":
                            # branch 1
                            if any(word in str(forefoot) for word in varus_fore):
                                def Ask1(Trial, text_side):
                                    result = messagebox.askyesno("Varus and diplegia", "For trial " + str(Trial) + ", " + str(text_side) + ", Is the tibialis anterior "+\
                                        "constant on?")
                                    return result
                                tibialis_anterior1 = Ask1(Trial)
                                if tibialis_anterior1 == True:
                                    recommend1 = 'Transfer the whole tibialis anterior to lateral cuneiform.'
                                    recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  ", Trial:" + str(Trial) +\
                                     ", Recommendation: " + recommend1 + "\n"
                                    explanation_str += "The criteria for this recommedation is having:\n-Fixed deformity or not tolerating AFO\n-Varus\n-Diplegia"+\
                                        "\n-Tibialis anterior constant on with forefoot varus\n In this case the patient matches all the criteria\n\n"
                                    foot_recommend.append(recommend1)
                                else: 
                                    recommend1 = 'No recommendation.'
                                    recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  ", Trial:" + str(Trial) + \
                                        ", " + recommend1 + "\n"
                                    explanation_str += "This patient has diplegia and has forefoot varus but the tibial anterior is not constant on\n\n"
                                    foot_recommend.append(recommend1)
                            # branch 2
                            elif any(word in str(midfoot) for word in mid_cavus) and any(word in str(hindfoot) for word in varus_hind):
                                def Ask2(Trial, text_side):
                                    result = messagebox.askyesno("Varus and diplegia", "For trial " + str(Trial) + ", " + str(text_side) + ", Is the tibialis posterior "+\
                                        "constant on?")
                                    return result
                                tibialis_posterior1 = Ask2(Trial, text_side)
                                if tibialis_posterior1 == True:
                                    recommend1 = 'Do a Z-lengthening tibialis posterior.'
                                    recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  ", Trial:" + str(Trial) + \
                                        ", Recommendation: " + recommend1 + "\n"
                                    explanation_str += "The criteria for this recommedation is having:\n-Fixed deformity or not tolerating AFO\n-Varus\n-Diplegia"+\
                                        "\n-Tibialis posterior constant on with hindfoot varus and cavus\n In this case the patient matches all the criteria\n\n"
                                    foot_recommend.append(recommend1)
                                else:
                                    recommend1 = 'No recommendation.'
                                    recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  ", Trial:" + str(Trial) + \
                                        ", " + recommend1 + "\n"
                                    explanation_str += "This patient has diplegia and has hindfoot varus and cavus but the tibial posterior is not constant on\n\n"
                                    foot_recommend.append(recommend1)
                            # branch 3
                            elif any(word in str(hindfoot) for word in varus_hind):
                                # branch 4
                                if str(supple_foot).upper() == "N" or str(supple_foot).upper() == "NO":
                                    recommend1 = 'Do a lateral column shortening (Evans procedure) and Z-lengthening tibialis posterior.'
                                    recommendation_str  += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  ", Trial:" + str(Trial) + \
                                        ", Recommendation: " + recommend1 + "\n"
                                    explanation_str += "The criteria for this recommedation is having:\n-Fixed deformity or not tolerating AFO\n- Severe fixed varus"+\
                                        "\n-Diplegia\n In this case the patient matches all the criteria\n\n"
                                    foot_recommend.append(recommend1)
                                elif str(supple_foot).upper() == "Y" or str(supple_foot).upper() == "YES":
                                    def Ask3(Trial, text_side):
                                        result = messagebox.askyesno("Varus and diplegia", "For trial " + str(Trial) + ", " + str(text_side) + ", Is the tibialis "+\
                                            "posterior out of phase?")
                                        return result
                                    tibialis_posterior2 = Ask3(Trial, text_side)
                                    if tibialis_posterior2 == True:
                                        recommend1 = 'Do a myofascial lengthening tibialis posterior.'
                                        recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  ", Trial:" + str(Trial) +\
                                             ",  Recommendation: " + recommend1 + "\n"
                                        explanation_str += "The criteria for this recommedation is having:\n-Fixed deformity or not tolerating AFO\n-Varus\n-Diplegia"+\
                                            "\n-Tibialis posterior out of phase\n In this case the patient matches all the criteria\n\n"
                                        foot_recommend.append(recommend1)
                                    else:
                                        recommend1 = 'No recommendation'
                                        recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  ", Trial:" + str(Trial) + \
                                            ", " + recommend1 + "\n"
                                        explanation_str += "This patient has diplegia and has hindfoot varus but the tibial posterior is not out of phase\n\n"
                                        foot_recommend.append(recommend1)
                                else:
                                    recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  ", Trial:" + str(Trial) + \
                                        ", " + recommend1 + "\n"
                                    explanation_str += "The patients has diplegia and hindfoot varus but it is not specified if the deformity is fixed or supple.\n\n"
                                    foot_recommend.append(recommend1)
                    # valgus 
                    elif any(word in str(hindfoot) for word in valgus_hind): #check for valgus

                        #independent ambulator
                        if GMFCS_level == 1 or GMFCS_level == 2:
                            
                            # toleration of orthotics
                            if 'Y' in str(AFO_tolerated).upper():
                                recommend1= 'Continue to use foot orthotics.'
                                recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  ", Trial:" + str(Trial) + \
                                    ",  Recommendation: " + recommend1 + "\n"
                                explanation_str += "The criteria for this recommendation is having a valgus deformity and AFO toleration\n\n"
                                foot_recommend.append(recommend1)
                            # no toleration of orthotics   
                            else:
                                if float(phi_hind) > 10:
                                    recommend1= 'Do a medial epiphysiodesis screw if growth remaining.'
                                    recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  ", Trial:" + str(Trial) +\
                                         ",  Recommendation: " + recommend1 + "\n"
                                    explanation_str += "The patient is between 5 a 10 years old, has varus, is an independent ambulator, does not tolerates orthotics "+\
                                        "and has ankle valgus greater than 10 degrees.\n\n"
                                    foot_recommend.append(recommend1)
                                else:
                                    if (tone_gastroc == 0) or (peroneals == 0): # no spasticity
                                        if severity == "Moderate":
                                            recommend1 = 'Do a subtalar fusion'
                                            if any(word in str(forefootOff) for word in fore_hallux_pos):
                                                recommend1= recommend1 + " AND transfer tibialis anterior to the midfoot and do a fusion or osteotomy naviculocuneiform joint"
                                                recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  \
                                                    ", Trial:" + str(Trial) + ",  Recommendation: " + recommend1 + "\n"
                                                explanation_str += "The patient has varus, is an independent ambulator, has no spasticity, has a moderate deformity and "+\
                                                    "a forefoot supination or dorsal bunion or 1st ray elevation" + "\n\n"
                                                foot_recommend.append(recommend1)
                                            else:
                                                recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  ", Trial:" + str(Trial) +\
                                                    ",  Recommendation: " + recommend1 + "\n"
                                                explanation_str += "The patient has varus, is an independent ambulator, has no spasticity and has a moderate deformity" + "\n\n"
                                                foot_recommend.append(recommend1)
                                        else:
                                            recommend1 = 'No recommendation'
                                            recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  ", Trial:" + str(Trial) +\
                                                ", " + recommend1 + "\n"
                                            explanation_str += "The patient has varus and is an independent ambulator, has no spasticity, but does not have a moderate "+\
                                                "deformity." + "\n\n"
                                            foot_recommend.append(recommend1)

                                    elif (tone_gastroc >= 2):
                                        #branch 1
                                        if severity == "Mild": # mild deformity 
                                            recommend1 = 'Do a lateral column lengthening' 
                                            if any(word in str(forefootOff) for word in fore_hallux_pos):
                                                recommend1= recommend1 + " AND transfer tibialis anterior to the midfoot and do a fusion or osteotomy naviculocuneiform joint"
                                                recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  \
                                                    ", Trial:" + str(Trial) + ",  Recommendation: " + recommend1 + "\n"
                                                explanation_str += "The patient has varus, is an independent ambulator, has spasticity, has a mild supple deformity and "+\
                                                    "a forefoot supination or dorsal bunion or 1st ray elevation" + "\n\n"
                                                foot_recommend.append(recommend1)
                                            else:
                                                recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  \
                                                    ", Trial:" + str(Trial) + ",  Recommendation: " + recommend1 + "\n"
                                                explanation_str += "The patient has varus, is an independent ambulator, has spasticity and  has a mild supple deformity\n\n"
                                                foot_recommend.append(recommend1)
                                        elif severity == "Moderate" or severity == "Severe": # moderate or severe deformity
                                            if any(word in str(per_clonus) for word in peroneals_clonus):
                                                recommend1 = 'Myofascial lengthening of perineus brevis ONLY with lateral column lengthening or subtalar fusion'
                                                if any(word in str(forefootOff) for word in fore_hallux_pos):
                                                    recommend1= recommend1 + " AND transfer tibialis anterior to the midfoot and do a fusion or osteotomy "+\
                                                        "naviculocuneiform joint"
                                                    recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  \
                                                        ", Trial:" + str(Trial) + ",  Recommendation: " + recommend1 + "\n"
                                                    explanation_str += "The patient has varus, is an independent ambulator, has severe or moderate spastic foot "+\
                                                        "deformity with contracted perineal tendons, and a forefoot supination or dorsal bunion or 1st ray elevation" + "\n\n"
                                                    foot_recommend.append(recommend1)
                                                else:
                                                    recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  \
                                                        ", Trial:" + str(Trial) + ",  Recommendation: " + recommend1 + "\n"
                                                    explanation_str += "The patient has varus, is an independent ambulator and has severe or moderate spastic foot "+\
                                                        "deformity with contracted perineal tendons\n\n" 
                                                    foot_recommend.append(recommend1)
                                            else: # no contracted perineal tendons
                                                recommend1= 'Do a subtalar fusion'
                                                if any(word in str(forefootOff) for word in fore_hallux_pos):
                                                    recommend1= recommend1 + " AND transfer tibialis anterior to the midfoot and do a fusion or osteotomy naviculocuneiform "+\
                                                        "joint"
                                                    recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) + \
                                                        ", Trial:" + str(Trial) + ",  Recommendation: " + recommend1 + "\n"
                                                    explanation_str += "The patient has varus, is an independent ambulator and has severe or moderate spastic foot "+\
                                                        "deformity and a forefoot supination or dorsal bunion or 1st ray elevation\n\n"
                                                    foot_recommend.append(recommend1)
                                                else:
                                                    recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  \
                                                        ", Trial:" + str(Trial) + ", Recommendation: " + recommend1 + "\n"
                                                    explanation_str += "The patient has varus, is an independent ambulator and has severe or moderate spastic foot deformity\n\n"
                                                    foot_recommend.append(recommend1)
                                        else:
                                            recommend1 = 'No recommendation'
                                            recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  \
                                                ", Trial:" + str(Trial) + ", " + recommend1 + "\n"
                                            explanation_str += "The patient has varus, is an independent ambulator but does not have a deformity\n\n"
                                            foot_recommend.append(recommend1)
                                    else:
                                        recommend1 = 'No recommendation'
                                        recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  ", Trial:" + str(Trial) + \
                                            ", " + recommend1 + "\n"
                                        explanation_str += "The patient is spastic but the value needed for a recommendation is having a muscle tone greater or equal "+\
                                            "to 2, and in this case the patient has a value of 1\n\n"
                                        foot_recommend.append(recommend1)

                        elif GMFCS_level == 3 or GMFCS_level == 4:
                            if any(word in str(hindfootOff) for word in pos_hindfoot):
                                recommend1 = 'Do a subtalar fusion'
                                if any(word in str(midfootOff) for word in pos_midfoot):
                                    recommend1= recommend1 + ' AND calcaneocuboid lengthening fusion'
                                    if (any(word in str(forefootOff) for word in pos_hallux)) or (any(word in str(forefootOff) for word in pos_forefoot)):
                                        recommend1 = recommend1 + ' AND transfer of tibialis anterior & fusion or naviculocuneiform osteotomy'
                                        recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  ", Trial:" + str(Trial) + \
                                            ",  Recommendation: " + recommend1 + "\n"
                                        explanation_str += "The patient has varus, is nonambulator or has limited ambulation with device, has hindfoot valgus, has "+\
                                            "planovalgus with severe forefoot abduction, forefoot supination, and 1st ray elevation\n\n"
                                        foot_recommend.append(recommend1)
                                    else:
                                        recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  ", Trial:" + str(Trial) + \
                                            ",  Recommendation: " + recommend1 + "\n"
                                        explanation_str += "The patient has varus, is nonambulator or has limited ambulation with device, has hindfoot valgus and has "+\
                                            "planovalgus with severe forefoot abduction\n\n"
                                        foot_recommend.append(recommend1)
                                else: 
                                    recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  ", Trial:" + str(Trial) + \
                                        ",  Recommendation: " + recommend1 + "\n"
                                    explanation_str += "The patient has varus, is nonambulator or has limited ambulation with device and has hindfoot valgus\n\n"
                                    foot_recommend.append(recommend1)
                            else: 
                                recommend1 = 'No recommendation'
                                recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  ", Trial:" + str(Trial) + \
                                    ", " + recommend1 + "\n"
                                explanation_str += "The patient has varus, is nonambulator or has limited ambulation with device, but does not have hindfoot valgus.\n\n"   
                                foot_recommend.append(recommend1)   
                        
                        elif math.isnan(GMFCS_level):
                            recommend1 = 'No recommendation'
                            recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  ", Trial:" + str(Trial) + \
                                ", " + recommend1 + "\n" 
                            explanation_str += "There is not a gross motor function classification system (GMFCS level) value.\n\n"
                            foot_recommend.append(recommend1)
                    
                    else:
                        recommend1 = 'No recommendation'
                        recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  ", Trial:" + str(Trial) + ", " + recommend1 + "\n"
                        explanation_str += "The patient is between 5 and 10 years old has a fixed deformity or does not tolerate AFO, but has no valgus or varus\n\n"
                        foot_recommend.append(recommend1)

                else:
                    recommend1 = 'No recommendation'
                    recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  ", Trial:" + str(Trial) + ", " + recommend1 + "\n"
                    explanation_str += "It is not specified whether or not the patient has a fixed deformity and/or tolerates orthotics\n\n"
                    foot_recommend.append(recommend1)
            
            #Third Age Group 
            if EncounterAge > 10:
                # check for varus
                if any(word in str(hindfoot) for word in varus_hind):
                    if (str(supple_foot).upper() == "Y") or (str(supple_foot).upper() == "YES"): # has supple deformity
                        def Ask1(Trial, text_side):
                            result = messagebox.askyesno("Varus and supple deformity", "For trial " + str(Trial) + ", " + str(text_side) + ", Is the tibialis anterior "+\
                                "muscle out of phase or constant on?")
                            return result
                        tibialis_anterior = Ask1(Trial, text_side)
                        # branch 1
                        if tibialis_anterior == True:
                            recommend1 = 'Do split transfer tibialis anterior'
                            recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  ", Trial:" + str(Trial) + \
                                ", Recommendation: " + recommend1 + "\n"
                            explanation_str += "The patient has varus, a supple deformity and has the tibialis anterior out of phase or constant on.\n\n"
                            foot_recommend.append(recommend1)
                            #tibialis posterior branch
                        else: 
                            def Ask2(Trial, text_side):
                                result = messagebox.askyesno("Varus and supple deformity", "For trial " + str(Trial) + ", " + str(text_side) + ", Is the tibialis "+\
                                    "posterior out of phase or constant on?")
                                return result
                            tibialis_posterior = Ask2(Trial, text_side)
                            if tibialis_posterior == True:
                                recommend1 = 'Do split transfer tibialis posterior.'
                                recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  ", Trial:" + str(Trial) + \
                                    ", Recommendation: " + recommend1 + "\n"
                                explanation_str += "The patient has varus, a supple deformity and has the tibialis posterior out of phase or constant on.\n\n"
                                foot_recommend.append(recommend1)
                            else:
                                def Ask3(Trial, text_side):
                                    result = messagebox.askyesno("Varus and supple deformity", "For trial " + str(Trial) + ", " + str(text_side) + ", Is there a "+\
                                        "fixed contracture of tibialis posterior?")
                                    return result
                                fixed_contracture = Ask3(Trial, text_side)
                            # branch 3
                                if fixed_contracture == True:
                                    recommend1 = 'Do a Z-lengthening tibialis posterior.'
                                    recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  ", Trial:" + str(Trial) + \
                                        ", Recommendation: " + recommend1 + "\n"
                                    explanation_str += "The patient has varus, a supple deformity and has a fixed contracture of tibialis posterior.\n\n"
                                    foot_recommend.append(recommend1)
                                else: 
                                    recommend1 = 'No recommendation'
                                    recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  ", Trial:" + str(Trial) + \
                                        ", " + recommend1 + "\n"
                                    explanation_str += "The patient has varus, a supple deformity, but does not have at least one of the following criteria:\n - "+\
                                        "The tibialis anterior out of phase or constant on.\n - The tibialis posterior out of phase or constant on.\n - A fixed contracture "+\
                                        "of tibialis posterior.\n\n"
                                    foot_recommend.append(recommend1)

                    elif (str(supple_foot).upper() == "N") or (str(supple_foot).upper() == "NO"): # has fixed deformity
                        if severity == "Mild":
                            recommend1 = 'No recommendation'
                            recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  ", Trial:" + str(Trial) + \
                                ", " + recommend1 + "\n"
                            explanation_str += "The patient is more than 10 years old, has a varus and has a mild fixed deformity\nFor fixed deformities with varus "+\
                                "there is just recommendation for moderate or severe deformities.\n\n"
                            foot_recommend.append(recommend1)
                        elif severity == "Moderate":
                            recommend1 = 'Do a calcaneal osteotomy.'
                            recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  ", Trial:" + str(Trial) + \
                                ", Recommendation: " + recommend1 + "\n"
                            explanation_str += "The patient is more than 10 years old, has a varus and has a moderate fixed deformity.\n\n"
                            foot_recommend.append(recommend1)
                        elif severity == "Severe":
                            recommend1 = 'Do a talectomy or triple arthrodesis.'
                            recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  ", Trial:" + str(Trial) + \
                                ", Recommendation: " + recommend1 + "\n"
                            explanation_str += "The patient is more than 10 years old, has a varus and has a severe fixed deformity.\n\n"
                            foot_recommend.append(recommend1)
                        else:
                            recommend1 = 'No recommendation'
                            recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  ", Trial:" + str(Trial) + \
                                ", " + recommend1 + "\n"
                            explanation_str += "The patient is more than 10 years old, has a varus but has no deformity.\n\n"
                            foot_recommend.append(recommend1)

                    elif math.isnan(supple_foot):
                        recommend1 = 'No recommendation'
                        recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  ", Trial:" + str(Trial) + ", " + recommend1 + "\n"
                        explanation_str += "The patient is more than 10 years old and has varus, but it's not specified if the deformity is fixed or supple.\n\n"
                        foot_recommend.append(recommend1)
                
                # check for valgus
                elif (any(word in str(hindfoot) for word in valgus_hind2)):
                    # ambulator without devices
                    if GMFCS_level == 1 or GMFCS_level == 2:
                        if severity == "Mild":
                            recommend1 = 'Do a lateral column lengthening through the calcaneus'
                            if any(word in str(hallux) for word in severe_hallux):
                                recommend1 = recommend1 + ' AND correct bunion'
                                if (number_thigh > 20) and ('E' in ext_int_thigh.upper()):
                                    # this condition is to get the external or internal text from the thigh angle cell
                                    if ('E' in str(ext_int_thigh).upper()):
                                        ext_int_thigh = 'external'
                                    elif ('I' in str(ext_int_thigh).upper()):
                                        ext_int_thigh = 'internal'
                                    elif ext_int_thigh == None:
                                        ext_int_thigh = 'unspecified'                                                 
                                    recommend1 = recommend1 + ' AND correct planovalgus and do a tibial derotation osteotomy'
                                    recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  ", Trial:" + str(Trial) + \
                                        ", Recommendation: " + recommend1 + "\n"
                                    explanation_str += "The patient has a mild supple deformity, planovalgus and has a external tibial torsion greater than 20 degrees "+\
                                        "external thigh to foot angle\n\n"
                                    foot_recommend.append(recommend1)
                                else:
                                    recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  ", Trial:" + str(Trial) + \
                                        ", Recommendation: " + recommend1 + "\n"
                                    explanation_str += "The patient has a mild supple deformity and planovalgus.\nTo get an additional recommendation it is necessary "+\
                                        "to have an external tibial torsion greater than 20 degrees\n - In this case the patient has an " + str(ext_int_thigh) + \
                                        " tibial torsion of " + str(number_thigh) + " degrees \n\n"
                                    foot_recommend.append(recommend1)
                            else:
                                if (number_thigh > 20) and ('E' in ext_int_thigh.upper()):
                                    # this condition is to get the external or internal text from the thigh angle cell
                                    if ('E' in str(ext_int_thigh).upper()):
                                        ext_int_thigh = 'external'
                                    elif ('I' in str(ext_int_thigh).upper()):
                                        ext_int_thigh = 'internal'
                                    elif ext_int_thigh == None:
                                        ext_int_thigh = 'unspecified'                                                 
                                    recommend1 = recommend1 + ' AND correct planovalgus and do a tibial derotation osteotomy'
                                    recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  ", Trial:" + str(Trial) + \
                                        ", Recommendation: " + recommend1 + "\n"
                                    explanation_str += "The patient has a mild supple deformity and has a external tibial torsion greater than 20 degrees external "+\
                                        "thigh to foot angle.\n - The exact value is an " + str(ext_int_thigh) + " tibial torsion of " + str(number_thigh) + " degrees \n\n"
                                    foot_recommend.append(recommend1)
                                else:
                                    # this condition is to get the external or internal text from the thigh angle cell
                                    if ('E' in str(ext_int_thigh).upper()):
                                        ext_int_thigh = 'external'
                                    elif ('I' in str(ext_int_thigh).upper()):
                                        ext_int_thigh = 'internal'
                                    elif ext_int_thigh == None:
                                        ext_int_thigh = 'unspecified'   
                                    recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  ", Trial:" + str(Trial) + \
                                        ", Recommendation: " + recommend1 + "\n"
                                    explanation_str += "The patient has a mild supple deformity \nTo get an additional recommendation it is necessary to have an "+\
                                        "external tibial torsion greater than 20 degrees\n - In this case the patient has an " + str(ext_int_thigh) + " tibial torsion "+\
                                        "of " + str(number_thigh) + " degrees \n\n"
                                    foot_recommend.append(recommend1)
                        elif severity == "Moderate":
                            recommend1= 'Do a subtalar fusion'
                            if any(word in str(hallux) for word in severe_hallux):
                                recommend1 = recommend1 + ' AND correct bunion'
                                if (number_thigh > 20) and ('E' in ext_int_thigh.upper()):
                                    # this condition is to get the external or internal text from the thigh angle cell
                                    if ('E' in str(ext_int_thigh).upper()):
                                        ext_int_thigh = 'external'
                                    elif ('I' in str(ext_int_thigh).upper()):
                                        ext_int_thigh = 'internal'
                                    elif ext_int_thigh == None:
                                        ext_int_thigh = 'unspecified'                                          
                                    recommend1 = recommend1 + ' AND correct planovalgus and do a tibial derotation osteotomy'
                                    recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  ", Trial:" + str(Trial) + \
                                        ", Recommendation: " + recommend1 + "\n"
                                    explanation_str += "The patient has a moderate deformity, planovalgus and has a external tibial torsion greater than 20 degrees "+\
                                        "external thigh to foot angle\n\n"
                                    foot_recommend.append(recommend1)
                                else:
                                    # this condition is to get the external or internal text from the thigh angle cell
                                    if ('E' in str(ext_int_thigh).upper()):
                                        ext_int_thigh = 'external'
                                    elif ('I' in str(ext_int_thigh).upper()):
                                        ext_int_thigh = 'internal'
                                    elif ext_int_thigh == None:
                                        ext_int_thigh = 'unspecified'  
                                    recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  ", Trial:" + str(Trial) + \
                                        ", Recommendation: " + recommend1 + "\n"
                                    explanation_str += "The patient has a moderate deformity and planovalgus.\nTo get an additional recommendation it is necessary to "+\
                                        "have an external tibial torsion greater than 20 degrees\n - In this case the patient has an " + str(ext_int_thigh) + " tibial "+\
                                            "torsion of " + str(number_thigh) + " degrees \n\n"  
                                    foot_recommend.append(recommend1)
                            else:
                                if (number_thigh > 20) and ('E' in ext_int_thigh.upper()):
                                    # this condition is to get the external or internal text from the thigh angle cell
                                    if ('E' in str(ext_int_thigh).upper()):
                                        ext_int_thigh = 'external'
                                    elif ('I' in str(ext_int_thigh).upper()):
                                        ext_int_thigh = 'internal'
                                    elif ext_int_thigh == None:
                                        ext_int_thigh = 'unspecified'  
                                    recommend1 = recommend1 + ' AND correct planovalgus and do a tibial derotation osteotomy'
                                    recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  ", Trial:" + str(Trial) + \
                                        ", Recommendation: " + recommend1 + "\n"
                                    explanation_str += "The patient has a moderate deformity and has a external tibial torsion greater than 20 degrees external thigh "+\
                                        "to foot angle.\n - The exact value is an " + str(ext_int_thigh) + " tibial torsion of " + str(number_thigh) + " degrees \n\n"
                                    foot_recommend.append(recommend1)
                                else:
                                    # this condition is to get the external or internal text from the thigh angle cell
                                    if ('E' in str(ext_int_thigh).upper()):
                                        ext_int_thigh = 'external'
                                    elif ('I' in str(ext_int_thigh).upper()):
                                        ext_int_thigh = 'internal'
                                    elif ext_int_thigh == None:
                                        ext_int_thigh = 'unspecified'  
                                    recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  ", Trial:" + str(Trial) + \
                                        ", Recommendation: " + recommend1 + "\n"
                                    explanation_str += "The patient has a moderate deformity \nTo get an additional recommendation it is necessary to have an external "+\
                                        "tibial torsion greater than 20 degrees\n - In this case the patient has an " + str(ext_int_thigh) + " tibial torsion "+\
                                            "of " + str(number_thigh) + " degrees \n\n"
                                    foot_recommend.append(recommend1)

                        # severe
                        elif severity == "Severe":
                            if phi_hind >= 10:
                                def Ask1(Trial):
                                    result = messagebox.askyesno("Valgus and independent ambulator", "For trial: " + str(Trial) + ", " + str(text_side) + ", Is the "+\
                                        "growth plate open?")
                                    return result
                                open_growth = Ask1(Trial)
                                if open_growth == True:
                                    recommend1='Do an epiphysiodesis screw'
                                    text_open = 'has open growth plate'
                                else:
                                    recommend1='Do a tibia osteotomy'
                                    text_open = 'has closed growth plate'
                                
                                if any(word in str(hallux) for word in severe_hallux):
                                    recommend1 = recommend1 + ' AND correct bunion'
                                    if (number_thigh > 20) and ('E' in ext_int_thigh.upper()):
                                        # this condition is to get the external or internal text from the thigh angle cell
                                        if ('E' in str(ext_int_thigh).upper()):
                                            ext_int_thigh = 'external'
                                        elif ('I' in str(ext_int_thigh).upper()):
                                            ext_int_thigh = 'internal'
                                        elif ext_int_thigh == None:
                                            ext_int_thigh = 'unspecified'  
                                        recommend1 = recommend1 + ' AND correct planovalgus and do a tibial derotation osteotomy'
                                        recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  ", Trial:" + str(Trial) + \
                                            ", Recommendation: " + recommend1 + "\n"
                                        explanation_str += "The patient has planovalgus with greater than 10 degrees ankle valgus and " + text_open +".\nIn addition, the "+\
                                            "patient has an external tibial torsion greater than 20 degrees external thigh to foot angle.\n - The exact value is "+\
                                                "an " + str(ext_int_thigh) + " tibial torsion of " + str(number_thigh) + " degrees \n\n"
                                        foot_recommend.append(recommend1)
                                    else:
                                        recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  ", Trial:" + str(Trial) + \
                                            ", Recommendation: " + recommend1 + "\n"
                                        explanation_str += "The patient has planovalgus with greater than 10 degrees ankle valgus and " + text_open + "\nTo get an "+\
                                            "additional recommendation it is necessary to have an external tibial torsion greater than 20 degrees\n - In this case the "+\
                                            "patient has an " + str(ext_int_thigh) + " tibial torsion of " + str(number_thigh) + " degrees\n\n"
                                        foot_recommend.append(recommend1)
                                else:
                                    if (number_thigh > 20) and ('E' in ext_int_thigh.upper()):
                                        # this condition is to get the external or internal text from the thigh angle cell
                                        if ('E' in str(ext_int_thigh).upper()):
                                            ext_int_thigh = 'external'
                                        elif ('I' in str(ext_int_thigh).upper()):
                                            ext_int_thigh = 'internal'
                                        elif ext_int_thigh == None:
                                            ext_int_thigh = 'unspecified'  
                                        recommend1 = recommend1 + ' AND correct planovalgus and do a tibial derotation osteotomy'
                                        recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  ", Trial:" + str(Trial) + \
                                            ", Recommendation: " + recommend1 + "\n"
                                        explanation_str += "The patient has more than 10 degrees ankle valgus and " + text_open +".\nIn addition, the patient has an "+\
                                            "external tibial torsion greater than 20 degrees external thigh to foot angle.\n - The exact value is an " + str(ext_int_thigh) + \
                                            " tibial torsion of " + str(number_thigh) + " degrees \n\n"
                                        foot_recommend.append(recommend1)
                                    else:
                                        # this condition is to get the external or internal text from the thigh angle cell
                                        if ('E' in str(ext_int_thigh).upper()):
                                            ext_int_thigh = 'external'
                                        elif ('I' in str(ext_int_thigh).upper()):
                                            ext_int_thigh = 'internal'
                                        elif ext_int_thigh == None:
                                            ext_int_thigh = 'unspecified'  
                                        recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  ", Trial:" + str(Trial) + \
                                            ", Recommendation: " + recommend1 + "\n"
                                        explanation_str += "The patient has more than 10 degrees ankle valgus and " + text_open +".\nTo get an additional recommendation "+\
                                            "it is necessary to have an external tibial torsion greater than 20 degrees or to have planovalgus\n - In this case the "+\
                                            "patient has an " + str(ext_int_thigh) + " tibial torsion of " + str(number_thigh) + " degrees and has no planovalgus\n\n"
                                        foot_recommend.append(recommend1)
                            else:
                                recommend1 = 'No recommendation'
                                recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  ", Trial:" + str(Trial) + \
                                    ", " + recommend1 + "\n"
                                explanation_str += "For patients older than 10 years old and GMFCS level I or II, has a severe deformity, but has no more than 10 "+\
                                    "degrees ankle valgus.\n\n"
                                foot_recommend.append(recommend1)
                            
                        else:
                            recommend1 = 'No recommendation'
                            recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  ", Trial:" + str(Trial) + \
                                ", " + recommend1 + "\n"
                            explanation_str += "For patients older than 10 years old and GMFCS level I or II, there is just recommendations for mild, moderate or "+\
                                "severe deformities. This patients has none\n\n"
                            foot_recommend.append(recommend1)

                    # nonambulator or with limited ambulation with device
                    elif GMFCS_level == 3 or GMFCS_level == 4:
                        if 'valgus' in str(hindfoot):
                            recommend1 = 'Do subtalar fussions'
                            if any(word in str(midfoot) for word in severe_mid):
                                recommend1 = recommend1 + ' AND calcaneocuboid lengthening fusion'
                                if any(word in str(forefoot) for word in severe_fore):
                                    recommend1= recommend1 + ' AND a fusion of naviculocuneiform joint or fuse whole medial column'
                                    if any(word in str(hallux) for word in severe_hallux):
                                        recommend1 = recommend1 + ' AND add a fusion of the 1st MTP joint'
                                        recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  ", Trial:" + str(Trial) + \
                                            ", Recommendation: "+ recommend1 + "\n"
                                        explanation_str += "The patient is considered a nonambulator or ambulated with devices, has planovalgus feet with hindfoot "+\
                                            "valgus, forefoot abduction and forefoot supination & 1st ray elevation & dorsal bunion, with symptomatic bunion\n\n"
                                        foot_recommend.append(recommend1)
                                    else:
                                        recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  ", Trial:" + str(Trial) +\
                                             ", Recommendation: "+ recommend1 + "\n"
                                        explanation_str += "The patient is considered a nonambulator or ambulated with devices, has planovalgus feet with hindfoot "+\
                                            "valgus, forefoot abduction and forefoot supination & 1st ray elevation & dorsal bunion\n\n"
                                        foot_recommend.append(recommend1)
                                else:
                                    if any(word in str(hallux) for word in severe_hallux):
                                        recommend1 = recommend1 + ' AND add a fusion of the 1st MTP joint'
                                        recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  ", Trial:" + str(Trial) + \
                                            ", Recommendation: "+ recommend1 + "\n"
                                        explanation_str += "The patient is considered a nonambulator or ambulated with devices, has planovalgus feet with hindfoot valgus "+\
                                            "and forefoot abduction, with symptomatic bunion\n\n"
                                        foot_recommend.append(recommend1)
                                    else:
                                        recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  ", Trial:" + str(Trial) + \
                                            ", Recommendation: "+ recommend1 + "\n"
                                        explanation_str += "The patient is considered a nonambulator or ambulated with devices, has planovalgus feet with hindfoot "+\
                                            "valgus and forefoot abduction\n\n"
                                        foot_recommend.append(recommend1)
                            else:
                                if any(word in str(hallux) for word in severe_hallux):
                                    recommend1 = recommend1 + ' AND add a fusion of the 1st MTP joint'
                                    recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  ", Trial:" + str(Trial) + \
                                        ", Recommendation: "+ recommend1 + "\n"
                                    explanation_str += "The patient is considered a nonambulator or ambulated with devices, has planovalgus feet with hindfoot valgus\n\n"
                                    foot_recommend.append(recommend1)
                                else:
                                    recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  ", Trial:" + str(Trial) + \
                                        ", Recommendation: "+ recommend1 + "\n"
                                    explanation_str += "The patient is considered a nonambulator or ambulated with devices, has planovalgus feet with hindfoot valgus, "+\
                                        "and forefoot supination & 1st ray elevation & dorsal bunion\n\n"
                                    foot_recommend.append(recommend1)
                        else:
                            recommend1 = 'No recommendation'
                            recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  ", Trial:" + str(Trial) + \
                                ", " + recommend1 + "\n"
                            explanation_str += "For patients more than 10 years old, with a GMFCS level III or IV, there is just a recommendation if the patient "+\
                                "has any type of planovalgus feet deformity, specifically hindfoot valgus\n\n"
                            foot_recommend.append(recommend1)
                    else: 
                        recommend1 = 'No recommendation'
                        recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  ", Trial:" + str(Trial) + ", " + recommend1 + "\n"
                        explanation_str += "For patients more than 10 years old, that have a valgus deformity there is only recommendations if the patient has GMFCS "+\
                            "levels 1, 2, 3 or 4\n\n"
                        foot_recommend.append(recommend1)

                # no valgus or varus
                else: 
                    recommend1 = 'No recommendation'
                    recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  ", Trial:" + str(Trial) + ", " + recommend1 + "\n"
                    explanation_str += "No words from the groups valgus presented\n\n"
                    foot_recommend.append(recommend1)                            

        else:
            recommend1 = 'No recommendation'
            recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E) +  ", Trial:" + str(Trial) + ", " + recommend1 + "\n"
            explanation_str += " Foot valgus or varus not present in this patient.\n\n"
            foot_recommend.append(recommend1)   

    # this part is to add the recommendations to the text widget called recommendation_text
    recommendation_text.delete('1.0', END)
    # add format to the recommendation text displayed
    recommendation_text.insert(END, recommendation_str, ('tag1', 'color_blue', 'space1'))
    recommendation_text.tag_configure('tag1', font=('Open Sans', 12))
    recommendation_text.tag_configure('color_blue', foreground='black')
    recommendation_text.tag_configure('space1', spacing3=15)

    def export_Excel(foot_recommend, df):
        df['Freeman Miller Foot Recommendations'] = foot_recommend

        export_file_path = filedialog.asksaveasfilename(defaultextension='.xlsx')
        df.to_excel(export_file_path, index=False)

    # function to export the recommendations and criteria report to a pdf file
    def export_pdf_full(recommendation_str, explanation_str, MRN, selected_dates2):
        messagebox.showinfo("Save report", "Select the folder to export the report\n\nThe report is in pdf format")
        folder_pathpdf = filedialog.askdirectory()
        if not folder_pathpdf:
            return
        dates = [str(date.year) for date in selected_dates2]

        # join dates with underscore
        joined_dates = '_'.join(dates)

        filename=f"{folder_pathpdf}/" + str(MRN) + "_FootRec_"+ joined_dates +"_full_report.pdf"

        c=canvas.Canvas(filename, pagesize=letter)
        max_width = 95
        page_height = c._pagesize[1] - 5

        recommendation_lines = recommendation_str.split("\n")
        explanation_lines = explanation_str.split("\n\n")

        y = page_height - 30
        add_extra_text = True
        # add 1 recommendation at a time
        for i in range(len(recommendation_lines)):
            if add_extra_text and i == 0:
                y -= 5 # space to add the header
                c.setFont("Helvetica-Bold", 16)
                c.drawString(135, y, "Recommendations from the Freeman Miller")
                c.drawString(180, y-20, "Foot Treatment Decision Tree")
                y -= 60
                add_extra_text = False

            c.setFont("Helvetica", 12)
            c.setFillColorRGB(0, 0, 1) # blue color
            wrapped_recommendation = textwrap.wrap(recommendation_lines[i], width = max_width)
           
            num_lines_rec = len(wrapped_recommendation)
            if y - num_lines_rec*20 < 35:
                
                c.showPage()
                y = page_height - 30
                ##
            c.setFont("Helvetica", 12)
            c.setFillColorRGB(0, 0, 1) # blue color
            for line in wrapped_recommendation:
                c.drawString(50, y, line)
                y -= 20
            
            # check the space left in the page to avoid cropped information
            exp_lines = explanation_lines[i].split('\n')
            num_lines = sum(len(textwrap.wrap(line, width=max_width+5)) for line in exp_lines)
            if y - num_lines*20 < 35:
                c.showPage()
                y = page_height - 30
            
            # show every explanation for each recommendation
            c.setFont("Helvetica", 10)
            for line in exp_lines:
                wrapped_explanation = textwrap.wrap(line, width=(max_width+15))
                for exp_line in wrapped_explanation:
                    c.setFillColorRGB(0, 0, 0) # black color
                    c.drawString(70, y, exp_line.strip())
                    y -= 20
            y -= 20
        c.save()
        messagebox.showinfo("Export", "Report exported")

    # function to export JUST the recommendation report to a pdf file
    def export_pdf(recommendation_str, MRN, selected_dates2):
        messagebox.showinfo("Save report", "Select the folder to export the report\n\nThe report is in pdf format")
        folder_pathpdf = filedialog.askdirectory()
        if not folder_pathpdf:
            return
        
        dates = [str(date.year) for date in selected_dates2]

        # join dates with underscore
        joined_dates = '_'.join(dates)

        filename=f"{folder_pathpdf}/" + str(MRN) + "_FootRec_"+ joined_dates +"_report.pdf"

        c=canvas.Canvas(filename, pagesize=letter)
        max_width = 115 # this control the horizontal size in which the text can be displayed
        page_height = c._pagesize[1] - 5
        recommendation_lines = recommendation_str.split("\n")

        y = page_height - 30
        add_extra_text = True # verify that there is enough space left in the same page for the full recommendation or explanation
        for i in range(len(recommendation_lines)):
            if add_extra_text and i == 0:
                y -= 5 # space to add the header
                c.setFont("Helvetica-Bold", 16)
                c.drawString(135, y, "Recommendations from the Freeman Miller")
                c.drawString(180, y-20, "Foot Treatment Decision Tree")
                y -= 60
                add_extra_text = False

            c.setFont("Helvetica", 10)
            c.setFillColorRGB(0, 0, 0) # black color
            wrapped_recommendation = textwrap.wrap(recommendation_lines[i], width = max_width)
            for line in wrapped_recommendation:
                c.drawString(50, y, line)
                y -= 12
            
            # check the space left in the page to avoid cropped information
            if y < 40:
                c.showPage()
                y = page_height - 30
            y -= 15
            
        c.save()
        messagebox.showinfo("Export", "Report exported")

    # button to export the recommendations and criteria in a pdf file
    button_export1=Button(root, text="Export\nreport\n(with criteria)", font=("Open Sans", 10), bg="#BFBFBF", padx=10, pady=8, \
        command= lambda: export_pdf_full(recommendation_str, explanation_str, MRN, selected_dates2))
    button_export1.place(x=1230, y=190)

    # button to export Excel file
    button_export2=Button(root, text="Export\nExcel\nfile", font=("Open Sans", 10), bg="#BFBFBF", padx=15, pady=8, command= lambda: export_Excel(foot_recommend, df))
    button_export2.place(x=1100, y=190)

    # # button to export the recommendations in a pdf file
    # button_export2=Button(root, text="Export\nreport\n(no criteria)", font=("Open Sans", 10), bg="#BFBFBF", padx=10, pady=8, \
    #   command= lambda: export_pdf(recommendation_str, MRN, selected_dates2))
    # button_export2.place(x=1100, y=190)

    # button for the pop up window with the criteria
    button_criteria=Button(root, text="Recommendation\nCriteria", font=("Open Sans", 12), bg="#D9E1F2", padx=10, pady=8, \
        command=lambda: show_popup(recommendation_str, explanation_str))
    button_criteria.place(x=850, y=195)

    # pop up window that shows the criteria used for each recommendation
    def show_popup(recommendation_str, explanation_str):
        # create a child window
        popup_window = Toplevel(root)
        popup_window.geometry("950x575+700+350")
        popup_window.title("Recommendation Criteria")

        popup_window.resizable(False,False)
        # create a Frame to hold the Text widgets and scrollbar
        frame = Frame(popup_window)
        frame.place(x=10, y=10)

        # create a scrollbar for the Text widgets
        vscrollbar = Scrollbar(frame)
        vscrollbar.pack(side=RIGHT, fill=Y)

        # create a horizontal scrollbar for the Text widget
        hscrollbar = Scrollbar(frame, orient=HORIZONTAL)
        hscrollbar.pack(side=BOTTOM, fill=X)

        # create a Text widget to hold the recommendation text
        recommendation_text = Text(frame, height=25, width=115, wrap=NONE, yscrollcommand=vscrollbar.set, xscrollcommand=hscrollbar.set, spacing3=5)
        recommendation_text.pack(side=LEFT, fill=Y)

        # cet the scrollbar to control the Text widget
        vscrollbar.config(command=recommendation_text.yview)

        # cet the horizontal scrollbar to control the Text widget
        hscrollbar.config(command=recommendation_text.xview)

        # add the recommendation and explanation text to the new window
        recommendation_text.delete('1.0', END)
        lines = recommendation_str.split("\n")

        #use a loop to add 1 recommendation and 1 explanation at a time
        for i in range(len(lines)):
            recommendation_text.insert(END, lines[i] + "\n", ('tag1', 'color_blue'))
            recommendation_text.tag_configure('tag1', font=('Open Sans', 12))
            recommendation_text.tag_configure('color_blue', foreground='blue')
            if i<len(explanation_str.split("\n\n")):
                recommendation_text.insert(END, explanation_str.split("\n\n")[i] + "\n\n", ('tag2', 'color_black'))
                recommendation_text.tag_configure('tag2', font=('Open Sans', 10))
                recommendation_text.tag_configure('color_black', foreground='black')

def GaitDev():
    global error_label, ins_label, button_export1, button_export2, button_criteria, sagittal_col, trans_col, instructions_label
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

    #start of the function
    recommendation_str = ""
    explanation_str = ""

    fold_name = folder_name.get()
    if fold_name == "":
        messagebox.showerror("Missing data", "There was not MRN typed or any browsed file\nPlease try again")
        return
    elif fold_name.isdigit():
        file_name = str(fold_name) + "_DT_query.xlsx"
        folder_path = os.path.join(folder_location, str(fold_name))
        file_path = os.path.join(folder_path, file_name)
        error_label = Label(root, text=" ", fg="red")
        error_label.place(x=75, y=320)
    else:
        file_path= fold_name
        error_label = Label(root, text=" ", fg="red")
        error_label.place(x=75, y=320)
    
    df = pd.read_excel(file_path) # read the excel file

    if 'complete_query' in fold_name:
        num_rows = len(df.index)
        filtered_rows = range(num_rows)
    else:
        # Select dates
        selected_dates2 = select_dates(df)
        df['Date'] = pd.to_datetime(df['Date']).dt.date
        filtered_rows = df[df['Date'].isin(selected_dates2)].index

    title_label.config(text="Gait Deviation Classification", fg="black", font=('Open Sans - Bold', 16))

    ######
    # stablish cutoff values (they can be an input from the user)

    for row_index in filtered_rows:
        # global variables for this part
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
        Trial = df.loc[row_index, 'GcdFile'] # trial number that is in the GcdFile column
        Date = df.loc[row_index, 'Date']
        MRN = int(df.loc[row_index, 'MRN'])
        e_date = str(Date).split()
        if len(e_date) > 1:
            t_date = None
            Date_E = str(e_date[0]) 
        else:
            Date_E = str(e_date[0]) 

        # variables for right side
        if 'R' in str(Side).upper():
            text_side = 'Right side'

        # variables for left side
        elif 'L' in str(Side).upper():
            text_side = 'Left side'
        
        # if there's no L or R it wont go to the decision tree
        else: 
            text_side = None
            continue

        # sagittal PLANE 
        if KneeFlexMin <= 0:
            if AnkleDorsi <= 4.4:
                if HipFlexMin <= 6.9:
                    sagittal = 'Knee Recurvatum, Ankle Plantar Flexed and Hip Not Flexed'
                else: #HipFlexMin > 6.9:
                    sagittal = 'Knee Recurvatum, Ankle Plantar Flexed and Hip Flexed'
            elif (AnkleDorsi <= 20.3) and (AnkleDorsi > 4.4):
                if HipFlexMin <= 6.9:
                    sagittal = 'Knee Recurvatum, Ankle Normal Flexed and Hip Not Flexed'
                else: #HipFlexMin > 6.9:
                    sagittal = 'Knee Recurvatum, Ankle Normal Flexed and Hip Flexed'
            elif AnkleDorsi > 20.3:
                if HipFlexMin <= 6.9:
                    sagittal = 'Knee Recurvatum, Ankle Dorsi Flexed and Hip Not Flexed'
                else: #HipFlexMin > 6.9:
                    sagittal = 'Knee Recurvatum, Ankle Dorsi Flexed and Hip Flexed'
            else:
                sagittal = 'No Classification' # no classification
        
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
                    elif (KneeFlexMean <= 23.1) and (HipFlexMin<= 6.9):
                        sagittal = 'Apparent Ankle Equinus & No Excess Hip or Knee Flexion'
                    else:
                        sagittal = 'No Classification' #no classification
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
                        sagittal = 'No Classification' #no classification
                else:
                    if (KneeRange <= 43.3) and (KneeFlexMean <= 23.1) and (HipFlexMin <= 6.9):
                        sagittal = 'Exclusive Stiff Knee'
                    else:
                        sagittal = 'No Classification'
            else:
                sagittal = 'No Classification'
        else:
            sagittal = 'No Classification'

        #TRANSVERSAL PLANE
        #internal
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
            
        #external
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

        #neutral 
        elif (FootProgress > -21.5) and (FootProgress <= 6.3):
            if ((KneeRot <= 17.4) and (KneeRot > -15.8)) and ((HipRot <= 19) and (HipRot > -17.4)) and ((PelvicRot <= 7.8) and (PelvicRot > -7.3)):
                transversal = 'True Neutral'
            else:
                if KneeRot > 17.4:
                    if HipRot > 19:
                        if PelvicRot>7.8:
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
                elif HipRot> 19:
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
        
        recommendation_str += "Patient " + str(MRN) + ", " + text_side + ", Encounter Date: " + str(Date_E)  + "\n  Sagittal: " + str(sagittal) + \
            "\n  Transversal: " + str(transversal) + "\n\n"
        sagittal_col.append(sagittal)
        trans_col.append(transversal)

    lines = recommendation_str.split("\n\n")
    recommendation_text.delete('1.0', END)
    #use a loop to add 1 recommendation and 1 explanation at a time
    for i in range(len(lines)):
        recommendation_text.insert(END, lines[i] + "\n\n", ('tag1', 'color_black'))
        recommendation_text.tag_configure('tag1', font=('Open Sans', 12))
        recommendation_text.tag_configure('color_black', foreground='black')
    
    def export_Excel(sagittal_col, trans_col, df):
        df['Gait Deviation Sagittal Plane'] = sagittal_col
        df['Gait Deviation Transversal Plane'] = trans_col
        export_file_path = filedialog.asksaveasfilename(defaultextension='.xlsx')
        df.to_excel(export_file_path, index=False)

    def export_pdf(recommendation_str, MRN, selected_dates2):
        messagebox.showinfo("Save report", "Select the folder to export the report\n\nThe report is in pdf format")
        folder_pathpdf = filedialog.askdirectory()
        if not folder_pathpdf:
            return
        dates = [str(date.year) for date in selected_dates2]

        # join dates with underscore
        joined_dates = '_'.join(dates)

        filename=f"{folder_pathpdf}/" + str(MRN) + "_GaitClassif_"+ joined_dates +"_report.pdf"

        c=canvas.Canvas(filename, pagesize=letter)
        max_width = 115 # this control the horizontal size in which the text can be displayed
        page_height = c._pagesize[1] - 5
        recommendation_lines = recommendation_str.split("\n\n")

        y = page_height - 30
        add_extra_text = True # verify that there is enough space left in the same page for the full recommendation or explanation
        for i in range(len(recommendation_lines)):
            if add_extra_text and i == 0:
                y -= 5 # space to add the header
                c.setFont("Helvetica-Bold", 16)
                c.drawString(230, y, "Gait Classifications")
                #c.drawString(180, y-20, "Knee Treatment Decision Tree")
                y -= 40
                add_extra_text = False

            c.setFont("Helvetica", 10)
            c.setFillColorRGB(0, 0, 0) # black color
            wrapped_recommendation = textwrap.wrap(recommendation_lines[i], width = max_width)
            for line in wrapped_recommendation:
                c.drawString(50, y, line)
                y -= 12
        
            # check the space left in the page to avoid cropped information
            if y < 40:
                c.showPage()
                y = page_height - 30
            y -= 15
        c.save()
        messagebox.showinfo("Export", "Report exported")
    
    # button to export Excel file
    button_export2=Button(root, text="Export\nExcel\nfile", font=("Open Sans", 10), bg="#BFBFBF", padx=20, pady=8, \
        command= lambda: export_Excel(sagittal_col, trans_col, df))
    button_export2.place(x=1100, y=190)

    # button to export the recommendations in a pdf file
    button_export1=Button(root, text="Export report", font=("Open Sans", 12), bg="#BFBFBF", padx=25, pady=15, \
        command= lambda: export_pdf(recommendation_str, MRN, selected_dates2))
    button_export1.place(x=1230, y=190)

def save_Excel(knee_recommend, foot_recommend, sagittal_col, trans_col, selected_dates2):
    fold_name = folder_name.get()
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
                
                    initial_dir = folder_path
                    export_file_path = filedialog.asksaveasfilename(initialdir=initial_dir, defaultextension='.xlsx')

                    df.to_excel(export_file_path, index=False)
                    messagebox.showinfo("Export", "File exported")
    else:
        messagebox.showwarning("Export Error", "There is no data to export. \n First select the patient's data and try again.")

# function to open the pdf reference of the knee
def openRef_knee():
    file_p = os.path.join(script_dir, 'resources', '2005_Book_CerebralPalsy_Freeman Miller(Knee).pdf')
    os.startfile(file_p)

# function to open the pdf reference of the foot
def openRef_foot():
    file_p = os.path.join(script_dir, 'resources', '2005_Book_CerebralPalsy_Freeman Miller(Foot Decision Tree).pdf')
    os.startfile(file_p)

#function to delete the previous data and look for a new patient
def new_patient():
    global error_label, ins_label, button_export1, button_export2, button_criteria, diagnosis, diagnosis_label, knee_recommend, foot_recommend, sagittal_col, trans_col
    recommendation_text.delete('1.0', END)
    folder_name.delete(0, tk.END)
    first_name.delete(0, tk.END)
    last_name.delete(0, tk.END)
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

def all_Data():
    global error_label, ins_label, diagnosis, diagnosis_label
    folder_name.delete(0, tk.END)
    # Destroy previous labels if they exist
    if error_label:
        error_label.destroy()
    if ins_label:
        ins_label.destroy()
    if diagnosis:
        diagnosis.destroy()
    if diagnosis_label:
        diagnosis_label.destroy()

    conn = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb)};DBQ=G:\Gait Lab\_MAL Database\PatientsAndServicesDatabase.mdb;')
    cursor = conn.cursor()

    primary_diagnosis = "cerebral palsy"

    sql_query="""
    SELECT Patients.MRN, Patients.LastName, Patients.FirstName, Patients.PrimaryDiagnosis, 
        Patients.PatternOfInvolvement, Patients.Side AS Patients_Side, Patients.DOB, Encounters.Date, 
        DateDiff('yyyy',[Patients].[DOB],Date()) AS AgeNow, 
        DateDiff('yyyy',[Patients].[DOB],[Encounters].[Date]) AS EncounterAge, 
        MotionParams.GcdFile, PhysicalExams.GmfcsNew, 
        PhysicalExams.LeftForefootWtBearing, PhysicalExams.RightForefootWtBearing, PhysicalExams.LeftMidfootWtBearing, 
        PhysicalExams.RightMidfootWtBearing, PhysicalExams.LeftHindFootWtBearing, PhysicalExams.RightHindFootWtBearing, 
        PhysicalExamsMore.LeftHalluxWtBearing, PhysicalExamsMore.RightHalluxWtBearing, FootMilwaukeeAnglesEtc.LeftPhiHind, 
        FootMilwaukeeAnglesEtc.RightPhiHind, PhysicalExams.LeftFootSupple, PhysicalExams.RightFootSupple, 
        PhysicalExams.LeftOrthoticTolerated, PhysicalExams.RightOrthoticTolerated, PhysicalExams.LeftToneGastroc, 
        PhysicalExams.RightToneGastroc, PhysicalExams.LeftTonePeroneals, PhysicalExams.RightTonePeroneals, 
        PhysicalExams.LeftAnkleRomInver, PhysicalExams.RightAnkleRomInver, PhysicalExams.LeftAnkleRomEver, 
        PhysicalExams.RightAnkleRomEver, PhysicalExams.LeftThighFootAngle, PhysicalExams.RightThighFootAngle, 
        MotionParams.Side AS MotionParams_Side, MotionParams.AnkleDorsiPlantarMeanStanceDF, MotionParams.KneeFlexionAtIC, 
        PhysicalExams.LeftForefootOffWtBear, PhysicalExams.RightForefootOffWtBear, PhysicalExams.LeftMidfootOffWtBear, 
        PhysicalExams.RightMidfootOffWtBear, PhysicalExams.LeftHindFootOffWtBear, PhysicalExams.RightHindFootOffWtBear, 
        PhysicalExams.LeftFootOther, PhysicalExams.RightFootOther, PhysicalExams.LeftTonePeronealsClonus, 
        PhysicalExams.RightTonePeronealsClonus, PhysicalExams.LeftKneeRomPopAngle, PhysicalExams.RightKneeRomPopAngle, 
        PhysicalExams.LeftKneeRomPopAngleR1, PhysicalExams.RightKneeRomPopAngleR1, PhysicalExams.LeftKneeRomExt, 
        PhysicalExams.RightKneeRomExt, MotionParams.KneeFlexionMin, MotionParams.HipFlexionMin, 
        MotionParams.AnkleDorsiPlantarPeakDFStance, MotionParams.KneeFlexionMeanStance, 
        MotionParams.KneeFlexionMinIpsiStance, MotionParams.KneeFlexionTotalRange, 
        MotionParams.FootProgressionMeanStance, MotionParams.KneeRotationMeanStance, MotionParams.HipRotationMean, 
        MotionParams.PelvicRotationMean
    FROM Patients 
        INNER JOIN ((((Encounters 
                        INNER JOIN FootMilwaukeeAnglesEtc ON (Encounters.TestNumber = FootMilwaukeeAnglesEtc.TestNumber) 
                                                        AND (Encounters.Date = FootMilwaukeeAnglesEtc.Date) 
                                                        AND (Encounters.MRN = FootMilwaukeeAnglesEtc.MRN)) 
                        INNER JOIN PhysicalExams ON (Encounters.TestNumber = PhysicalExams.TestNumber) 
                                                AND (Encounters.Date = PhysicalExams.Date) 
                                                AND (Encounters.MRN = PhysicalExams.MRN)) 
                        INNER JOIN PhysicalExamsMore ON (Encounters.TestNumber = PhysicalExamsMore.TestNumber) 
                                                    AND (Encounters.Date = PhysicalExamsMore.Date) 
                                                    AND (Encounters.MRN = PhysicalExamsMore.MRN)) 
                        INNER JOIN MotionParams ON (Encounters.TestNumber = MotionParams.TestNumber)
                                                    AND (Encounters.Date = MotionParams.Date)
                                                    AND (Encounters.MRN = MotionParams.MRN))
                    ON Patients.MRN = Encounters.MRN
    WHERE Patients.PrimaryDiagnosis = ? AND Encounters.Date > #1/1/2005# AND Patients.PatternOfInvolvement Not Like 'Monoplegia' 
            AND Patients.PatternOfInvolvement Not Like 'Quadriplegia' 
        AND (Encounters.Date)=(SELECT MIN(Date) FROM Encounters AS E WHERE E.MRN=Encounters.MRN)
    """

    # Execute the query and fetch the results into a Pandas DataFrame
    df_q = pd.read_sql(sql_query, conn, params=primary_diagnosis)
    
    # For the two columns with dates, remove the timestamp
    df_q['Date'] = pd.to_datetime(df_q['Date']).dt.date
    df_q['DOB'] = pd.to_datetime(df_q['DOB']).dt.date

    # Close the cursor and connection
    cursor.close()
    conn.close()

    # Path to save the query exported
    file_path=f"\\\\CHI-FS-APP03\\Motion Analysis Center\\Gait Lab\\_Patients\\{mrn_input}\\{mrn_input}_DT_query.xlsx"
    #file_path=f"G:\\Gait Lab\\_Patients\\{mrn_input}\\{mrn_input}_DT_query.xlsx"
    folder_name.insert(0, file_path)
    # Export the DataFrame to an Excel file
    df_q.to_excel(file_path, index=False)

    error_label = Label(root, text="The query is now created", fg="green")
    error_label.place(x=75, y=330)
    ins_label = Label(root, text="Click on the Decision Tree you want\nor to get the Get Deviation Classification:", fg="blue")
    ins_label.place(x=90, y=440)

# Search button
buttonS = Button(root, text="Search", font=("Open Sans", 12, 'bold'), padx=20, pady=10, bg="#A7C4FF", command=findExcel)
buttonS.place(x=450, y=165)

# Browse button
ButtonB = Button(root, text="Browse", font=("Open Sans", 12, 'bold'), padx=20, pady=10, bg="#A7C4C2", command=browse_file)
ButtonB.place(x=300, y=355)

# Button for the Knee DT
buttonKneeDT = Button(root, text="Freeman Miller\nKnee Decision Tree", padx=20, pady=20, bg="#C8E1B5", command=KneeFlexDefor)
buttonKneeDT.place(x=90, y=500)

# button for the reference of the knee
buttonRefKnee = Button(root, text="View\nreference", font=("Open Sans", 7), command=openRef_knee)
buttonRefKnee.place(x=255, y=520)

# Button for the Foot DT
buttonFootDT = Button(root, text="Freeman Miller\nFoot Decision Tree", padx=20, pady=20, bg="#FBD8C5", command=Foot)
buttonFootDT.place(x=90, y=640)

# button for the reference of the foot
buttonRefFoot = Button(root, text="View\nreference", font=("Open Sans", 7), command=openRef_foot)
buttonRefFoot.place(x=255, y=660)

# Button for the Gait Deviation Classification
buttonGaitDT = Button(root, text="Gait Deviation\nClassification", padx=8, pady=15, bg="#FFF2CC", command=GaitDev)
buttonGaitDT.place(x=355, y=565)

# Button to run the whole query 
button_complete = Button(root, text="Run complete\nquery", padx=10, pady=10, bg="#D8D6E6", command=all_Data)
button_complete.place(x=310, y=420)

logoe_path = os.path.join(script_dir, 'resources', 'excel_logo.png')
logo_e = Image.open(logoe_path)
logo1_e=logo_e.resize((30,30), Image.LANCZOS) # the original size of the image is too big
logo2_e = ImageTk.PhotoImage(logo1_e)

# Button to export excel files
button_export_Ex = Button(root, text="Export Excel File", image=logo2_e, compound="left", padx=10, pady=15, bg="#C8F0DA", \
    command=lambda: save_Excel(knee_recommend, foot_recommend, sagittal_col, trans_col, selected_dates2))
button_export_Ex.place(x=310, y=800)

logo_n_path = os.path.join(script_dir, 'resources', 'logo_new.png')
# logo_n = Image.open('\\Users\\AGraf\\Desktop\\python_app\\logo_new.png')
logo_n = Image.open(logo_n_path)
logo1_n=logo_n.resize((30,30), Image.LANCZOS) # the original size of the image is too big
logo2_n = ImageTk.PhotoImage(logo1_n)

button_new= Button(root, text="New patient", image=logo2_n, compound="left", padx=10, pady=15, command = new_patient)
button_new.place(x=60, y=800)

# keep all above in the window all the time 
root.mainloop()
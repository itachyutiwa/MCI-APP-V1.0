
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
from ttkbootstrap import Style  
from ttkwidgets import Calendar

import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import openpyxl
import datetime as dt
import pandas as pd
import os


# Global variables to store the file paths
selected_beneficiaire_path = ""
selected_emission_path = ""
selected_consommation_path = ""

def upload_beneficiaire_file():
    global selected_beneficiaire_path
    file_path = filedialog.askopenfilename()
    if file_path:
        selected_beneficiaire_path = file_path
        beneficiaire_entry.delete(0, tk.END)
        beneficiaire_entry.insert(0, selected_beneficiaire_path)

def upload_emission_file():
    global selected_emission_path
    file_path = filedialog.askopenfilename()
    if file_path:
        selected_emission_path = file_path
        emission_entry.delete(0, tk.END)
        emission_entry.insert(0, selected_emission_path)

def upload_consommation_file():
    global selected_consommation_path
    file_path = filedialog.askopenfilename()
    if file_path:
        selected_consommation_path = file_path
        consommation_entry.delete(0, tk.END)
        consommation_entry.insert(0, selected_consommation_path)

# Function to print the selected file paths
def print_selected_file_paths():
    print("Selected BENEFICIAIRE File Path:", selected_beneficiaire_path)
    print("Selected EMISSION File Path:", selected_emission_path)
    print("Selected CONSOMMATION File Path:", selected_consommation_path)
    print("Start Date:", dt.datetime.strptime(str(start_date_entry.selection),'%Y-%m-%d %H:%M:%S').date())
    print("End Date:",  dt.datetime.strptime(str(end_date_entry.selection),'%Y-%m-%d %H:%M:%S').date())
    print("Seuil Coûteux:", seuil_entry.get())
    print("Montant Plafond:", montant_entry.get())

def age(x):
    if isinstance(x, str):
        return -5555555
    else:
        today = dt.datetime.strptime(str(end_date_entry.selection),'%Y-%m-%d %H:%M:%S').date()
        return today.year - x.year - ((today.month, today.day) < (x.month, x.day))
    
def critere_benef(file):
    nom_dossier = "RESULTATS"
    sous_dossier = "BENEFICIAIRE"
    if not os.path.exists(nom_dossier):
        os.mkdir(nom_dossier)
    if not os.path.exists(os.path.join(nom_dossier, sous_dossier)):
        os.mkdir(os.path.join(nom_dossier, sous_dossier))

    benef = pd.read_excel(file)
    benef_doublon = benef[benef.duplicated(keep=False)]
    benef_sans_doublon = benef.drop_duplicates()

    police_benef = benef_sans_doublon.groupby(["Num Police"])["Num Police"].agg({"count"}).reset_index()
    benef_sans_doublon["age"] = pd.to_datetime(benef_sans_doublon['Date Naissance']).apply(age)

    if benef_sans_doublon.shape[0] != 0:
        ENFANT_SUP_25 = benef_sans_doublon[benef_sans_doublon["Statut Ace"].isin(['E']) & (benef_sans_doublon['age'] > 25)]
        ADULTES_SUP_60 = benef_sans_doublon[benef_sans_doublon["Statut Ace"].ne('E') & (benef_sans_doublon['age'] > 60)]
    else:
        ENFANT_SUP_25 = benef_sans_doublon
        ADULTES_SUP_60 = benef_sans_doublon
    
    
    if len(benef_sans_doublon) !=0:  

            path_sans_doublon = os.path.join(nom_dossier, sous_dossier, "SANS_DOUBLONS BENEF.xlsx") 
            benef_sans_doublon.to_excel(path_sans_doublon)

    if len(benef_doublon) !=0:
        path_doublon = os.path.join(nom_dossier, sous_dossier, "DOUBLONS BENEF.xlsx")
        benef_doublon.to_excel(path_doublon)

    if len(police_benef) !=0:
            path_police_benef = os.path.join(nom_dossier, sous_dossier, "POLICE BENEF.xlsx")
            police_benef.to_excel(path_police_benef)

    if len(ENFANT_SUP_25) !=0:
            path_ENFANT_SUP_25 = os.path.join(nom_dossier, sous_dossier, "ENFANT_25.xlsx")
            ENFANT_SUP_25.to_excel(path_ENFANT_SUP_25)

    if len(ADULTES_SUP_60) !=0:
        path_ADULTES_SUP_60 = os.path.join(nom_dossier, sous_dossier, "ADULTE_60.xlsx")
        ADULTES_SUP_60.to_excel(path_ADULTES_SUP_60)
    return

def critere_emission(file):
    nom_dossier = "RESULTATS"
    sous_dossier = "EMISSION"
    if not os.path.exists(nom_dossier):
        os.mkdir(nom_dossier)
    if not os.path.exists(os.path.join(nom_dossier, sous_dossier)):
        os.mkdir(os.path.join(nom_dossier, sous_dossier))

    emission = pd.read_excel(file)
    doublons_emission = emission[emission.duplicated(keep=False)]
    sans_doublons_emission =emission.drop_duplicates()
    police_emission = emission.groupby(["Gestionnaire"])["Mt Prime Net"].agg({"sum","count"}).sort_values(by=['sum'], ascending=False).reset_index()
    
    if len(doublons_emission) !=0:
        path_doublon = os.path.join(nom_dossier, sous_dossier, "SANS_DOUBLONS BENEF.xlsx") 
        doublons_emission.to_excel(path_doublon)
        
    if len(sans_doublons_emission) !=0:
        path_sans_doublon = os.path.join(nom_dossier, sous_dossier, "SANS_DOUBLONS EMISSION.xlsx") 
        sans_doublons_emission.to_excel(path_sans_doublon)
       
    if len(police_emission) !=0:
        path_police_emission = os.path.join(nom_dossier, sous_dossier, "POLICE EMMISSION.xlsx") 
        police_emission.to_excel(path_police_emission)
    return

def critere_conso(file):
    nom_dossier = "RESULTATS"
    sous_dossier = "CONSOMMATION"
    if not os.path.exists(nom_dossier):
        os.mkdir(nom_dossier)
    if not os.path.exists(os.path.join(nom_dossier, sous_dossier)):
        os.mkdir(os.path.join(nom_dossier, sous_dossier))

    conso = pd.read_excel(file)
    doublons_conso = conso[conso.duplicated(keep=False)]
    sans_doublons_conso = conso.drop_duplicates()
    police_conso = conso.groupby(['Num Police'])['Montant Paye'].agg({"count","sum"}).sort_values(by=["count"], ascending=False).reset_index()
    
    if len(doublons_conso) != 0:
        path_doublon = os.path.join(nom_dossier, sous_dossier, "DOUBLONS CONSO.xlsx") 
        doublons_conso.to_excel(path_doublon)

    if len(sans_doublons_conso) != 0:
        path_sans_doublons_conso = os.path.join(nom_dossier, sous_dossier,"SANS_DOUBLONS CONSO.xlsx")
        sans_doublons_conso.to_excel(path_sans_doublons_conso)

    if len(police_conso) !=0:
        path_police_conso = os.path.join(nom_dossier, sous_dossier,"POLICE CONSO.xlsx")
        police_conso.to_excel(path_police_conso)
    return

def resume(file_benef, file_emission, file_conso, seuil_ctx):
    nom_dossier = "RESULTATS"
    sous_dossier = "RESUME"
    if not os.path.exists(nom_dossier):
        os.mkdir(nom_dossier)
    if not os.path.exists(os.path.join(nom_dossier, sous_dossier)):
        os.mkdir(os.path.join(nom_dossier, sous_dossier))

    benef = pd.read_excel(file_benef)
    emission = pd.read_excel(file_emission)
    conso = pd.read_excel(file_conso)

    benef_sans_doublon = benef.drop_duplicates()
    benef_sans_doublon["age"] = pd.to_datetime(benef_sans_doublon['Date Naissance']).apply(age)
    sans_doublons_emission =emission.drop_duplicates()
    sans_doublons_conso = conso.drop_duplicates()
    if benef_sans_doublon.shape[0] != 0:
        ENFANT_SUP_25 = benef_sans_doublon[benef_sans_doublon["Statut Ace"].isin(['E']) & (benef_sans_doublon['age'] > 25)]
        ADULTES_SUP_60 = benef_sans_doublon[benef_sans_doublon["Statut Ace"].ne('E') & (benef_sans_doublon['age'] > 60)]
    else:
        ENFANT_SUP_25 = benef_sans_doublon
        ADULTES_SUP_60 = benef_sans_doublon
    
    
    liste_police_anormale = set(sans_doublons_conso["Num Police"]).difference(set(sans_doublons_emission["Num Police"]))

    cond1 = sans_doublons_conso['Num Police'].isin(liste_police_anormale) & pd.to_datetime(sans_doublons_conso["Date Soins"]).gt(pd.to_datetime('2022-01-01 00:00:00'))
    police_anomale = sans_doublons_conso[cond1].astype({'Montant Paye': 'float'})
   
    police_anomale_resume = police_anomale.groupby(['Num Police'])['Montant Paye'].agg(['count', 'sum']).reset_index()

    if len(ADULTES_SUP_60) !=0:
        try:
            cond2 = sans_doublons_conso['Matricule Beneficiaire'].isin(ADULTES_SUP_60["Matricule"].unique())
            ADULTE_60_CONSO = sans_doublons_conso[cond2]
        except:
            print("ADULTES_SUP_60 vide")
    
    if len(ENFANT_SUP_25) !=0:
        try:
            cond3 = sans_doublons_conso['Matricule Beneficiaire'].isin(ENFANT_SUP_25["Matricule"].unique())
            ENFANT_25_CONSO = sans_doublons_conso[cond3]
        except:
            print("ENFANT_25_CONSO vide")

    cond4 = pd.to_datetime(sans_doublons_conso["Date Soins"]).lt(pd.to_datetime('2022-01-01 00:00:00')) & pd.to_datetime(sans_doublons_conso["Date Recepetion"]).gt(pd.to_datetime('2022-12-31 00:00:00'))
    PRESTATIONS_TARDIVES = sans_doublons_conso[cond4]
    
    cond5 = sans_doublons_conso['Montant Paye'].astype("float").gt(float(seuil_ctx))
    CONSO_COUTEUSES = sans_doublons_conso[cond5]

    if len(police_anomale) != 0:
        path_police_anomale =  os.path.join(nom_dossier, sous_dossier, "POLICES ANORMALES.xlsx")
        police_anomale.to_excel(path_police_anomale)

    if len(police_anomale_resume) != 0:   
        path_police_anomale_resume =  os.path.join(nom_dossier, sous_dossier, "POLICE ANORMALES RESUME.xlsx")
        police_anomale_resume.to_excel(path_police_anomale_resume)

    if len(PRESTATIONS_TARDIVES) != 0:
        path_PRESTATIONS_TARDIVES =  os.path.join(nom_dossier, sous_dossier, "CONSO TARDIVES.xlsx")
        PRESTATIONS_TARDIVES.to_excel(path_PRESTATIONS_TARDIVES)

    if len(CONSO_COUTEUSES) != 0:
        path_CONSO_COUTEUSES =  os.path.join(nom_dossier, sous_dossier, "CONSO COUTEUSES.xlsx")
        CONSO_COUTEUSES.to_excel(path_CONSO_COUTEUSES)
    return


def start_progress():
    progress_bar['value'] = 0
    critere_benef(selected_beneficiaire_path)
    critere_emission(selected_emission_path)
    critere_conso(selected_consommation_path)
    resume(selected_beneficiaire_path, selected_emission_path, selected_consommation_path,seuil_entry.get())
    update_progress()
    return

def update_progress():
    current_value = progress_bar['value']
    if current_value < 100:
        progress_bar['value'] += 5
        percentage = current_value + 5
        progress_label.config(text=f"{percentage}%")
        window.after(2000, update_progress)
    else:
        progress_label.config(text="Terminé")
    return


# Create the main window
window = tk.Tk()
window.title("********************** CONTROLE DELEGUE MCI-CARE COTE D'IVOIRE *************************")

# Create a ttkbootstrap style
style = Style(theme="superhero")

# Créez le cadre (frame) pour contenir tous les éléments de l'interface utilisateur
frame = ttk.Frame(window)
frame.pack(fill='both', expand=True)

# Ajoutez une barre de défilement verticale au cadre
scrollbar = ttk.Scrollbar(frame, orient='vertical')
scrollbar.pack(side='right', fill='y')

# Configurer la barre de défilement pour faire défiler le cadre
frame.config(yscrollcommand=scrollbar.set)
scrollbar.config(command=frame.yview)

# Add a label for CHARGEMENT DES FICHIERS
restrictions_label = ttk.Label(frame, text="CHARGEMENT DES FICHIERS", font=("Helvetica", 12, "bold"))
restrictions_label.grid(row=0, column=0, columnspan=3, padx=10, pady=10, sticky="w")

# Create Label, Entry field, and Upload button for BENEFICIAIRE
beneficiaire_label = ttk.Label(frame, text="BENEFICIAIRE:")
beneficiaire_label.grid(row=1, column=0, padx=10, pady=5)

beneficiaire_entry = ttk.Entry(frame, width=50)
beneficiaire_entry.grid(row=1, column=1, padx=10, pady=5)

beneficiaire_upload_button = ttk.Button(window, text="Chargez fichier BENEFICIAIRE", command=upload_beneficiaire_file)
beneficiaire_upload_button.grid(row=1, column=2, padx=10, pady=10)

# Create Label, Entry field, and Upload button for EMISSION
emission_label = ttk.Label(frame, text="EMISSION:")
emission_label.grid(row=2, column=0, padx=10, pady=5)

emission_entry = ttk.Entry(frame, width=50)
emission_entry.grid(row=2, column=1, padx=10, pady=5)

emission_upload_button = ttk.Button(frame, text="Chargez fichier EMISSION", command=upload_emission_file)
emission_upload_button.grid(row=2, column=2, padx=10, pady=10)

# Create Label, Entry field, and Upload button for CONSOMMATION
consommation_label = ttk.Label(frame, text="CONSOMMATION:")
consommation_label.grid(row=3, column=0, padx=10, pady=5)

consommation_entry = ttk.Entry(frame, width=50)
consommation_entry.grid(row=3, column=1, padx=10, pady=5)

consommation_upload_button = ttk.Button(frame, text="Chargez fichier CONSOMMATION", command=upload_consommation_file)
consommation_upload_button.grid(row=3, column=2, padx=10, pady=10)

# Add a label for PERIODE
restrictions_label = ttk.Label(frame, text="PERIODES", font=("Helvetica", 12, "bold"))
restrictions_label.grid(row=4, column=0, columnspan=3, padx=10, pady=10, sticky="w")

# Label and Entry for Start Date
date_label = ttk.Label(frame, text="Date début:")
date_label.grid(row=5, column=0, padx=10, pady=10, sticky="w")
start_date_entry = Calendar(frame)
start_date_entry.grid(row=5, column=1, padx=10, pady=10, sticky="w")

# Label and Entry for End Date
end_label = ttk.Label(frame, text="Date de fin:")
end_label.grid(row=6, column=0, padx=10, pady=10, sticky="w")
end_date_entry = Calendar(frame)
end_date_entry.grid(row=6, column=1, padx=10, pady=10, sticky="w")

# Add a label for RESTRICTIONS
restrictions_label = ttk.Label(frame, text="RESTRICTIONS", font=("Helvetica", 12, "bold"))
restrictions_label.grid(row=7, column=0, columnspan=3, padx=10, pady=10, sticky="w")

# Create IntVar variables for the checkboxes
benef_checkbox_var = tk.IntVar()
emission_checkbox_var = tk.IntVar()
conso_checkbox_var = tk.IntVar()

# Create a label for "Date de naissance" in italics
date_naissance_label = ttk.Label(frame, text="Date de naissance")
date_naissance_label.grid(row=8, column=0, padx=10, pady=5, sticky="w")

# Create an entry for "Antérieure à:"
anterieure_label = ttk.Label(frame, text="Antérieure à:")
anterieure_label.grid(row=9, column=1, padx=10, pady=5, sticky="e")

anterieure_entry = ttk.Entry(frame)
anterieure_entry.grid(row=9, column=2, padx=10, pady=5, sticky="w")

# Create a label for "Postérieure à:"
posterieure_label = ttk.Label(frame, text="Postérieure à:")
posterieure_label.grid(row=9, column=3, padx=10, pady=5, sticky="e")

posterieure_entry = ttk.Entry(frame)
posterieure_entry.grid(row=9, column=4, padx=10, pady=5, sticky="w")

# Create checkboxes for BENEF, EMISSION, and CONSO
benef_checkbox = ttk.Checkbutton(frame, text="BENEF.", variable=benef_checkbox_var)
benef_checkbox.grid(row=10, column=0, padx=10, pady=5, sticky="w")

emission_checkbox = ttk.Checkbutton(frame, text="EMISSION", variable=emission_checkbox_var)
emission_checkbox.grid(row=10, column=1, padx=10, pady=5, sticky="w")

conso_checkbox = ttk.Checkbutton(frame, text="CONSO.", variable=conso_checkbox_var)
conso_checkbox.grid(row=10, column=2, padx=10, pady=5, sticky="w")

# Label and Entry for Seuil coûteux
seuil_label = ttk.Label(frame, text="Seuil coûteux:")
seuil_label.grid(row=11, column=0, padx=10, pady=10, sticky="w")
seuil_entry = ttk.Entry(frame)
seuil_entry.grid(row=11, column=1, padx=10, pady=10, sticky="w")

# Label and Entry for Montant plafond
montant_label = ttk.Label(frame, text="Montant plafond:")
montant_label.grid(row=12, column=0, padx=10, pady=10, sticky="w")
montant_entry = ttk.Entry(frame)
montant_entry.grid(row=12, column=1, padx=10, pady=10, sticky="w")

# Add a button to start the control process
start_button = ttk.Button(frame, text="COMMENCER LE CONTROLE", style="danger.Outline.TButton", command=start_progress)
start_button.grid(row=13, column=0, columnspan=3, pady=10, padx=10, sticky="we")

# Create a progress bar
progress_bar = ttk.Progressbar(frame, orient="horizontal", length=200, mode="determinate", takefocus=True, maximum=100)
progress_bar.grid(row=14, column=0, columnspan=3, pady=10, padx=10, sticky="we")

# Add a label to display the percentage
progress_label = ttk.Label(frame, text="0%")
progress_label.grid(row=15, column=0, columnspan=3, pady=10, padx=10, sticky="we")
# Start the Tkinter main loop
window.mainloop()


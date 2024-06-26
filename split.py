from tkinter import *
from tkinter import filedialog
from tkinter.filedialog import askopenfile
import pandas as pd
import datetime

root = Tk()

root.title("Splitter_Table_Ca-gip")
root.geometry("400x400")
root.minsize(200,200)
root.config(background='light grey')
label_title = Label(root, text="Bienvenue !", font=("Courrier",30), bg='grey',fg='light yellow')
label_title.pack()

def choisirtable():
    global df2
    file = filedialog.askopenfilename()
    df = pd.read_excel(file)
    df2 = df.rename(columns={'Montant Conso':'Montant Conso (€)','Remise Comm':'Remise Comm (€)',
                       'Appel Nationaux':'Appel Nationaux (€)','Appel Opérateur':'Appel Opérateur (€)',
                       'Appel Autre Ope':'Appel Autre Ope (€)','Appel No Speciaux':'Appel No Speciaux (€)',
                       'Appel Recu Inter':'Appel Recu Inter (€)','Appel Emis Vers Inter':'Appel Emis Vers Inter (€)',
                       'Appel Emis depuis Inter':'Appel Emis depuis Inter (€)','Montant SMS':'Montant SMS (€)',
                       'DataNat':'DataNat (€)','Data Internat.':'Data Internat. (€)','Autre Presta':'Autre Presta (€)',
                       'Voix M ':'Voix M  (€)','Voix M-1':'Voix M-1 (€)','Voix M-2':'Voix M-2 (€)','SMS M':'SMS M (€)',
                       'SMS M-1':'SMS M-1 (€)','SMS M-2':'SMS M-2 (€)','Moyenne Voix':'Moyenne Voix (€)',
                       'Moyenne SMS':'Moyenne SMS (€)','Montant Conso':'Montant Conso (€)',
                       'N° spéciaux et autres services voix':'N° spéciaux et autres services voix (€)',
                       'Monde':'Monde (€)','Abo + options Profil':'Abo + options Profil (€)',
                       'Total Conso (Voix +SMS) Old':'Total Conso (Voix +SMS) Old (€)'})
    df2['Date Facture'] = pd.to_datetime(df2['Date Facture']).dt.date

def splitter():
    split_values = df2['Entité'].unique()
    for value in split_values:
        df3 = df2[df2['Entité'] == value]
        output_file_name = str(value) + ".xls"
        df3.to_excel(output_file_name, index=False)
        
bouton1 = Button(root,text="Choisir le fichier Table",command=choisirtable)
bouton1.pack(pady=15)

bouton2= Button(root,text="Splitter les données par entité",command=splitter)
bouton2.pack(pady=15)

root.mainloop()

from tkinter import *
from tkinter import filedialog
from tkinter.filedialog import askopenfile
import sqlite3
import openpyxl
import pandas as pd
import numpy as np

root = Tk()

root.title("Rapport CES pour Lexmark, OBS, Bouygues, BT et SFR")
root.geometry("500x500")
root.minsize(350,450)
root.config(background='light blue')
label_title = Label(root, text="Bienvenue !", font=("Courrier",40), bg='light blue',fg='light yellow')
label_title.pack()

def chooselex():
#    global dflex
    file = filedialog.askopenfilename()
    dflex = pd.read_excel(file)
    
    Lexmark1 = dflex[dflex['Type de charge']=='ABO']
    Lexmark2 = pd.pivot_table(Lexmark1, values = 'Montant HT', index = ['AAAAMM'], columns = 'Opérateur', aggfunc = np.sum, margins = 'TRUE', margins_name='Grand Total')
    
    Lexmark3 = dflex[dflex['Type de charge']=='CON']
    Lexmark4 = pd.pivot_table(Lexmark3, values = 'Montant HT', index = ['AAAAMM'], columns = 'Opérateur', aggfunc = np.sum, margins = 'TRUE', margins_name='Grand Total')
    
    Lexmark5 = pd.pivot_table(Lexmark1, values = 'Montant HT', index = 'Produit facturé', columns = 'AAAAMM', aggfunc = np.sum,
                             margins = True, margins_name = 'Grand Total')
    
    Lexmark6 = dflex[(dflex['Produit facturé']=='Color') | (dflex['Produit facturé']=='LTPC') | (dflex['Produit facturé']=='Mono')]
    Lexmark7 = pd.pivot_table(Lexmark6, values = 'Montant HT', index = 'Produit facturé', columns = 'AAAAMM', aggfunc = np.sum,
                             margins = True, margins_name = 'Totals')
    
    Lexmark8 = dflex[dflex['Type de charge']=='ABO']
    Lexmark9 = pd.pivot_table(Lexmark8, values = 'NDI', index = ['Produit facturé'], columns = 'AAAAMM', aggfunc = lambda x: len(x.unique()),
                             margins = True, margins_name = "Grand Total")
    Lexmark10 = Lexmark9.fillna('')
    
    Lexmark11 = dflex.rename(columns={'Description':'Sites', 'Sous-Compte':'Coda', 'NDI':'Nombre imprimantes par site'})
    Lexmark12 = pd.pivot_table(Lexmark11, values = ['Nombre imprimantes par site'], index = ['Sites','Coda'], columns = 'AAAAMM',
                               aggfunc = lambda x: len(x.unique()),margins=True,margins_name="Grand Total" )
    Lexmark13 = Lexmark12.fillna('')
    
    Lexmark14 = pd.pivot_table(Lexmark11, values = ['Montant HT'], index = ['Sites','Coda'], columns = 'AAAAMM', aggfunc = np.sum, margins = 'True', margins_name="Grand Total")
    Lexmark15 = Lexmark14.fillna('')
    
    Lexmark16 = dflex.rename(columns={'NDI':'Imprimantes', 'Sous-Compte':'Coda', 'Description':'Nom du site', 'Produit facturé':'Abonnement'})
    Lexmark17 = pd.pivot_table(Lexmark16, values = ['Montant HT'], index = ['Imprimantes','Coda','Nom du site','Abonnement'],
                               columns = 'AAAAMM', aggfunc = np.sum, margins=True, margins_name="Grand Total")
    Lexmark18 = Lexmark17.fillna('')
    
    with pd.ExcelWriter('Lexmark.xlsx') as writer:
        Lexmark2.to_excel(writer, sheet_name="LEXMARK AAAAMM -Total A,A-1,A-2",startcol = 0)
        Lexmark4.to_excel(writer, sheet_name="LEXMARK AAAAMM -Total A,A-1,A-2",startcol = 4)
        Lexmark5.to_excel(writer, sheet_name="ABONNEMENT")
        Lexmark7.to_excel(writer, sheet_name="CONSOMMATIONS")
        Lexmark10.to_excel(writer, sheet_name="NOMBRES IMPRIMANTES")
        Lexmark13.to_excel(writer, sheet_name="Facturation par Sites-NDI")
        Lexmark15.to_excel(writer, sheet_name="Facturation par Sites-M_ HT")
        Lexmark18.to_excel(writer, sheet_name="Imprimantes par sites")
    
def chooseobs():
#    global dfobs
    file = filedialog.askopenfilename()
    dfobs = pd.read_excel(file)

    Obs1 = dfobs[dfobs['Type de charge']=='ABO']
    Obs2 = pd.pivot_table(Obs1, values = 'Montant HT', index = ['AAAAMM'], columns = 'Opérateur', aggfunc = np.sum)
    
    Obs3 = dfobs[dfobs['Type de charge']=='REM']
    Obs4 = pd.pivot_table(Obs3, values = 'Montant HT', index = ['AAAAMM'], columns = 'Opérateur', aggfunc = np.sum)
    
    Obs5 = dfobs[dfobs['Type de charge']=='CON']
    Obs6 = pd.pivot_table(Obs5, values = 'Montant HT', index = ['AAAAMM'], columns = 'Opérateur', aggfunc = np.sum)
    
    Obs7 = dfobs[dfobs['Type de charge']=='AUT']
    Obs8 = pd.pivot_table(Obs7, values = 'Montant HT', index = ['AAAAMM'], columns = 'Opérateur', aggfunc = np.sum)
    
    Obs9 = dfobs[dfobs['Type de charge']=='PRE']
    Obs10 = pd.pivot_table(Obs9, values = 'Montant HT', index = ['AAAAMM'], columns = 'Opérateur', aggfunc = np.sum)
    
    Obs11 = dfobs[(dfobs['Type de charge']=='ABO') | (dfobs['Type de charge']=='CON')]
    Obs12 = pd.pivot_table(Obs11, values = ['Montant HT'], index = ['Type de charge','Produit facturé'], columns = 'AAAAMM',
                      aggfunc = np.sum, margins=True, margins_name="Grand Total")
    Obs13= Obs12.fillna('')    
    
    Obs14 = dfobs[dfobs['Service']=='Accueil']
    Obs15 = pd.pivot_table(Obs14, values = ['Montant HT'], index = ['Produit facturé'], columns = 'AAAAMM', aggfunc = np.sum, margins = True, margins_name = "Grand Total")
    Obs16 = Obs15.fillna('')
    
    Obs17 = dfobs[(dfobs['Type de charge']=='ABO') | (dfobs['Type de charge']=='AUT') | (dfobs['Type de charge']=='CON') | (dfobs['Service']=='Transfert de données')]
    Obs18 = pd.pivot_table(Obs17, values = ['Montant HT'], index = ['Type de charge','Produit facturé'], columns = 'AAAAMM',
                      aggfunc = np.sum, margins=True, margins_name = "Grand Total")
    Obs19 = Obs18.fillna('')
    
    Obs20 = dfobs[dfobs['Type de charge']=='CON']
    Obs21 = pd.pivot_table(Obs20, values = ['Montant HT'], index = ['Produit facturé'], columns = 'AAAAMM',
                      aggfunc = np.sum, margins=True, margins_name = "Grand Total")
    Obs22 = Obs21.fillna('')

    Obs23 = dfobs[dfobs['Type de charge']=='PRE']
    Obs24 = pd.pivot_table(Obs23, values = ['Montant HT'], index = ['Produit facturé'], columns = 'AAAAMM',
                      aggfunc = np.sum, margins=True, margins_name = "Grand Total")
    Obs25 = Obs24.fillna('')
    
    Obs26 = dfobs[dfobs['Service']=='Internet']
    Obs27 = pd.pivot_table(Obs26, values = ['Montant HT'], index = ['Produit facturé'], columns = 'AAAAMM',
                      aggfunc = np.sum, margins=True, margins_name = "Grand Total")
    Obs28 = Obs27.fillna('')
        
    with pd.ExcelWriter('OBS.xlsx') as writer:
        Obs2.to_excel(writer, sheet_name="O.B.S AAAAMM TOTAL A,A-1,A-2", startcol=0)
        Obs4.to_excel(writer, sheet_name="O.B.S AAAAMM TOTAL A,A-1,A-2", startcol=3)
        Obs6.to_excel(writer, sheet_name="O.B.S AAAAMM TOTAL A,A-1,A-2", startcol=6)
        Obs8.to_excel(writer, sheet_name="O.B.S AAAAMM TOTAL A,A-1,A-2", startcol=9)
        Obs10.to_excel(writer, sheet_name="O.B.S AAAAMM TOTAL A,A-1,A-2", startcol=12)
        Obs13.to_excel(writer, sheet_name="DISMO OBS")
        Obs16.to_excel(writer, sheet_name="ACCUEIL")
        Obs19.to_excel(writer, sheet_name="TRANSFERTS DE DONNEES")
        Obs22.to_excel(writer, sheet_name="CONSOMMATIONS")
        Obs25.to_excel(writer, sheet_name="PRESTATIONS")
        Obs28.to_excel(writer, sheet_name="INTERNET")

def choosebyt():
    global dfbyt
    file = filedialog.askopenfilename()
    dfbyt = pd.read_excel(file)
    
    Byt1 = pd.pivot_table(dfbyt, values = ['Montant HT'], index = ['Compte'], columns = 'AAAAMM', aggfunc = np.sum,
                          margins = True, margins_name = "Grand Total")
    Byt2 = Byt1.fillna('')
    
    Byt3 = dfbyt[(dfbyt['Type de charge']=='ABO') | (dfbyt['Type de charge']=='CON')]
    Byt4 = pd.pivot_table(Byt3, values = ['Montant HT'], index = ['Compte','Produit facturé'], columns = 'AAAAMM',
                          aggfunc = np.sum, margins = True, margins_name = "Grand Total")
    Byt5 = Byt4.fillna('')
    
    with pd.ExcelWriter('Bouygues.xlsx') as writer:
        Byt2.to_excel(writer, sheet_name="Compte Facturation")
        Byt5.to_excel(writer, sheet_name="Détail Produits Facturés")
#       Byt6.to_excel(writer, sheet_name="Facturation par site")
    
def choosebt():
    global dfbt
    file = filedialog.askopenfilename()
    dfbt = pd.read_excel(file)

    Bt1 = dfbt[dfbt['Type de charge']=='ABO']
    Bt2 = pd.pivot_table(Bt1, values = 'Montant HT', index = ['AAAAMM'], columns = 'Opérateur', aggfunc = np.sum)
    
    Bt3 = dfbt[dfbt['Type de charge']=='CON']
    Bt4 = pd.pivot_table(Bt3, values = 'Montant HT', index = ['AAAAMM'], columns = 'Opérateur', aggfunc = np.sum)
    
    Bt5 = dfbt[dfbt['Type de charge']=='PRE']
    Bt6 = pd.pivot_table(Bt5, values = 'Montant HT', index = ['AAAAMM'], columns = 'Opérateur', aggfunc = np.sum)
        
    Bt7 = dfbt[dfbt['Compte']=='564-10786']
    Bt8 = pd.pivot_table(Bt7, values = 'Montant HT', index = ['Produit facturé'], columns = 'AAAAMM', aggfunc = np.sum, margins = True, margins_name = "Grand Total")
    Bt9 = Bt8.fillna('')
    
    Bt13 = dfbt[dfbt['Compte']=='564-10657']
    Bt14 = pd.pivot_table(Bt13, values = ['Durée (s)','Montant HT'], index = ['Produit facturé'], columns = 'AAAAMM',
                      aggfunc = np.sum, margins=True, margins_name="Grand Total")
    Bt15 = Bt14.dropna()
    
    Bt16 = pd.pivot_table(Bt1, values = ['Montant HT'], index = ['Produit facturé'], columns = 'AAAAMM',
                      aggfunc = np.sum, margins=True, margins_name="Grand Total")
    Bt17 = Bt16.fillna('')
    
    Bt10 = dfbt[dfbt['Service']=='Transfert de données']
    Bt11 = pd.pivot_table(Bt10, values = ['Montant HT'], index = ['NDI'], columns = 'AAAAMM', aggfunc = np.sum, margins = True, margins_name = "Grand Total")
    Bt12 = Bt11.fillna('')
    
    with pd.ExcelWriter('BT.xlsx') as writer:
        Bt2.to_excel(writer, sheet_name="BT AAAAMM TOTAL A,A-1,A-2", startcol=0)
        Bt4.to_excel(writer, sheet_name="BT AAAAMM TOTAL A,A-1,A-2", startcol=3)
        Bt6.to_excel(writer, sheet_name="BT AAAAMM TOTAL A,A-1,A-2", startcol=6)
        Bt9.to_excel(writer, sheet_name="DISMO BT")
        Bt15.to_excel(writer, sheet_name="CONSOMMATIONS")
        Bt17.to_excel(writer, sheet_name="ABONNEMENT")
        Bt12.to_excel(writer, sheet_name="Site Actifs")

def choosesfr():
    global dfsfr
    file = filedialog.askopenfilename()
    dfsfr = pd.read_excel(file)
    
    Sfr1 = dfsfr[dfsfr['Type de charge']=='ABO']
    Sfr2 = pd.pivot_table(Sfr1, values = 'Montant HT', index = ['AAAAMM'], columns = 'Opérateur', aggfunc = np.sum,
                          margins = True, margins_name = "Grand Total")
    
    Sfr3 = dfsfr[dfsfr['Type de charge']=='CON']
    Sfr4 = pd.pivot_table(Sfr3, values = 'Montant HT', index = ['AAAAMM'], columns = 'Opérateur', aggfunc = np.sum,
                          margins = True, margins_name = "Grand Total")  
    
    Sfr5 = dfsfr[dfsfr['Type de charge']=='AUT']
    Sfr6 = pd.pivot_table(Sfr5, values = 'Montant HT', index = ['AAAAMM'], columns = 'Opérateur', aggfunc = np.sum,
                          margins = True, margins_name = "Grand Total")
    
    for i in dfsfr['Type de charge']:
        if 'PRE' in i:
            Sfr7 = dfsfr[dfsfr['Type de charge']=='PRE']
            Sfr8 = pd.pivot_table(Sfr7, values = 'Montant HT', index = ['AAAAMM'], columns = 'Opérateur', aggfunc = np.sum,
                          margins = True, margins_name = "Grand Total")
            with pd.ExcelWriter('SFR.xlsx') as writer:
                Sfr8.to_excel("SFR AAAAMM -TOTAL A,A-1,A-2", startcol=12)
        else:
            pass
            
    Sfr9 = pd.pivot_table(Sfr1, values = 'Montant HT', index = ['Produit facturé'], columns = 'AAAAMM', aggfunc = np.sum,
                           margins = True, margins_name = "Grand Total")    

    Sfr10 = pd.pivot_table(Sfr3, values = 'Montant HT', index = ['Produit facturé'], columns = 'AAAAMM', aggfunc = np.sum,
                           margins = True, margins_name = "Grand Total")        
    
    with pd.ExcelWriter('SFR.xlsx') as writer:
        Sfr2.to_excel(writer, sheet_name="SFR AAAAMM -TOTAL A,A-1,A-2", startcol=0)
        Sfr4.to_excel(writer, sheet_name="SFR AAAAMM -TOTAL A,A-1,A-2", startcol=4)
        Sfr6.to_excel(writer, sheet_name="SFR AAAAMM -TOTAL A,A-1,A-2", startcol=8)
        Sfr9.to_excel(writer, sheet_name="ABONNEMENT")
        Sfr10.to_excel(writer, sheet_name="CONSOMMATIONS")
        
myButtonlex = Button(root,text="Lecture du fichier Lexmark",command=chooselex)
myButtonlex.pack(pady=20)

myButtonobs = Button(root,text="Lecture du fichier OBS",command=chooseobs)
myButtonobs.pack(pady=20)

myButtonbyt = Button(root,text="Lecture du fichier Bouygues",command=choosebyt)
myButtonbyt.pack(pady=20)

myButtonbt = Button(root,text="Lecture du fichier BT",command=choosebt)
myButtonbt.pack(pady=20)

myButtonsfr = Button(root,text="Lecture du fichier SFR",command=choosesfr)
myButtonsfr.pack(pady=20)

root.mainloop()

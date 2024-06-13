import shutil
import pandas as pd
import os
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
import xlsxwriter
import re
import fnmatch
gui = Tk()
gui.geometry("400x400")
gui.title("FC")

def getFolderPath():
    folder_selected = filedialog.askdirectory()
    folderPath.set(folder_selected)

def classer():
    folder = folderPath.get()
#    print("Doing stuff with folder", folder)
    names = os.listdir(folder)
    folder_name = ['Abo','CAT_Mtt','CAT_Unite','Facture','TypeAppel_Mtt','TypeAppel_Unite']
    for x in range(0,6):
        os.makedirs(folder +'/'+ folder_name[x])
        
    for files in names:
        for i in range(0,6):
            if folder_name[i] in files and not os.path.exists(folder + '/' + folder_name[i] + '/' + files):
                shutil.move(folder + '/' + files, folder + '/' + folder_name[i]+ '/' +files)
                
def facture():
    global dffacture
    file = filedialog.askopenfilename()
    dffacture = pd.read_csv(file,sep = ";",encoding='unicode_escape')

def typeappelmtt():
    global dfmtt
    file = filedialog.askopenfilename()
    dfmtt = pd.read_csv(file,sep = ";",encoding='unicode_escape')
    
def typeappelunite():
    global dfunite
    file = filedialog.askopenfilename()
    dfunite = pd.read_csv(file,sep = ";",encoding='unicode_escape')
    
def final():
    workbook = xlsxwriter.Workbook('20AAMM_SFR_.xlsx')
    worksheet = workbook.add_worksheet()
    bold = workbook.add_format({'bold': True})
    
    worksheet.write('A1', 'Nom client', bold)
    worksheet.write('B1', 'Filiale', bold)
    worksheet.write('C1', 'Operateur', bold)
    worksheet.write('D1', 'NoTitulaire', bold)
    worksheet.write('E1', 'NoContrat', bold)
    worksheet.write('F1', 'No ligne', bold)
    worksheet.write('G1', 'Code PF', bold)
    worksheet.write('H1', 'CodeListe', bold)
    worksheet.write('I1', 'Date Ouverture', bold)
    worksheet.write('J1', 'Date Résiliation', bold)
    worksheet.write('K1', 'Date FDC', bold)
    worksheet.write('L1', 'Type Ligne', bold)
    worksheet.write('M1', 'Abo Principal', bold)
    worksheet.write('N1', 'Option BB', bold)
    worksheet.write('O1', 'Option Iphone', bold)
    worksheet.write('P1', 'Mail && Surf', bold)
    worksheet.write('Q1', 'Option SMS', bold)
    worksheet.write('R1', 'OptionSeuil', bold)
    worksheet.write('S1', 'MontantSeuil (€)', bold)
    worksheet.write('T1', 'Date Facture', bold)
    worksheet.write('U1', 'Montant Abo (€)', bold)
    worksheet.write('V1', 'Montant Conso (€)', bold)
    worksheet.write('W1', 'Montant divers (€)', bold)
    worksheet.write('X1', 'Montant HT (€)', bold)
    worksheet.write('Y1', 'Montant TVA (€)', bold)
    worksheet.write('Z1', 'Montant TTC (€)', bold)
    worksheet.write('AA1', 'Remise Comm (€)', bold)
    worksheet.write('AB1', 'Appels Nationaux (€)', bold)
    worksheet.write('AC1', 'Appel Opérateur (€)', bold)
    worksheet.write('AD1', 'Appel Autre Ope (€)', bold)
    worksheet.write('AE1', 'Appel No Speciaux (€)', bold)
    worksheet.write('AF1', "Appel Recu Inter (€)", bold)
    worksheet.write('AG1', 'Appel Emis vers Inter (€)', bold)
    worksheet.write('AH1', 'Appel Emis depuis Inter (€)', bold)
    worksheet.write('AI1', 'Montant SMS (€)', bold)
    worksheet.write('AJ1', 'Montant Data BB Nat (€)', bold)
    worksheet.write('AK1', 'Montant Data BB Internat (€)', bold)
    worksheet.write('AL1', 'DataNat (€)', bold)
    worksheet.write('AM1', 'Data Internat. (€)', bold)
    worksheet.write('AN1', 'Autre Presta (€)', bold)
    worksheet.write('AO1', 'Durée Appel Nat', bold)
    worksheet.write('AP1', 'Unit Appel Opérateur', bold)
    worksheet.write('AQ1', 'Unit Appel Autre Ope', bold)
    worksheet.write('AR1', 'Unit No Speciaux', bold)
    worksheet.write('AS1', "Unit Appel Recu Inter", bold)
    worksheet.write('AT1', 'Unit Appel Emis vers Inter', bold)
    worksheet.write('AU1', 'Unit Appel Emis Depuis Inter', bold)
    worksheet.write('AV1', 'Unit SMS', bold)
    worksheet.write('AW1', 'Volume Data BB Nat', bold)
    worksheet.write('AX1', 'Volume Data BB Internat', bold)
    worksheet.write('AY1', 'Volume Data Nat', bold)
    worksheet.write('AZ1', 'Volume Data Internat.', bold)
    worksheet.write('BA1', 'Unit Autre Presta', bold)
    worksheet.write('BB1', 'Nom Utilisateur', bold)
    worksheet.write('BC1', 'Matricule', bold)
    worksheet.write('BD1', 'Code1', bold)
    worksheet.write('BE1', 'Carte SIM', bold)
    worksheet.write('BF1', 'Pin1', bold)
    worksheet.write('BG1', 'Pin2', bold)
    worksheet.write('BH1', 'PUK1', bold)
    worksheet.write('BI1', 'PUK2', bold)
    worksheet.write('BJ1', 'IMEI', bold)
    worksheet.write('BK1', 'Marque', bold)
    worksheet.write('BL1', 'Modele', bold)
    worksheet.write('BM1', 'Date Réengagement', bold)
    worksheet.write('BN1', 'NoFacture', bold)
    worksheet.write('BO1', 'Sans Conso', bold)
    worksheet.write('BP1', 'LIbCode1', bold)
    worksheet.write('BQ1', 'Code2', bold)
    worksheet.write('BR1', 'LibCode2', bold)
    worksheet.write('BS1', 'Code3', bold)
    worksheet.write('BT1', 'LibCode3', bold)
#fac
    dffacture1 = dffacture.filter(['N° de ligne','Nom Utilisateur','N° de titulaire','Code PF','Code liste','N° de contrat',
                                   'N° de facture','Date de Facture','Mnt total HT net','Mnt total abo HT net',
                                   'Mnt total consos HT net','Mnt total divers HT net','Mnt HT net Facture  Entreprise'])
    dffacture2 = dffacture1.fillna('')
    dffacture3 = dffacture2[(dffacture2['Nom Utilisateur'] != 'Nom Utilisateur') & (dffacture2['Nom Utilisateur'] != '')
                            & (dffacture2['Nom Utilisateur'] != '')]
    dffacture4 = dffacture3.sort_values(by = 'N° de ligne', ascending = True)
#mtt
    dfmtt1 = dfmtt.filter(['N° de ligne','Appel Nationaux Fixe HT Brut','Appel SFR HT Brut','Appel Autres opérateurs HT Brut',
                           'Appels Numéro Spéciaux HT Brut','Appel reçu à l international HT Brut',
                           'Appel Emis vers l international depuis la france HT Brut',"Appel Emis depuis l'international HT Brut",
                           'Nb de sms HT Brut','Quantité de data en bb (Nat) HT Brut',
                           'Quantité de data en bb (Inter) HT Brut','Quantité data (Nat) HT Brut',
                           'Quantité data (Inter) HT Brut','Autres Prestation HT Brut'])
    dfmtt2 = dfmtt1.fillna('')
    dfmtt3 = dfmtt2[(dfmtt2['N° de ligne'] != 'N° de ligne') & (dfmtt2['N° de ligne'] != '')
                           & (dfmtt2['N° de ligne'] != '')]
    dfmtt4 = dfmtt3.sort_values(by = 'N° de ligne', ascending = True)
#unite
    dfunite1 = dfunite.drop(['N° de titulaire','N° de contrat','N° de facture','Date de Facture',
               'Date ouverture de ligne','Date de résillation'], axis=1)
    dfunite2 = dfunite1.fillna('')
    dfunite3 = dfunite2[(dfunite2['N° de ligne'] != 'N° de ligne') & (dfunite1['N° de ligne'] != '')
                       & (dfunite2['N° de ligne'] != '')]
    dfunite4 = dfunite3.sort_values(by = 'N° de ligne', ascending = True)
# d e f t u v w x bn
    dffacture5 = dffacture4['N° de titulaire']
    dffacture6 = dffacture5.values.tolist()
    rowD = 1
    colD = 3
    for d in dffacture6:
        worksheet.write(rowD,colD, d)
        rowD +=1
        
    dffacture7 = dffacture4['N° de contrat']
    dffacture8 = dffacture7.values.tolist()
    rowE = 1
    colE = 4
    for e in dffacture8:
        worksheet.write(rowE,colE, e)
        rowE +=1
        
    dffacture9 = dffacture4['N° de ligne'].astype(float)
    dffacture10 = dffacture9.values.tolist()
    rowF = 1
    colF = 5
    for f in dffacture10:
        worksheet.write(rowF,colF, f)
        rowF +=1
    
    dffacture11 = dffacture4['Date de Facture']
    dffacture12 = dffacture11.values.tolist()
    rowT = 1
    colT = 19
    for t in dffacture12:
        worksheet.write(rowT,colT, t)
        rowT +=1
    
    dffacture15 = dffacture4['Mnt total abo HT net'].str.replace(',','.').astype(float)
    dffacture16 = dffacture15.values.tolist()
    rowU = 1
    colU = 20
    for u in dffacture16:
        worksheet.write(rowU,colU, u)
        rowU +=1
        
    dffacture17 = dffacture4['Mnt total consos HT net'].str.replace(',','.').astype(float)
    dffacture18 = dffacture17.values.tolist()
    rowV = 1
    colV = 21
    for v in dffacture18:
        worksheet.write(rowV,colV, v)
        rowV +=1
        
    dffacture19 = dffacture4['Mnt total divers HT net'].str.replace(',','.').astype(float)
    dffacture20 = dffacture19.values.tolist()
    rowW = 1
    colW = 22
    for w in dffacture20:
        worksheet.write(rowW,colW, w)
        rowW +=1
        
    dffacture21 = dffacture4['Mnt total HT net'].str.replace(',','.').astype(float)
    dffacture22 = dffacture21.values.tolist()
    rowX = 1
    colX = 23
    for x in dffacture22:
        worksheet.write(rowX,colX, x)
        rowX +=1
        
    dffacture25 = dffacture4['Mnt total HT net'].str.replace(',','.').astype(float)
    dffacture25['TVA'] = pd.to_numeric(dffacture25,errors='coerce') * 0.2
    dffacture26 = dffacture25['TVA'].values.tolist()
    rowY = 1
    colY = 24
    for y in dffacture26:
        worksheet.write(rowY,colY, y)
        rowY +=1

    dffacture26 = dffacture4['Mnt total HT net'].str.replace(',','.').astype(float)
    dffacture26['TOTAL'] = pd.to_numeric(dffacture26,errors='coerce') * 1.2
    dffacture27 = dffacture26['TOTAL'].values.tolist()
    rowZ = 1
    colZ = 25
    for z in dffacture27:
        worksheet.write(rowZ,colZ, z)
        rowZ +=1
        
    dffacture13 = dffacture4['Nom Utilisateur']
    dffacture14 = dffacture13.values.tolist()
    rowBB = 1
    colBB = 53
    for bb in dffacture14:
        worksheet.write(rowBB,colBB, bb)
        rowBB +=1
        
    dffacture23 = dffacture4['N° de facture']
    dffacture24 = dffacture23.values.tolist()
    rowBN = 1
    colBN = 65
    for bn in dffacture24:
        worksheet.write(rowBN,colBN, bn)
        rowBN +=1
# ab ac af ag ah 
    dfmtt5 = dfmtt4['Appel reçu à l international HT Brut'].str.replace(',','.').astype(float)
    dfmtt6 = dfmtt5.values.tolist()
    rowAF = 1
    colAF = 31
    for af in dfmtt6:
        worksheet.write(rowAF,colAF, af)
        rowAF +=1
        
    dfmtt7 = dfmtt4['Appel Emis vers l international depuis la france HT Brut'].str.replace(',','.').astype(float)
    dfmtt8 = dfmtt7.values.tolist()
    rowAG = 1
    colAG = 32
    for ag in dfmtt8:
        worksheet.write(rowAG,colAG, ag)
        rowAG +=1
        
    dfmtt9 = dfmtt4["Appel Emis depuis l'international HT Brut"].str.replace(',','.').astype(float)
    dfmtt10 = dfmtt9.values.tolist()
    rowAH = 1
    colAH = 33
    for ah in dfmtt10:
        worksheet.write(rowAH,colAH, ah)
        rowAH +=1
        
    dfmtt11 = dfmtt4["Appel Nationaux Fixe HT Brut"].str.replace(',','.').astype(float)
    dfmtt12 = dfmtt11.values.tolist()
    rowAB = 1
    colAB = 27
    for ab in dfmtt12:
        worksheet.write(rowAB,colAB, ab)
        rowAB +=1
    
    dfmtt13 = dfmtt4["Appel SFR HT Brut"].str.replace(',','.').astype(float)
    dfmtt14 = dfmtt13.values.tolist()
    rowAC = 1
    colAC = 28
    for ac in dfmtt14:
        worksheet.write(rowAC,colAC, ac)
        rowAC +=1
    
    dfmtt15 = dfmtt4["Appel Autres opérateurs HT Brut"].str.replace(',','.').astype(float)
    dfmtt16 = dfmtt15.values.tolist()
    rowAD = 1
    colAD = 29
    for ad in dfmtt16:
        worksheet.write(rowAD,colAD, ad)
        rowAD +=1
        
    dfmtt17 = dfmtt4["Appels Numéro Spéciaux HT Brut"].str.replace(',','.').astype(float)
    dfmtt18 = dfmtt17.values.tolist()
    rowAE = 1
    colAE = 30
    for ae in dfmtt18:
        worksheet.write(rowAE,colAE, ae)
        rowAE +=1
        
    dfmtt19 = dfmtt4["Nb de sms HT Brut"].str.replace(',','.').astype(float)
    dfmtt20 = dfmtt19.values.tolist()
    rowAI = 1
    colAI = 34
    for ai in dfmtt20:
        worksheet.write(rowAI,colAI, ai)
        rowAI +=1
    
    dfmtt21 = dfmtt4["Quantité data (Nat) HT Brut"].str.replace(',','.').astype(float)
    dfmtt22 = dfmtt21.values.tolist()
    rowAL = 1
    colAL = 37
    for al in dfmtt22:
        worksheet.write(rowAL,colAL, al)
        rowAL +=1

    dfmtt23 = dfmtt4["Quantité data (Inter) HT Brut"].str.replace(',','.').astype(float)
    dfmtt24 = dfmtt23.values.tolist()
    rowAM = 1
    colAM = 38
    for am in dfmtt24:
        worksheet.write(rowAM,colAM, am)
        rowAM +=1
        
    dfmtt25 = dfmtt4["Autres Prestation HT Brut"].str.replace(',','.').astype(float)
    dfmtt26 = dfmtt25.values.tolist()
    rowAN = 1
    colAN = 39
    for an in dfmtt26:
        worksheet.write(rowAN,colAN, an)
        rowAN +=1
# g h ao:ba
    dfunite5 = dfunite4['Code PF']
    dfunite6 = dfunite5.values.tolist()
    rowG = 1
    colG = 6
    for g in dfunite6:
        worksheet.write(rowG,colG, g)
        rowG +=1
        
    dfunite7 = dfunite4['Code liste']
    dfunite8 = dfunite7.values.tolist()
    rowH = 1
    colH = 7
    for h in dfunite8:
        worksheet.write(rowH,colH, h)
        rowH +=1
    
    dfunite9 = dfunite4['Appel Nationaux Fixe (en Durée)'].astype(float)
    dfunite10 = dfunite9.values.tolist()
    rowAO = 1
    colAO = 40
    for ao in dfunite10:
        worksheet.write(rowAO,colAO, ao)
        rowAO +=1
    
    dfunite11 = dfunite4['Appel SFR (en Durée)'].astype(float)
    dfunite12 = dfunite11.values.tolist()
    rowAP = 1
    colAP = 41
    for ap in dfunite12:
        worksheet.write(rowAP,colAP, ap)
        rowAP +=1
        
    dfunite13 = dfunite4['Appel Autres opérateurs (en Durée)'].astype(float)
    dfunite14 = dfunite13.values.tolist()
    rowAQ = 1
    colAQ = 42
    for aq in dfunite14:
        worksheet.write(rowAQ,colAQ, aq)
        rowAQ +=1
        
    dfunite15 = dfunite4['Appels Numéro Spéciaux (en Durée)'].astype(float)
    dfunite16 = dfunite15.values.tolist()
    rowAR = 1
    colAR = 43
    for ar in dfunite16:
        worksheet.write(rowAR,colAR, ar)
        rowAR +=1
    
    dfunite17 = dfunite4['Appel reçu à l international (en Durée)'].astype(float)
    dfunite18 = dfunite17.values.tolist()
    rowAS1 = 1
    colAS1 = 44
    for as1 in dfunite18:
        worksheet.write(rowAS1,colAS1, as1)
        rowAS1 +=1
    
    dfunite19 = dfunite4['Appel Emis vers l international depuis la france (en Durée)'].astype(float)
    dfunite20 = dfunite19.values.tolist()
    rowAT = 1
    colAT = 45
    for at in dfunite20:
        worksheet.write(rowAT,colAT, at)
        rowAT +=1
        
    dfunite21 = dfunite4["Appel Emis depuis l'international (en Durée)"].astype(float)
    dfunite22 = dfunite21.values.tolist()
    rowAU = 1
    colAU = 46
    for au in dfunite22:
        worksheet.write(rowAU,colAU, au)
        rowAU +=1
                         
    dfunite23 = dfunite4["Nb de sms (en Durée)"].astype(float)
    dfunite24 = dfunite23.values.tolist()
    rowAV = 1
    colAV = 47
    for av in dfunite24:
        worksheet.write(rowAV,colAV, av)
        rowAV +=1
        
    dfunite25 = dfunite4["Quantité de data en bb (Nat) (en Durée)"].astype(float)
    dfunite26 = dfunite25.values.tolist()
    rowAW = 1
    colAW = 48
    for aw in dfunite26:
        worksheet.write(rowAW,colAW, aw)
        rowAW +=1
    
    dfunite27 = dfunite4["Quantité de data en bb (Inter) (en Durée)"].astype(float)
    dfunite28 = dfunite27.values.tolist()
    rowAX = 1
    colAX = 49
    for ax in dfunite28:
        worksheet.write(rowAX,colAX, ax)
        rowAX +=1

    dfunite29 = dfunite4["Quantité data (Nat) (en Durée)"].astype(float)
    dfunite30 = dfunite29.values.tolist()
    rowAY = 1
    colAY = 50
    for ay in dfunite30:
        worksheet.write(rowAY,colAY, ay)
        rowAY +=1
        
    dfunite31 = dfunite4["Quantité data (Inter) (en Durée)"].astype(float)
    dfunite32 = dfunite31.values.tolist()
    rowAZ = 1
    colAZ = 51
    for az in dfunite32:
        worksheet.write(rowAZ,colAZ, az)
        rowAZ +=1
        
    dfunite33 = dfunite4["Autres Prestation (en Durée)"].astype(float)
    dfunite34 = dfunite33.values.tolist()
    rowBA = 1
    colBA = 52
    for ba in dfunite34:
        worksheet.write(rowBA,colBA, ba)
        rowBA +=1
# A:C, N:S, AA, AC, AJ:AK
    df = dffacture4['Code PF'].values.tolist()
    for a in range(1, len(dffacture4)+1):
        #if (('25186110BS' in df) | ('1121645H1J' in df) | ('244637000P' in df) | ('251861100H' in df)):
        #   worksheet.write(a,0,'CROIX ROUGE FRANCAISE')
        if (('253200400J' in df) | ('2532004001' in df)):
            worksheet.write(a,0,'SOUFFLET')
        elif (('2702269003' in df) | ('2702269009' in df)):
            worksheet.write(a,0,'SILCA')
        else:
            worksheet.write(a,0,'CROIX ROUGE FRANCAISE')
    
    df1 = dfunite4['Code PF'].values.tolist()
    for b in range(1, len(dfunite4)+1):
        if (('253200400J' in df1) | ('2532004001' in df1)):
            worksheet.write(b,1,'SOUFFLET')
        elif (('2702269003' in df1) | ('2702269009' in df1)):
            worksheet.write(b,1,'SILCA')
        else:
            worksheet.write(b,1,'CROIX ROUGE FRANCAISE')
            
# BO
    rowBO = 1
    colBO = 66
    for mm in range(len(dffacture4)):
        if ((dfunite22[mm] + dfunite24[mm] + dfunite16[mm] + dfunite18[mm] + dfunite34[mm] + dfunite32[mm] + dfmtt26[mm] +
            dfmtt24[mm] + dfmtt22[mm] + dfmtt20[mm] + dfmtt10[mm] + dfmtt8[mm] + dfmtt12[mm] + dfmtt18[mm] + dfmtt6[mm] +
            dfunite10[mm] + dfunite24[mm] + dfunite30[mm] + dfunite12[mm] + dfunite14[mm]) == 0):
            worksheet.write(rowBO, colBO, 1)
        else:
            worksheet.write(rowBO,colBO, 0)
        rowAU += 1
        rowAT += 1
        rowAR += 1
        rowAS1 += 1
        rowBA += 1
        rowAZ += 1
        rowAN += 1
        rowAM += 1
        rowAL += 1
        rowAI += 1
        rowAH += 1
        rowAG += 1
        rowAB += 1
        rowAE += 1
        rowAF += 1
        rowAO += 1
        rowAV += 1
        rowAY += 1
        rowAP += 1
        rowAQ += 1
        rowBO += 1

    for c in range(1, len(dffacture4)+1):
        worksheet.write(c,2,'SFR')
        
    for n in range(1, len(dffacture4)+1):
        worksheet.write(n,13, 0)
        
    for o in range(1, len(dffacture4)+1):
        worksheet.write(o,14, 0)
        
    for p in range(1, len(dffacture4)+1):
        worksheet.write(p,15, 0)   

    for q in range(1, len(dffacture4)+1):
        worksheet.write(q,16, 0)
   
    for r in range(1, len(dffacture4)+1):
        worksheet.write(r,17, 0)   

    for s in range(1, len(dffacture4)+1):
        worksheet.write(s,18, 0)
        
    for aa in range(1, len(dffacture4)+1):
        worksheet.write(aa,26, 0)
  
    for ad in range(1, len(dffacture4)+1):
        worksheet.write(ad,29, 0)
        
    for aj in range(1, len(dffacture4)+1):
        worksheet.write(aj,35, 0)
        
    for ak in range(1, len(dffacture4)+1):
        worksheet.write(ak,36, 0)
        
    for aw in range(1, len(dffacture4)+1):
        worksheet.write(aw,48, 0)
        
    for ax in range(1, len(dffacture4)+1):
        worksheet.write(ax,49, 0)
        
    workbook.close()
    
folderPath = StringVar()
a = Label(gui ,text="Sélection du repertoire:")
a.grid(row=0,column = 0)

E = Entry(gui,textvariable=folderPath)
E.grid(row=0,column=1)

btnFind = ttk.Button(gui, text="Rechercher",command=getFolderPath)
btnFind.grid(row=0,column=2)

c = ttk.Button(gui ,text="Classer les données .txt", command=classer)
c.grid(row=1,column=1, pady = 25)
#####
a1 = Label(gui ,text="Lecture des fichiers global:")
a1.grid(row=3,column = 0, pady = 10)

btnfacture = ttk.Button(gui, text="Facture",command=facture)
btnfacture.grid(row=3,column=1, pady = 10)

btnmtt = ttk.Button(gui, text="TypeAppel_Mtt",command=typeappelmtt)
btnmtt.grid(row=4,column=1, pady = 10)

btnunite = ttk.Button(gui, text="TypeAppel_Unite",command=typeappelunite)
btnunite.grid(row=5,column=1, pady = 10)

btnfinal = ttk.Button(gui, text="Génération xls final",command=final)
btnfinal.grid(row=6,column=1, pady = 25)

gui.mainloop()

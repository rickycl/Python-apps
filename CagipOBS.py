from tkinter import *
from tkinter import filedialog
from tkinter.filedialog import askopenfile
import sqlite3
import openpyxl
import pandas as pd
import numpy as np
import xlsxwriter

root = Tk()

root.title("Conso Ca-gip/Schindler")
root.geometry("500x500")
root.minsize(350,450)
root.config(background='light grey')
label_title = Label(root, text="Bienvenue !", font=("Courrier",40), bg='light blue',fg='light yellow')
label_title.pack()

def choosefacdynam():
    global dffac
    file = filedialog.askopenfilename()
    dffac = pd.read_excel(file)  
    
def chooseedpop():
    global dfop
    file = filedialog.askopenfilename()
    dfop = pd.read_excel(file)
    
def chooseedpngc():
    global dfngc
    file = filedialog.askopenfilename()
    dfngc = pd.read_excel(file)

def generer():
    workbook = xlsxwriter.Workbook('20AAMM_OBS_CLIENT.xlsx')
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
    worksheet.write('P1', 'Mail __ Surf', bold)
    worksheet.write('Q1', 'Option SMS', bold)
    worksheet.write('R1', 'OptionSeuil', bold)
    worksheet.write('S1', 'MontantSeuil', bold)
    worksheet.write('T1', 'Date Facture', bold)
    worksheet.write('U1', 'Montant Abo', bold)
    worksheet.write('V1', 'Montant Conso', bold)
    worksheet.write('W1', 'Montant divers', bold)
    worksheet.write('X1', 'Montant HT', bold)
    worksheet.write('Y1', 'Montant TVA', bold)
    worksheet.write('Z1', 'Montant TTC', bold)
    worksheet.write('AA1', 'Remise Comm', bold)
    worksheet.write('AB1', 'Montant Appels Nationaux', bold)
    worksheet.write('AC1', 'Montant Appel Opérateur', bold)
    worksheet.write('AD1', 'Montant Appel Autre Ope', bold)
    worksheet.write('AE1', 'Montant Numéros Spéciaux', bold)
    worksheet.write('AF1', "Montant Appels Reçus à l'International", bold)
    worksheet.write('AG1', 'Montant  Appels France vers International', bold)
    worksheet.write('AH1', 'Montant Appels émis depuis International', bold)
    worksheet.write('AI1', 'Montant SMS_MMS', bold)
    worksheet.write('AJ1', 'Montant Data BB Nat', bold)
    worksheet.write('AK1', 'Montant Data BB Internat', bold)
    worksheet.write('AL1', 'Montant Data national', bold)
    worksheet.write('AM1', 'Montant Data International', bold)
    worksheet.write('AN1', 'Montant Autres Prestations', bold)
    worksheet.write('AO1', 'Durée Appels Nationaux', bold)
    worksheet.write('AP1', 'Unit Appel Opérateur', bold)
    worksheet.write('AQ1', 'Unit Appel Autre Ope', bold)
    worksheet.write('AR1', 'Durée Numéros Spéciaux', bold)
    worksheet.write('AS1', "Durée Appels Reçus à l'International", bold)
    worksheet.write('AT1', 'Durée Appels France vers International', bold)
    worksheet.write('AU1', 'Durée Appels émis depuis International', bold)
    worksheet.write('AV1', 'Quantité SMS_MMS', bold)
    worksheet.write('AW1', 'Volume Data BB Nat', bold)
    worksheet.write('AX1', 'Volume Data BB Internat', bold)
    worksheet.write('AY1', 'Volume Data national', bold)
    worksheet.write('AZ1', 'Volume Data International', bold)
    worksheet.write('BA1', 'Quantité Autres Prestations', bold)
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

    dffac1 = dffac.rename(columns={'Numéro accès':'Ligne'})
    dffac2 = pd.pivot_table(dffac1, values = ['Quantité (Acte)','Quantité (Volume (Mo))','Quantité (Durée (hh:mm:ss))',
                                              'Quantité (Volume (Ko))','Quantité (Hors conso)','Montant (€ HT)','Montant TVA',
                                              'Montant (€ TTC)'], index = ['Ligne'], aggfunc = np.sum)
    dffac777 = dffac1.drop(['Code groupe','Code entreprise','Numéro compte','Montant (€ HT)','Montant (€ TTC)',
                        'Montant TVA','Prix Unitaire','Quantité (Acte)','Raison Sociale','Votre référence',
                        'Quantité (Durée (hh:mm:ss))','Quantité (Hors conso)','Quantité (Volume (Ko))','Numéro facture',
                        'Quantité (Volume (Mo))','Libellé ligne facture','affectation','Type de charge',
                        'Siren','Taux TVA'], axis=1)
    dffac7777 = dffac777.drop_duplicates('Ligne')
    dffac77777a = dffac7777['Date']
    dffac77777 = dffac77777a.tolist()
    rowT = 1
    colT = 19
    for t in dffac77777:
        worksheet.write(rowT,colT, t)
        rowT += 1
    
# EDP OPERATEUR
    dfop2 = dfop.rename(columns={"No d'appel":'Ligne'})
    dfop2['Cle PUK'].replace('', np.nan, inplace=True)
    dfop2.dropna(subset=['Cle PUK'], inplace=True)

    list1 = dfop2['Ligne'].tolist()
    list2 = dfop2['Cle PUK'].tolist()
    list3 = dfop2['No Abonne'].tolist()
    list4 = dfop2["Date de fin d'engagement"].tolist()
    list5 = dfop2['Types Abonnements'].tolist()
    list6 = dfop2['No carte SIM'].tolist()

    data1 = {"Ligne": list1,
             'Cle PUK': list2,
             'No Abonne': list3,
             "Date de fin d'engagement": list4,
             'Types Abonnements': list5,
             'No carte SIM': list6}

    dfop3 = pd.DataFrame(data1, columns = ['Ligne','Cle PUK','No Abonne',"Date de fin d'engagement",'Types Abonnements',
                                      'No carte SIM'])
    
    join = pd.merge(dffac2, dfop3, on='Ligne',how='left')
    join100 = join.fillna('')
# E
    join1001 = join100['No Abonne']
    join1002 = join1001.values.tolist()
    rowE = 1
    colE = 4
    for e in join1002:
        worksheet.write(rowE,colE, e)
        rowE +=1
# F, K, M, BE, BH --------------------------------------- 6
    join101 = join100['Ligne']
    join102 = join101.values.tolist()
    rowF = 1
    colF = 5
    for f in join102:
        worksheet.write(rowF,colF, f)
        rowF +=1
        
    join105 = join100["Date de fin d'engagement"]
    join106 = join105.values.tolist()
    rowK = 1
    colK = 10
    for k in join106:
        worksheet.write(rowK,colK, k)
        rowK +=1
    
    join107 = join100['Types Abonnements']
    join108 = join107.values.tolist()
    rowM = 1
    colM = 12
    for m in join108:
        worksheet.write(rowM,colM, m)
        rowM +=1
     
    rowL = 1
    colL = 11
    for m in range(len(join108)):
        if ((join108[m] == 'BE Smartphone Connect 2016') | (join108[m] == 'BE Abondance 2014') |
        (join108[m] == 'Business Everywhere illimité 3') | (join108[m] == 'Business Tablette Initial')):
            worksheet.write(rowL,colL, 'DATA')
        elif (join108[m] == ''):
            worksheet.write(rowL,colL, '')
        else:
            worksheet.write(rowL,colL, 'VOIX')
        rowM += 1
        rowL += 1

    join109 = join100['No carte SIM']
    join110 = join109.values.tolist()
    rowBE = 1
    colBE = 56
    for be in join110:
        worksheet.write(rowBE,colBE, be)
        rowBE +=1

    join111 = join100['Cle PUK']
    join112 = join111.values.tolist()
    rowBH = 1
    colBH = 59
    for bh in join112:
        worksheet.write(rowBH,colBH, bh)
        rowBH +=1        
# A:C, N:S, AA, AC, AJ:AK, AP:AQ, AW:AX -------------- 16
    for a in range(1, len(dffac2)+1):
        worksheet.write(a,0,'CA-GIP')
        
    for b in range(1, len(dffac2)+1):
        worksheet.write(b,1,'CA-GIP')

    for c in range(1, len(dffac2)+1):
        worksheet.write(c,2,'ORANGE')
        
    for n in range(1, len(dffac2)+1):
        worksheet.write(n,13, 0)
        
    for o in range(1, len(dffac2)+1):
        worksheet.write(o,14, 0)
        
    for p in range(1, len(dffac2)+1):
        worksheet.write(p,15, 0)   

    for q in range(1, len(dffac2)+1):
        worksheet.write(q,16, 0)
   
    for r in range(1, len(dffac2)+1):
        worksheet.write(r,17, 0)   

    for s in range(1, len(dffac2)+1):
        worksheet.write(s,18, 0)
        
    for aa in range(1, len(dffac2)+1):
        worksheet.write(aa,26, 0)
        
    for ac in range(1, len(dffac2)+1):
        worksheet.write(ac,28, 0)
        
    for ad in range(1, len(dffac2)+1):
        worksheet.write(ad,29, 0)
        
    for aj in range(1, len(dffac2)+1):
        worksheet.write(aj,35, 0)
        
    for ak in range(1, len(dffac2)+1):
        worksheet.write(ak,36, 0)
    
    for ap in range(1, len(dffac2)+1):
        worksheet.write(ap,41, 0)
        
    for aq in range(1, len(dffac2)+1):
        worksheet.write(aq,42, 0 )
        
    for aw in range(1, len(dffac2)+1):
        worksheet.write(aw,48, 0)
        
    for ax in range(1, len(dffac2)+1):
        worksheet.write(ax,49, 0)
# AO AV AY U V W --------------------------------------- 6
    dffac20 = dffac2['Quantité (Durée (hh:mm:ss))']
    dffac21 = dffac20.values.tolist()
    rowAO = 1
    colAO = 40
    for ao in dffac21:
        worksheet.write(rowAO,colAO, ao)
        rowAO +=1
    
    dffac22 = dffac2['Quantité (Acte)']
    dffac23 = dffac22.values.tolist()
    rowAV = 1
    colAV = 47
    for av in dffac23:
        worksheet.write(rowAV,colAV, av)
        rowAV +=1
    
    dffac24 = dffac2['Quantité (Volume (Mo))']
    dffac25 = dffac24.values.tolist()
    rowAY = 1
    colAY = 50
    for ay in dffac25:
        worksheet.write(rowAY,colAY, ay)
        rowAY +=1
# U:W ------------------------------------------- 3
    dffac26a = dffac1[dffac1['Type de charge']=="Abonnements, forfaits, formules et options"]
    dffac27a = dffac1[dffac1['Type de charge']=="Consommations (hors forfaits)"]
    dffac28a = dffac1[dffac1['Type de charge']=="Services ponctuels"]
    dffac29a = pd.pivot_table(dffac26a, values = ['Montant (€ HT)'], index = ["Ligne"], aggfunc = np.sum)
    dffac30a = pd.pivot_table(dffac27a, values = ['Montant (€ HT)'], index = ["Ligne"], aggfunc = np.sum)
    dffac31a = pd.pivot_table(dffac28a, values = ['Montant (€ HT)'], index = ["Ligne"], aggfunc = np.sum)
    
    join150re = dffac29a.rename(columns={'Montant (€ HT)':'MT ABO'})
    join150 = pd.merge(dffac2, join150re, on='Ligne',how='left')
    join151 = join150.fillna('')
    
    join152re = dffac30a.rename(columns={'Montant (€ HT)':'MT CON'}) 
    join152 = pd.merge(dffac2, join152re, on='Ligne',how='left')  
    join153 = join152.fillna(0)
    
    join154re = dffac31a.rename(columns={'Montant (€ HT)':'MT DIVERS'})
    join154 = pd.merge(dffac2, join154re, on='Ligne',how='left')
    join155 = join154.fillna(0)
    
    dffac31 = join151['MT ABO']
    dffac32 = dffac31.values.tolist()
    rowU = 1
    colU = 20
    for u in dffac32:
        worksheet.write(rowU,colU, u)
        rowU +=1
        
    dffac33 = join153['MT CON']
    dffac34 = dffac33.values.tolist()
    rowV = 1
    colV = 21
    for v in dffac34:
        worksheet.write(rowV,colV, v)
        rowV +=1
    
    dffac35 = join155['MT DIVERS']
    dffac36 = dffac35.values.tolist()
    rowW = 1
    colW = 22
    for w in dffac36:
        worksheet.write(rowW,colW, w)
        rowW +=1
        
# -----------------------------------------------------------------------------------
    dffac80 = dffac1[dffac1['affectation']=="Appels émis depuis l'international"]
    dffac170 = pd.pivot_table(dffac80, values = ['Quantité (Durée (hh:mm:ss))', 'Montant (€ HT)'], index = ["Ligne"], aggfunc = np.sum)
    dffac17ren = dffac170.rename(columns={'Quantité (Durée (hh:mm:ss))': 'AEDI QD', 'Montant (€ HT)': 'AEDI MTHT'})

    dffac90 = dffac1[dffac1['affectation']=="Appels émis vers international depuis la France"]
    dffac180 = pd.pivot_table(dffac90, values = ['Quantité (Durée (hh:mm:ss))', 'Montant (€ HT)'], index = ["Ligne"], aggfunc = np.sum)
    dffac18ren = dffac180.rename(columns={'Quantité (Durée (hh:mm:ss))': 'AEVIDF QD', 'Montant (€ HT)': 'AEVIDF MTHT' })

    dffac100 = dffac1[dffac1['affectation']=="Appels nationaux"]
    dffac190 = pd.pivot_table(dffac100, values = ['Quantité (Durée (hh:mm:ss))', 'Montant (€ HT)'], index = ["Ligne"], aggfunc = np.sum)
    dffac19ren = dffac190.rename(columns={'Quantité (Durée (hh:mm:ss))': 'Ap Nat QD', 'Montant (€ HT)': 'Ap Nat MTHT'})

    dffac110 = dffac1[dffac1['affectation']=="Appels Numéros spéciaux"]
    dffac200 = pd.pivot_table(dffac110, values = ['Quantité (Durée (hh:mm:ss))', 'Montant (€ HT)'], index = ["Ligne"], aggfunc = np.sum)
    dffac20ren = dffac200.rename(columns={'Quantité (Durée (hh:mm:ss))': 'ANS QD', 'Montant (€ HT)': 'ANS MTHT'})

    dffac120 = dffac1[dffac1['affectation']=="Appels reçu à l'international"]
    dffac210 = pd.pivot_table(dffac120, values = ['Quantité (Durée (hh:mm:ss))', 'Montant (€ HT)'], index = ["Ligne"], aggfunc = np.sum)
    dffac21ren = dffac210.rename(columns={'Quantité (Durée (hh:mm:ss))': 'ARAI QD', 'Montant (€ HT)': 'ARAI MTHT'})
# ------------------------------------------------------------------------------------------------------------------------    
    dffac130 = dffac1[dffac1['affectation']=="Autres prestations"]
    dffac220 = pd.pivot_table(dffac130, values = ['Quantité (Acte)', 'Montant (€ HT)'], index = ["Ligne"], aggfunc = np.sum)
    dffac22ren = dffac220.rename(columns={'Quantité (Acte)': 'Aut Pres QA', 'Montant (€ HT)': 'Aut Pres MTHT'})

    dffac140 = dffac1[dffac1['affectation']=="Data international"]
    dffac230 = pd.pivot_table(dffac140, values = ['Quantité (Volume (Mo))', 'Montant (€ HT)'], index = ["Ligne"], aggfunc = np.sum)
    dffac23ren = dffac230.rename(columns={'Quantité (Volume (Mo))': 'Dat Int QVm', 'Montant (€ HT)': 'Dat Int MTHT'})
    
    dffac150 = dffac1[dffac1['affectation']=="Data national"]
    dffac240 = pd.pivot_table(dffac150, values = ['Quantité (Volume (Mo))', 'Montant (€ HT)'], index = ["Ligne"], aggfunc = np.sum)
    dffac24ren = dffac240.rename(columns={'Quantité (Volume (Mo))': 'Dat Nat QVm', 'Montant (€ HT)': 'Dat Nat MTHT'})
    
    dffac160 = dffac1[dffac1['affectation']=="SMS_MMS"]
    dffac250 = pd.pivot_table(dffac160, values = ['Quantité (Acte)', 'Montant (€ HT)'], index = ["Ligne"], aggfunc = np.sum)
    dffac25ren = dffac220.rename(columns={'Quantité (Acte)': 'SMS_MMS QA', 'Montant (€ HT)': 'SMS_MMS MTHT'})
    
    join10 = pd.merge(dffac2, dffac17ren, on='Ligne',how='left')
    join11 = pd.merge(join10, dffac18ren, on='Ligne',how='left')
    join12 = pd.merge(join11, dffac19ren, on='Ligne',how='left')
    join13 = pd.merge(join12, dffac20ren, on='Ligne',how='left')
    join14 = pd.merge(join13, dffac21ren, on='Ligne',how='left')
    join15 = pd.merge(join14, dffac22ren, on='Ligne',how='left')
    join16 = pd.merge(join15, dffac23ren, on='Ligne',how='left')
    join17 = pd.merge(join16, dffac24ren, on='Ligne',how='left')
    join18 = pd.merge(join17, dffac25ren, on='Ligne',how='left')
    join19 = join18.fillna(0)
    
# AB, AE:AF, AV:BA --------------------------------------------18
    join20 = join19['AEDI QD']
    join21 = join20.values.tolist()
    rowAU = 1
    colAU = 46
    for au in join21:
        worksheet.write(rowAU,colAU, au)
        rowAU +=1
    
    join22 = join19['AEVIDF QD']
    join23 = join22.values.tolist()
    rowAT = 1
    colAT = 45
    for at in join23:
        worksheet.write(rowAT,colAT, at)
        rowAT +=1

    join26 = join19['ANS QD']
    join27 = join26.values.tolist()
    rowAR = 1
    colAR = 43
    for ar in join27:
        worksheet.write(rowAR,colAR, ar)
        rowAR +=1
        
    join28 = join19['ARAI QD']
    join29 = join28.values.tolist()
    rowAS = 1
    colAS = 44
    for as1 in join29:
        worksheet.write(rowAS,colAS, as1)
        rowAS +=1       
        
    join30 = join19['Aut Pres QA']
    join31 = join30.values.tolist()
    rowBA = 1
    colBA = 52
    for ba in join31:
        worksheet.write(rowBA,colBA, ba)
        rowBA +=1

    join32 = join19['Dat Int QVm']
    join33 = join32.values.tolist()
    rowAZ = 1
    colAZ = 51
    for az in join33:
        worksheet.write(rowAZ,colAZ, az)
        rowAZ +=1
# AB, AE:AH, AI, AL:AM ------------------ MONTANT HORS TAXE
    join38 = join19['Aut Pres MTHT']
    join39 = join38.values.tolist()
    rowAN = 1
    colAN = 39
    for an in join39:
        worksheet.write(rowAN,colAN, an)
        rowAN +=1
        
    join40 = join19['Dat Int MTHT']
    join41 = join40.values.tolist()
    rowAM = 1
    colAM = 38
    for am in join41:
        worksheet.write(rowAM,colAM, am)
        rowAM +=1
    
    join42 = join19['Dat Nat MTHT']
    join43 = join42.values.tolist()
    rowAL = 1
    colAL = 37
    for al in join43:
        worksheet.write(rowAL,colAL, al)
        rowAL +=1    
    
    join44 = join19['SMS_MMS MTHT']
    join45 = join44.values.tolist()
    rowAI = 1
    colAI = 34
    for ai in join45:
        worksheet.write(rowAI,colAI, ai)
        rowAI +=1        
        
    join46 = join19['AEDI MTHT']
    join47 = join46.values.tolist()
    rowAH = 1
    colAH = 33
    for ah in join47:
        worksheet.write(rowAH,colAH, ah)
        rowAH +=1
    
    join48 = join19['AEVIDF MTHT']
    join49 = join48.values.tolist()
    rowAG = 1
    colAG = 32
    for ag in join49:
        worksheet.write(rowAG,colAG, ag)
        rowAG +=1
    
    join50 = join19['Ap Nat MTHT']
    join51 = join50.values.tolist()
    rowAB = 1
    colAB = 27
    for ab in join51:
        worksheet.write(rowAB,colAB, ab)
        rowAB +=1
        
    join52 = join19['ANS MTHT']
    join53 = join52.values.tolist()
    rowAE = 1
    colAE = 30
    for ae in join53:
        worksheet.write(rowAE,colAE, ae)
        rowAE +=1
        
    join54 = join19['ARAI MTHT']
    join55 = join54.values.tolist()
    rowAF = 1
    colAF = 31
    for af in join55:
        worksheet.write(rowAF,colAF, af)
        rowAF +=1
    
# D G X Y Z BN -------------------------------------------6
    dffac11 = pd.pivot_table(dffac1, index = ['Ligne'])
    dffac12 = dffac11.drop(['Code groupe','Montant (€ HT)','Montant (€ TTC)','Montant TVA','Prix Unitaire','Quantité (Acte)',
                   'Quantité (Durée (hh:mm:ss))','Quantité (Hors conso)','Quantité (Volume (Ko))','Quantité (Volume (Mo))',
                   'Siren','Taux TVA'], axis=1)
    join = pd.merge(dffac2, dffac12, on='Ligne',how='left')
    
    join1 = join['Numéro compte']
    join3 = join1.values.tolist()
    rowG = 1
    colG = 6
    for g in join3:
        worksheet.write(rowG,colG, g)
        rowG +=1    
        
    rowD = 1
    colD = 3
    for d in join3:
        worksheet.write(rowD,colD, d)
        rowD +=1
        
    join2 = join['Numéro facture']
    join4 = join2.values.tolist()
    rowBN = 1
    colBN = 65
    for bn in join4:
        worksheet.write(rowBN,colBN, bn)
        rowBN +=1

    dffac5 = dffac2['Montant (€ HT)']
    dffac6 = dffac2['Montant TVA']
    dffac7 = dffac2['Montant (€ TTC)']
    
    dffac8 = dffac5.values.tolist()
    rowX = 1
    colX = 23
    for x in dffac8:
        worksheet.write(rowX,colX, x)
        rowX +=1
    
    dffac9 = dffac6.values.tolist()
    rowY = 1
    colY = 24
    for y in dffac9:
        worksheet.write(rowY,colY, y)
        rowY +=1
    
    dffac10 = dffac7.values.tolist()
    rowZ = 1
    colZ = 25
    for z in dffac10:
        worksheet.write(rowZ,colZ, z)
        rowZ +=1
        
# Sans Conso BO
    rowBO = 1
    colBO = 66
    for mm in range(len(dffac2)):
        if ((join21[mm] + join23[mm] + join27[mm] + join29[mm] + join31[mm] + join33[mm] + join39[mm] + join41[mm] +
            join43[mm] + join45[mm] + join47[mm] + join49[mm] + join51[mm] + join53[mm] + join55[mm] +
            dffac21[mm] + dffac23[mm] + dffac25[mm]) == 0):
            worksheet.write(rowBO, colBO, 1)
        else:
            worksheet.write(rowBO,colBO, 0)
        rowAU += 1
        rowAT += 1
        rowAR += 1
        rowAS += 1
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
        rowBO += 1
                           
# BB:BD, BJ -------------------------------------------------------------- EDP NGC
    dfngc1 = dfngc.rename(columns={'Numéro accès':'Ligne'})
    
    listA = dfngc1['Ligne'].tolist()
    listB = dfngc1['Nom Utilisateur'].tolist()
    listC = dfngc1['Matricule'].tolist()
    listD = dfngc1['Code1'].tolist()
    listE = dfngc1['IMEI'].tolist()
    listF = dfngc1['Etat'].tolist()
    
    data = {"Ligne": listA,
             'Nom Utilisateur': listB,
             'Matricule': listC,
             'Code1': listD,
             'IMEI': listE,
             'Etat': listF}

    dfngc2 = pd.DataFrame(data, columns = ['Ligne','Nom Utilisateur','Matricule','Code1','IMEI','Etat']) 
    dfngc3 = dfngc2[dfngc2['Etat'] == 'ACTIF']
    join113 = pd.merge(dffac2, dfngc3, on='Ligne',how='left')
    join114 = join113.fillna('')

    join115 = join114['Nom Utilisateur']
    join116 = join115.values.tolist()
    rowBB = 1
    colBB = 53
    for bb in join116:
        worksheet.write(rowBB,colBB, bb)
        rowBB +=1

    join117= join114['Matricule']
    join118= join117.values.tolist()
    rowBC= 1
    colBC= 54
    for bc in join118:
        worksheet.write(rowBC,colBC, bc)
        rowBC +=1
    
    join119 = join114['Code1']
    join120 = join119.values.tolist()
    rowBD = 1
    colBD = 55
    for bd in join120:
        worksheet.write(rowBD,colBD, bd)
        rowBD +=1
        
    join121 = join114['IMEI']
    join122 = join121.values.tolist()
    rowBJ = 1
    colBJ = 61
    for bj in join122:
        worksheet.write(rowBJ,colBJ, bj)
        rowBJ +=1
        
    workbook.close()  
    
myButtonfacdynam = Button(root,text="Lecture du fichier Facture Dynamique",command=choosefacdynam)
myButtonfacdynam.pack(pady=15)

myButtonedpop = Button(root,text="Lecture du fichier EDP Opérateur",command=chooseedpop)
myButtonedpop.pack(pady=15)

myButtonedpngc = Button(root,text="Lecture du fichier EDP NGC",command=chooseedpngc)
myButtonedpngc.pack(pady=15)

myButtongen = Button(root,text="Créer le fichier d'import",command=generer)
myButtongen.pack(pady=50)

root.mainloop()

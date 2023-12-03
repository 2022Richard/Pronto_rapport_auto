from django.shortcuts import render
import pandas as pd
from datetime import datetime
from django.http import HttpResponse
from .models import Dispatch_Engin
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
import locale
from openpyxl.styles import Font



def accueil(request): 
    return render(request, 'accueil.html', {'etat':'accueil'})

def contact(request): 
    return render(request, 'contact.html', {'etat':'contact'})


def rapport(request): 
    message = ''

    if request.method == "POST" : 
        
        file = request.FILES['files']

        mon_fichier, test = str(file), False
        type_file = mon_fichier.endswith("xlsx")

        if type_file == False : 
            type_file = mon_fichier.endswith("xls") 

        if type_file == False : 
            message = "Ce fichier n'est pas au format excel. Un fichier excel a pour extension xlsx ou xlx. "
        else : 
            ma_base = pd.read_excel(file)  
            
            if not (("Notepad" in ma_base.columns) or ("Xmit" in ma_base.columns) or ("Despatch" in ma_base.columns) or \
                    ("Vehicle" in ma_base.columns) or ("Arrived" in ma_base.columns) or ("Signal Time" in ma_base.columns) \
                        or ("Sig-Arr Time" in ma_base.columns) or ("Des-ArrTime" in ma_base.columns)) : 
                     
                     message = "Ce fichier est au format excel mais pas celui des declenchements alarme."  
                  
            else :       
                    if not(("custdesc" in ma_base.columns) and ("PhysicalAddress" in ma_base.columns)) :
                        message = "Ce fichier est celui des declenchements alarme. \
                               Mais veuillez ajouter les colonnes 'custdesc' et/ou 'PhysicalAddress' puis reessayez !" 
                        
                    else:  test = True
    
            if test == True : 
                # print("to be continued")

                if 'alarme' in request.POST :  
                    resultat = rapport_alarme(ma_base)  
                    resultat_fin = pd.DataFrame(resultat)

                    #print(resultat.columns)    
                    
                    response = HttpResponse(content_type='application/ms-excel')
                    response['Content-Disposition'] = 'attachment; filename="Rapport_alarme _ ' + str(date_en_francais()) + ' _ ' + \
                        str(datetime.now().strftime('%H')) + 'h ' + str(datetime.now().strftime('%M'))+ 'mn ' + str(datetime.now().strftime('%S')) + 's ' +'.xlsx"'

                    wb = openpyxl.Workbook()
                    ws = wb.active
                    ws.title = "Rapport Alarme"

                    # Add data from the model  datetime.datetime.now().strftime('%H')+ 'H'
                    for r in dataframe_to_rows(resultat_fin, index=False, header=True):
                        ws.append(r)    
                    
                    # Save the workbook to the HttpResponse  
                    wb.save(response)
                    return response

                if 'open' in request.POST :
                    resultat_open_close = rapport_open_close(ma_base) 
                    open_close = pd.DataFrame(resultat_open_close)

                    #print(resultat_open_close)

                    response = HttpResponse(content_type='application/ms-excel')
                    response['Content-Disposition'] = 'attachment; filename="Rapport_open_close ' + str(date_en_francais()) + ' _ ' + \
                        str(datetime.now().strftime('%H')) + 'h ' + str(datetime.now().strftime('%M'))+ 'mn ' + str(datetime.now().strftime('%S')) + 's ' +'.xlsx"'

                    wb = openpyxl.Workbook()
                    ws = wb.active
                    ws.title = "Rapport Open Close"

                    # Add data from the model
                    for r in dataframe_to_rows(open_close, index=False, header=True):
                        ws.append(r)    

                    # Save the workbook to the HttpResponse
                    wb.save(response)
                    return response
                    
            
    return render(request, 'rapport.html', {'etat':'rapport', 'message': message})


def rapport_alarme (ma_base) :
                
    ma_base['filtre'] = ma_base.Notepad.str[9:14] 
    ma_base['filtre'] = ma_base['filtre'].str.replace('ALARM','ALARME')

    def ma_selection(ma_base) : 
        if (ma_base['filtre'].find("PANIC") == 0) or (ma_base['filtre'].find("ALARM") == 0) : 
            return "ok" 
        else : 
            return "no"

    ma_base["test"] = ma_base.apply(ma_selection, axis=1)

    selection = ma_base.loc[(ma_base['test'] == 'ok')]
    ma_base = selection
	
    ### Je selectionne uniquement les alarmes (Type : alarme et panic) ici : 
                
    ####################### DEBUT TRAITEMENT DE LA BASE SELECTION OU MA_BASE Vehicle
                               
    #OPERATEUR = selection["Operator"] 
    TYPE_DECLENCHEMENT = selection["filtre"]
    CODE_ALARME = selection["Xmit"]
    COMMENTAIRE = selection["Notepad"]
    EQUIPAGE = selection["Vehicle"]            
    #TEMPS_PRONTO 

    date_et_heure = list(map(lambda x: x.partition(".")[0], selection["Signal Time"].astype(str)))
    #date_et_heure

    DATE = list(map(lambda x: datetime.strptime(x,'%Y-%m-%d %H:%M:%S').date().strftime('%d-%m-%Y'), date_et_heure))
    H_RECEPT = list(map(lambda x: datetime.strptime(x,'%Y-%m-%d %H:%M:%S').time(), date_et_heure))

    # FONCTION HEURE 
    def heure(value_1) :
        if value_1 == 'NaT' : return "" 
        else : return datetime.strptime(str(value_1), '%Y-%m-%d %H:%M:%S').time()

    # HEURE DE DESPATCH 
    desp = list(map(lambda x: x.partition(".")[0], selection["Despatch"].astype(str))) 
    db_desp = pd.DataFrame(desp, columns=["Desp"],)
    H_DESPATCH = db_desp.apply(lambda x: heure(x["Desp"]), axis=1)

    # HEURE D'ARRIVEE 
    arr = list(map(lambda x: x.partition(".")[0], selection["Arrived"].astype(str)))
    db_arr = pd.DataFrame(arr, columns=["Arr"],)
    H_ARRIVEE = db_arr.apply(lambda x: heure(x["Arr"]), axis=1)
            
    # FONCTION TEMPS : PATRL + INTERVENTION 
    def temps(val) : 
        val = str(val) 
        if val == 'nan' : return ""
        else : return val
                
    # Temps d'intervention (Heure d'arrivée de la patrouille - Heure de reception du signal)
    inter,  = selection["Sig-Arr Time"],
    db_inter,  = pd.DataFrame(list(inter), columns=["inter"],), 
    TEMPS_INTER,  = db_inter.apply(lambda x: temps(x["inter"]), axis=1),
                
    # Temps Patrouille : temps mis par la patrouille pour arriver chez le client dès que Pronto leur a donné l'alerte 
    patrl = selection["Des-ArrTime"]
    db_patrl = pd.DataFrame(list(patrl), columns=["patrl"],)
    TEMPS_PATRL = db_patrl.apply(lambda x: temps(x["patrl"]), axis=1)

    #  Temps mis par pronto pour lancer un équipage dès la reception du signal 
    ### FONCTION PRONTO : Je calcule le temps de PRONTO
    def pronto(val_1, val_2) :
        val_2, val_1 = str(val_2), str(val_1)
        if (val_1 =="") or (val_2 =="") or (val_1 =="NaT") or (val_2 =="NaT") :
            return ""
        else :
            format = "%H:%M:%S"
            val_2 = datetime.strptime(val_2, format)
            val_1 = datetime.strptime(val_1, format)
            val = val_2 - val_1
            val = datetime.strptime(str(val), format).time()
            return val

    # Temps Pronto : temps mis par Pronto pour lancer la patrouille dès la reception du code
    db_pronto = pd.DataFrame(list(zip(TEMPS_PATRL, TEMPS_INTER)), columns=["patrl", "inter"],)
    TEMPS_PRONTO = db_pronto.apply(lambda x: pronto(x["patrl"], x["inter"]), axis=1)
                
    ## Ajouter l'équipage après le code de l'alarme : 
    equipe = pd.DataFrame(list(EQUIPAGE), columns=["Area"]) 
    equipe_test = equipe['Area'].astype(str)
    equipe_test = equipe_test.str.replace("nan", "Richard")
    equipe = pd.DataFrame(list(equipe_test), columns=["Area"])
  
    colonne = [{"Area": x.Area, "Vehicule": x.Vehicule, "Description": x.Description} for x in Dispatch_Engin.objects.all()]
    mes_equipe = pd.DataFrame(colonne)

    fusion_equipe = equipe.merge(mes_equipe, how='left', on='Area')

    fusion_equipe_fin = fusion_equipe['Description'].astype(str)
    fusion_equipe_fin = fusion_equipe_fin.str.replace("nan", "")

    ############################################## RESULTAT ET AFFICHAGE  ###################################################
    resultat = pd.DataFrame(list(zip(TYPE_DECLENCHEMENT, DATE, H_RECEPT, H_DESPATCH, H_ARRIVEE, CODE_ALARME, fusion_equipe_fin,    
                                            selection["custdesc"], selection["PhysicalAddress"],COMMENTAIRE, TEMPS_INTER,  
                                            TEMPS_PATRL, TEMPS_PRONTO)),

                                    columns=["TYPE DECLEN.", "DATE", "H.RECEPT", "H.DESPATCH", "H.ARRIVEE", "CODE","EQUIPAGE", 
                                             "NOM CLIENT", "ADRESSE CLIENT", "COMMENTAIRE", "TEMPS INTER.",  "TEMPS PATRL", "TEMPS PRONTO"])
    return resultat


###################################################### RAPPORT OPEN / CLOSE ##############################################
# OPERATEUR, TYPE DECLENCHEMENT, DATE
def rapport_open_close (ma_base) :
                
    ma_base['filtre'] = ma_base.Notepad.str[9:14] 
    ma_base['filtre'] = ma_base['filtre'].str.replace('OPENI','OPENING')
    ma_base['filtre'] = ma_base['filtre'].str.replace('CLOSI','CLOSING')

    def ma_selection(ma_base) : 
        if (ma_base['filtre'].find("OPEN") == 0) or (ma_base['filtre'].find("CLOSE") == 0) or \
            (ma_base['filtre'].find("OPENING") == 0) or (ma_base['filtre'].find("CLOSING") == 0) : 
            return "ok" 
        else : 
            return "no"

    ma_base["test"] = ma_base.apply(ma_selection, axis=1)

    selection = ma_base.loc[(ma_base['test'] == 'ok')]
    ma_base = selection
	
    ### Je selectionne uniquement les alarmes (Type : Open et Close) ici : 
                
    ####################### DEBUT TRAITEMENT DE LA BASE SELECTION OU MA_BASE Vehicle
                               
    #OPERATEUR = selection["Operator"] 
    TYPE_DECLENCHEMENT = selection["filtre"]
    CODE_ALARME = selection["Xmit"]
    COMMENTAIRE = selection["Notepad"]
    EQUIPAGE = selection["Vehicle"]            
    #TEMPS_PRONTO 

    date_et_heure = list(map(lambda x: x.partition(".")[0], selection["Signal Time"].astype(str)))
    #date_et_heure

    DATE = list(map(lambda x: datetime.strptime(x,'%Y-%m-%d %H:%M:%S').date().strftime('%d-%m-%Y'), date_et_heure))
    H_RECEPT = list(map(lambda x: datetime.strptime(x,'%Y-%m-%d %H:%M:%S').time(), date_et_heure))

    # FONCTION HEURE 
    def heure(value_1) :
        if value_1 == 'NaT' : return "" 
        else : return datetime.strptime(str(value_1), '%Y-%m-%d %H:%M:%S').time()

    # HEURE DE DESPATCH 
    desp = list(map(lambda x: x.partition(".")[0], selection["Despatch"].astype(str))) 
    db_desp = pd.DataFrame(desp, columns=["Desp"],)
    H_DESPATCH = db_desp.apply(lambda x: heure(x["Desp"]), axis=1)

    # HEURE D'ARRIVEE 
    arr = list(map(lambda x: x.partition(".")[0], selection["Arrived"].astype(str)))
    db_arr = pd.DataFrame(arr, columns=["Arr"],)
    H_ARRIVEE = db_arr.apply(lambda x: heure(x["Arr"]), axis=1)
            
    # FONCTION TEMPS : PATRL + INTERVENTION 
    def temps(val) : 
        val = str(val) 
        if val == 'nan' : return ""
        else : return val
                
    # Temps d'intervention (Heure d'arrivée de la patrouille - Heure de reception du signal)
    inter,  = selection["Sig-Arr Time"],
    db_inter,  = pd.DataFrame(list(inter), columns=["inter"],), 
    TEMPS_INTER,  = db_inter.apply(lambda x: temps(x["inter"]), axis=1),
                
    # Temps Patrouille : temps mis par la patrouille pour arriver chez le client dès la reception de l'alerte par Pronto 
    patrl = selection["Des-ArrTime"]
    db_patrl = pd.DataFrame(list(patrl), columns=["patrl"],)
    TEMPS_PATRL = db_patrl.apply(lambda x: temps(x["patrl"]), axis=1)

    #  Temps Pronto : Temps mis par pronto pour lancer un équipage dès la reception du signal

    ### FONCTION PRONTO : Je calcule le temps de PRONTO
    def pronto(val_1, val_2) :
        val_2, val_1 = str(val_2), str(val_1)
        if (val_1 =="") or (val_2 =="") or (val_1 =="NaT") or (val_2 =="NaT") :
            return ""
        else :
            format = "%H:%M:%S"
            val_2 = datetime.strptime(val_2, format)
            val_1 = datetime.strptime(val_1, format)
            val = val_2 - val_1
            val = datetime.strptime(str(val), format).time()
            return val

    # Temps Pronto : Temps mis par pronto pour lancer un équipage dès la reception du signal
    db_pronto = pd.DataFrame(list(zip(TEMPS_PATRL, TEMPS_INTER)), columns=["patrl", "inter"],)
    TEMPS_PRONTO = db_pronto.apply(lambda x: pronto(x["patrl"], x["inter"]), axis=1)
                
    ## Ajouter l'équipage qui a pris le code : 
    equipe = pd.DataFrame(list(EQUIPAGE), columns=["Area"]) 
    equipe_test = equipe['Area'].astype(str)
    equipe_test = equipe_test.str.replace("nan", "Richard")
    equipe = pd.DataFrame(list(equipe_test), columns=["Area"])

    colonne = [{"Area": x.Area, "Vehicule": x.Vehicule, "Description": x.Description} for x in Dispatch_Engin.objects.all()]
    mes_equipe = pd.DataFrame(colonne)

    fusion_equipe = equipe.merge(mes_equipe, how='left', on='Area')

    fusion_equipe_fin = fusion_equipe['Description'].astype(str)
    fusion_equipe_fin = fusion_equipe_fin.str.replace("nan", "")

    ############################################## RESULTAT ET AFFICHAGE  ###################################################
    resultat = pd.DataFrame(list(zip(TYPE_DECLENCHEMENT, DATE, H_RECEPT, H_DESPATCH, H_ARRIVEE, CODE_ALARME, fusion_equipe_fin,    
                                           selection["custdesc"], selection["PhysicalAddress"], COMMENTAIRE, TEMPS_INTER,  
                                           TEMPS_PATRL, TEMPS_PRONTO)),

                                    columns=["TYPE DECLEN.", "DATE", "H.RECEPT", "H.DESPATCH", "H.ARRIVEE", "CODE","EQUIPAGE", 
                                             "NOM CLIENT", "ADRESSE CLIENT", "COMMENTAIRE", "TEMPS INTER.",  "TEMPS PATRL", "TEMPS PRONTO"])
    return resultat

###################################################### RAPPORT OPEN / CLOSE ##############################################

# Gestion de la date automatique  
def set_locale(locale_):
    locale.setlocale(category=locale.LC_ALL, locale=locale_)

def date_en_francais(date_entree = datetime.now()) : 
    set_locale('fr_FR.utf8')    #date = datetime.datetime.now().strftime("%d %B %Y")
    
    date = date_entree.strftime("%d %B %Y")
    entree, sortie = str(date).split(), ""
    for e in entree :
        e = e.replace(e[0], e[0].upper())  
        sortie = sortie + e + " "
    sortie = sortie[0: len(sortie) - 1]    
    return sortie

    
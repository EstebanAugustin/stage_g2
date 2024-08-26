import fitz
import os
import re
import pandas as pd

### IMPORTS ###
# pip install pymupdf
# pip install xlrd
# pip install openpyxl
# pip install pandas

MOTS_CLES = ["ref_parcelle",
             "Nature",
             "Nom / Prénom",
             "Adresse",
             "Droit",
             "Numéro SIREN",
             "Raison sociale",
             "Sexe",
             "Date de naissance",
             "Lieu de Naissance",
             "Compte MAJIC"]

REP_PDF = "./pdf"
REP_TAB_ENTREE = "./tableur"
REP_TAB_SORTIE = './tableur_rempli'

FORMATS_TABLEUR = ["xls", "xlsx", "csv"]


def pdf_vers_texte(chemin_fichier):
    pdf_document = fitz.open(chemin_fichier)
    texte = ""
    for i in range(len(pdf_document)):
        page = pdf_document.load_page(i)
        page_texte = page.get_text()
        texte += page_texte
    return texte

def extrait_infos(texte, mots_cles=MOTS_CLES):
    infos = []
    rows = texte.splitlines()
    for i in range(len(rows)):
        for mot in mots_cles:
            if rows[i].startswith(mot):
                infos.append({mot: rows[i+1]})
    if infos:
        return (0, infos)    
    return (-1, None)

def formate_adresse(adresse_complete):
    modele = r'^(?P<adresse>.*)\s(?P<code_postal>\d{5})\s(?P<commune>.*)$'       
    correspondance = re.match(modele, adresse_complete)
    if correspondance:
        adresse = correspondance.group("code_postal").strip()
        code_postal = correspondance.group("code_postal")
        commune = correspondance.group("commune").strip()
    else:
        modele = r'^(?P<code_postal>\d{5})\s(?P<ville>.*)$'
        correspondance = re.match(modele, adresse_complete)
        if correspondance:
            adresse = ''
            code_postal = correspondance.group('code_postal')
            commune = correspondance.group('ville').strip()
        else:
            adresse = adresse_complete
            code_postal = ''
            commune = ''
    return (adresse, code_postal, commune)

def cree_etat_parcellaire(infos):
    etat_parcellaire = {"Code_INSEE": None,
                        "Préfixe": None,
                        "Section": None,
                        "Numéro": None,
                        "Adresse": None}
    compteur_proprietaires = 0
    for i in range(len(infos)):
        if "ref_parcelle" in infos[i]:
            ref_parcelle = infos[i]["ref_parcelle"]
            etat_parcellaire["Ref_Parcelle"] = ref_parcelle
            etat_parcellaire["Code_INSEE"] = int(ref_parcelle[0:5])
            etat_parcellaire["Préfixe"] = int(ref_parcelle[5:8])
            etat_parcellaire["Section"] = ref_parcelle[8:10]
            etat_parcellaire["Numéro"] = int(ref_parcelle[10:])
        elif "Nature" in infos[i]:
            etat_parcellaire["Nature"] = infos[i]["Nature"]
        elif "Adresse" in infos[i] :
            adresse_complete = infos[i]["Adresse"]
            if compteur_proprietaires == 0:
                etat_parcellaire["Adresse"] = adresse_complete
            else:
                adresse, code_postal, commune = formate_adresse(adresse_complete)
                etat_parcellaire[f"Adresse_Propriétaire_{compteur_proprietaires}"] = adresse
                etat_parcellaire[f"Code_Postal_Propriétaire_{compteur_proprietaires}"] = code_postal
                etat_parcellaire[f"Commune_Propriétaire_{compteur_proprietaires}"] = commune
        elif "Nom / Prénom" in infos[i]:
            compteur_proprietaires += 1
            nom_prenom = infos[i]["Nom / Prénom"].split(' ')
            nom = nom_prenom[0]
            prenom = ' '.join(nom_prenom[1:])
            etat_parcellaire[f"Type_Propriétaire_{compteur_proprietaires}"] = "Personne physique"
            etat_parcellaire[f"Prénom_Propriétaire_{compteur_proprietaires}"] = prenom
            etat_parcellaire[f"Nom_RaisonSociale_Propriétaire_{compteur_proprietaires}"] = nom
        elif "Raison sociale" in infos[i]:
            compteur_proprietaires += 1
            etat_parcellaire[f"Type_Propriétaire_{compteur_proprietaires}"] = "Personne morale"
            etat_parcellaire[f"Nom_RaisonSociale_Propriétaire_{compteur_proprietaires}"] = infos[i]["Raison sociale"]
        elif "Numéro SIREN" in infos[i]:
            etat_parcellaire[f"Numéro_SIREN_Propriétaire_{compteur_proprietaires}"] = infos[i]["Numéro SIREN"]
        elif "Sexe" in infos[i]:
            etat_parcellaire[f"Sexe_Propriétaire_{compteur_proprietaires}"] = infos[i]["Sexe"]
        elif "Date de naissance" in infos[i]:
            etat_parcellaire[f"Date_de_naissance_Propriétaire_{compteur_proprietaires}"] = infos[i]["Date de naissance"]
        elif "Lieu de Naissance" in infos[i]:
            etat_parcellaire[f"Lieu_de_Naissance_Propriétaire_{compteur_proprietaires}"] = infos[i]["Lieu de Naissance"]
        elif "Droit" in infos[i]:
            etat_parcellaire[f"Droit_Propriétaire_{compteur_proprietaires}"] = infos[i]["Droit"]
        elif "Compte MAJIC" in infos[i]:
            etat_parcellaire[f"Identifiant_Foncier_Propriétaire_{compteur_proprietaires}"] = infos[i]["Compte MAJIC"]        
    return etat_parcellaire

def trouve_index_ligne(tab_xls, etat_parcellaire):
    cles = ["Code_INSEE", "Préfixe", "Section", "Numéro"]
    index_col_cles = {}
    colonnes = list(tab_xls.columns)
    for cle in cles:
        index_col_cles[cle] = colonnes.index(cle)
    for index_ligne, contenu_ligne in tab_xls.iterrows():
        score = 0
        for cle, index_col in index_col_cles.items():
            if str(contenu_ligne.iloc[index_col]) == str(etat_parcellaire[cle]):
                score += 1
        if score == len(cles):
            return (0, index_ligne)
    return (-1, None)
        
def remplit_excel(tab_xls, index_ligne, etat_parcellaire):
    for cle, valeur in etat_parcellaire.items():
        if cle in tab_xls.columns:
            tab_xls.loc[index_ligne, cle] = valeur

def affiche_etat_parcellaire(etat_parcellaire):
    max_cle_longueur = max(len(str(cle)) for cle in etat_parcellaire.keys())
    for cle, valeur in etat_parcellaire.items():
        print(f"{str(cle):<{max_cle_longueur}} : {str(valeur)}")

def sauver_tableur_rempli(df_tableur, fichier_tableur, extension="csv"):
    sortie = f"{REP_TAB_SORTIE}/{fichier_tableur}"
    if extension == "csv":
        df_tableur.to_csv(sortie, engine="c", index=False)
    elif extension == "xlsx":
        df_tableur.to_excel(sortie, index=False)
    elif extension == "xls":
        df_tableur.to_excel(sortie+'x', index=False)


if __name__ == "__main__":
    for fichier_tableur in os.listdir(REP_TAB_ENTREE):
        extension = os.path.splitext(fichier_tableur)[1].lstrip('.')
        if extension in FORMATS_TABLEUR:
            df_tableur = pd.read_excel(f"{REP_TAB_ENTREE}/{fichier_tableur}", dtype=str)
            for fichier_pdf in os.listdir(REP_PDF):
                if fichier_pdf.endswith("pdf"):
                    texte = pdf_vers_texte(f"{REP_PDF}/{fichier_pdf}")
                    code_sortie, infos = extrait_infos(texte)
                    if code_sortie == 0:
                        etat_parcellaire = cree_etat_parcellaire(infos)
                        code_sortie, index_ligne = trouve_index_ligne(df_tableur, etat_parcellaire)
                        if code_sortie == 0:
                            remplit_excel(df_tableur, index_ligne, etat_parcellaire)
                            print(f"{etat_parcellaire['Ref_Parcelle']} trouvée dans {fichier_tableur}")
                        else:
                            print(f"{etat_parcellaire['Ref_Parcelle']} non trouvée dans {fichier_tableur}")
                        affiche_etat_parcellaire(etat_parcellaire)
                        print('\n')
                    else:
                        print(f"Aucune info trouvée dans {fichier_pdf}")
            if not os.path.isdir(REP_TAB_SORTIE):
                os.mkdir(REP_TAB_SORTIE)
            sauver_tableur_rempli(df_tableur, fichier_tableur, extension)
            print(f"{fichier_tableur} rempli et sauvé dans le répertoire {REP_TAB_SORTIE}")
        else:
            print(f"{fichier_tableur} n'est pas un fichier tableur valide")

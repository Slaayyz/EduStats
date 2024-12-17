import csv
import matplotlib.pyplot as plt
import numpy as np

def calculer_moyenne_ponderee(liste_credits, liste_notes):
    total_credits = sum(liste_credits)
    if total_credits == 0:
        return 0
    return sum(credit * note for credit, note in zip(liste_credits, liste_notes)) / total_credits    

def lire_donnees_csv(nom_fichier):
    donnees_semestres = {}
    
    with open(nom_fichier, newline='') as fichier:
        lecteur_csv = csv.reader(fichier, delimiter=';')
        
        for ligne in lecteur_csv:
            nom_semestre, nom_ue, nom_module, nb_credits, valeur_note = ligne
            if valeur_note == '':
                continue
            try:
                nb_credits = int(nb_credits)
                valeur_note = float(valeur_note)
            except ValueError:
                continue
            
            if nom_semestre not in donnees_semestres:
                donnees_semestres[nom_semestre] = {}
            if nom_ue not in donnees_semestres[nom_semestre]:
                donnees_semestres[nom_semestre][nom_ue] = {'credits': [], 'notes': []}
            
            donnees_semestres[nom_semestre][nom_ue]['credits'].append(nb_credits)
            donnees_semestres[nom_semestre][nom_ue]['notes'].append(valeur_note)
    
    moyennes_par_semestre = {}
    for nom_semestre, ues in donnees_semestres.items():
        moyennes_par_semestre[nom_semestre] = {}
        for nom_ue, donnees in ues.items():
            moyennes_par_semestre[nom_semestre][nom_ue] = calculer_moyenne_ponderee(donnees['credits'], donnees['notes'])
    
    return moyennes_par_semestre

def generer_graphique_moyennes(moyennes_par_semestre):
    liste_semestres = list(moyennes_par_semestre.keys())
    liste_ues = sorted({nom_ue for semestre in moyennes_par_semestre.values() for nom_ue in semestre})
    valeurs_moyennes = {semestre: [moyennes_par_semestre[semestre].get(nom_ue, 0) for nom_ue in liste_ues] for semestre in liste_semestres}
    
    largeur_barre = 0.8 / len(liste_semestres)
    indices_ues = np.arange(len(liste_ues))
    
    plt.figure()
    plt.title("Moyennes pondérées par semestre")
    plt.ylim(0, 20)
    plt.axhline(10, color="red", linestyle="--", label="10 - Seuil de réussite")
    plt.axhline(8, color="orange", linestyle="--", label="8 - Seuil d'alerte")
    
    for i, nom_semestre in enumerate(liste_semestres):
        positions_barres = indices_ues + i * largeur_barre
        couleurs_barres = ["green" if moyenne >= 10 else "orange" if moyenne >= 8 else "red" for moyenne in valeurs_moyennes[nom_semestre]]
        plt.bar(positions_barres, valeurs_moyennes[nom_semestre], largeur_barre, label=nom_semestre, color=couleurs_barres)
    
    plt.xticks(indices_ues + (largeur_barre * (len(liste_semestres) - 1)) / 2, liste_ues)
    plt.xlabel("Unités d'enseignement")
    plt.ylabel("Moyenne pondérée")
    plt.legend()
    plt.show()

def main():
    nom_fichier = 'notes.csv'
    moyennes_par_semestre = lire_donnees_csv(nom_fichier)
    
    for nom_semestre, moyennes in moyennes_par_semestre.items():
        print(f"\n{nom_semestre} :")
        for nom_ue, moyenne in moyennes.items():
            print(f"  {nom_ue} : {moyenne:.2f}")
    
    generer_graphique_moyennes(moyennes_par_semestre)

if __name__ == "__main__":
    main()

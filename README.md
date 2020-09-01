# occupation_locaux
Gestion de l'occupation des locaux dans le cadre du COVID19


[![DOI](https://zenodo.org/badge/264869258.svg)](https://zenodo.org/badge/latestdoi/264869258)


## Contexte
- Besoin de gérer l'occupation des locaux de l'IRHS et de tracer leur utilisation par demi journée.  
- Un formulaire googleForm a été crée et est rempli par les RE.  
- La récupération des réponses au formulaire ne correspond pas à la mise en forme désirée.  

**Ecriture d'un script pour la mise en forme des réponses**

## Utilisation
Le formulaire à remplir est d'après le modèle suivant :   
Les résultats sont sur un fichier google Sheets du même nom.
Un script utilisant un fichier template crée un nouvel onglet appelé "Bilan" où les résultats du formulaire sont représenté comme souhaité.
Pour utiliser ce script, il faut le lancer depuis Outil > Macros > makeBilan

## Caractéristiques
- script en javascript
- utilisation du fichier template

# Google Apps Script
#### Lien de partage pour le script covid/code.gs : 
https://script.google.com/d/17mQd1O3R3pC8ttXfvIQtFXQMimO35Cf1M3zT9L22v45EmGQedllzlbir/edit?usp=sharing

#### Id du script :
17mQd1O3R3pC8ttXfvIQtFXQMimO35Cf1M3zT9L22v45EmGQedllzlbir

# Mise en place
- Sur Google Drive :
  - Duplication du formulaire
  - Renommage du fichier
  - Ouverture
- Sur le formulaire :
  - sélection de l'onglet "Réponses"
  - Clic sur "Afficher les réponses dans Sheets" (symbole google Sheet)
- Sur le fichier excel de résultats :
  - Outils > Editeur de scripts
- Sur l'Editeur de script :
  - Ressources > Bibliothèques...
    - Donner un nom au projet de script <nom du projet>
    - Add a library : copier l'Id du script ci dessus
    - Ajouter
    - Enregistrer
  - Coller la fonction  suivante : 
  ```
  function makeBilan() {
    Covid.makeBilan()
  }
  ```
  - Enregistrer
- De nouveau sur le fichier excel de résultats
  - Outils > Macro > Importer
  - makeBilan > Ajouter la fonction

Lors de la première utilisation du script, une demande d'autorisation est demander avant l'exécution :
  - Sélection du compte propriétaire
  - Paramètres avancés
  - Accéder à <Nom du projet> (non sécurisé)
  - Autoriser

# Utilisation du script :
Depuis la feuille excel : 
- Outils > Macros > makeBilan  

# Déroulement du script :
- Copie de la feuille template
- Mise en place de la feuille Bilan : 
  - Si une feuille Bilan existe déjà, elle est supprimée
  - Renommage de la copie de Template par "Bilan"
  - Création de la liste "zones"
  - Création de la liste "jours"
- Récupération de chaque données :
  - Récupération du nom de l'agent
  - Récupération de la période (jour)
  - Récupération de la zone
- Répartition des données dan le tableau de bilan :
  - Récupération de la ligne de la zone
  - Récupération de la colonne du jour
  - Si la cellule n'est pas vide : copie à la suite du nom de l'agent
  - Sinon (si la cellule est vide) : copie du nom de l'agent
- Mise en couleur en fonction du nombre max d'agent par zone (présentes dans le template)
  - Comparaison du nombre d'agent avec le nombre max autorisé
  - Si le nombre d'agent est égal : colorisation en jaune
  - Si le nombre d'agent est supérieur : colorisation en orange
- "Done" en A3 quand le script est fini.

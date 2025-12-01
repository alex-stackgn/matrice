# Traitement de Matrices ODS

## Description

Application graphique permettant de traiter par lots des fichiers ODS (LibreOffice Calc). Elle offre une interface complète pour modifier, formater et manipuler plusieurs fichiers simultanément en définissant une série d'opérations à appliquer automatiquement.

## Prérequis

```bash
pip install PySide6 ezodf qt_material --break-system-packages
```

## Lancement

```bash
python traitement_matrice.py
```

## Architecture de l'interface

L'application est organisée en 5 onglets principaux, chacun gérant un aspect spécifique du traitement des fichiers.

---

## 1. Onglet "Fichiers"

Cet onglet gère la sélection des fichiers à traiter et les options de sortie.

### Boutons et fonctions

**Bouton "Ajouter Fichiers"**
- Ouvre un explorateur de fichiers pour sélectionner un ou plusieurs fichiers .ods
- Les fichiers sélectionnés apparaissent dans la liste à gauche
- Permet la sélection multiple

**Bouton "Ajouter Dossier"**
- Scanne un dossier entier à la recherche de fichiers .ods
- Trouve automatiquement tous les fichiers avec l'extension .ods dans le dossier sélectionné
- Ajoute tous les fichiers trouvés à la liste de traitement

**Bouton "Retirer Sélection"**
- Supprime de la liste les fichiers actuellement sélectionnés
- Permet de retirer des fichiers sans recommencer la sélection depuis zéro

**Bouton "Vider Liste"**
- Supprime tous les fichiers de la liste en une seule action
- Réinitialise complètement la sélection

**Bouton "Choisir Dossier de Sortie"**
- Définit le répertoire où seront enregistrés les fichiers traités
- Par défaut, utilise le dossier "output" dans le répertoire courant

**Case à cocher "Utiliser dossier source comme sortie"**
- Lorsque cochée, les fichiers traités sont enregistrés dans le même dossier que le fichier source
- Ignore le dossier de sortie défini

---

## 2. Onglet "Structure"

Cet onglet permet de définir des modifications structurelles sur les feuilles de calcul.

### Formulaire d'ajout

**Champ "Feuille"**
- Nom de la feuille dans laquelle effectuer l'opération
- Exemple : "Feuille1", "Données", "Résultats"

**Menu déroulant "Action"**
- **Insérer Lignes** : Insère des lignes vides à une position donnée
- **Fusionner** : Fusionne plusieurs cellules en une seule
- **Effacer** : Efface le contenu d'une plage de cellules

**Champ "Détails" (contexte selon l'action)**
- Pour "Insérer Lignes" : Format "X ligne(s) à partir de Y" (ex: "5 ligne(s) à partir de 10")
- Pour "Fusionner" : Plage de cellules (ex: "A1:C3")
- Pour "Effacer" : Plage de cellules à effacer (ex: "B2:D10")

**Champ "Couleur Fond" (optionnel)**
- Code couleur hexadécimal pour appliquer un fond coloré
- Format : #RRGGBB (ex: #FF0000 pour rouge)
- Utilisé uniquement avec "Insérer Lignes"

**Bouton "Ajouter Action"**
- Ajoute l'opération définie à la liste des actions structurelles
- L'opération apparaît dans l'arbre en dessous

**Bouton "Retirer Sélection"**
- Supprime l'action sélectionnée dans l'arbre

---

## 3. Onglet "Contenu"

Cet onglet permet d'insérer des valeurs ou des grilles de données dans les cellules.

### Formulaire d'ajout

**Champ "Feuille"**
- Nom de la feuille cible

**Menu déroulant "Type"**
- **Valeur** : Insère une valeur unique dans une cellule ou remplit une plage avec la même valeur
- **Grille** : Colle un tableau de données à partir d'une cellule de départ

**Champ "Cellule/Plage"**
- Pour "Valeur" : Cellule unique (ex: "A1") ou plage (ex: "A1:C5")
- Pour "Grille" : Cellule de départ du coin supérieur gauche (ex: "B3")

**Champ "Détails"**
Pour le type "Valeur" :
- Format : "type: valeur"
- Types disponibles : string, int, float
- Exemples : "string: Bonjour", "int: 42", "float: 3.14"

Pour le type "Grille" :
- Cliquer sur le bouton "Éditer Grille" ouvre une fenêtre
- Permet de coller ou saisir un tableau de données
- Supporte le format tabulé (copier-coller depuis Excel/Calc)
- Supporte aussi le format CSV avec point-virgule

**Bouton "Éditer Grille"**
- Disponible uniquement pour le type "Grille"
- Ouvre une fenêtre modale avec un champ texte multiligne
- Permet de coller des données tabulées
- Valide et stocke les données lors de la confirmation

**Bouton "Ajouter Action"**
- Ajoute l'opération de contenu à la liste

**Bouton "Retirer Sélection"**
- Supprime l'action sélectionnée

---

## 4. Onglet "Style"

Cet onglet permet d'appliquer des formatages visuels aux cellules.

### Formulaire d'ajout

**Champ "Feuille"**
- Nom de la feuille dans laquelle appliquer le style

**Champ "Cellules"**
- Liste de cellules séparées par des virgules
- Exemple : "A1, B2, C3" ou "A1:A10"

**Case à cocher "Gras"**
- Applique le style gras au texte des cellules

**Case à cocher "Retour à la ligne"**
- Active le retour à la ligne automatique dans les cellules

**Champ "Couleur Fond"**
- Code couleur hexadécimal pour le fond de cellule
- Format : #RRGGBB

**Champ "Taille Police"**
- Taille de la police en points
- Exemple : 12, 14, 16

**Bouton "Ajouter Style"**
- Ajoute les paramètres de style à la liste
- Génère un résumé textuel des options choisies

**Bouton "Retirer Sélection"**
- Supprime le style sélectionné

---

## 5. Onglet "Copie"

Cet onglet permet de copier des plages de cellules d'un endroit à un autre, même entre différentes feuilles.

### Formulaire d'ajout

**Champ "Feuille Source"**
- Nom de la feuille d'où copier les données

**Champ "Plage Source"**
- Plage de cellules à copier
- Format : "A1:C5"

**Champ "Feuille Destination"**
- Nom de la feuille où coller les données
- Tapez "idem" pour coller dans la même feuille que la source

**Champ "Cellule Haut-Gauche Destination"**
- Cellule de départ pour le collage
- La plage copiée sera collée à partir de cette position

**Case à cocher "Transposer"**
- Lorsque cochée, inverse lignes et colonnes lors du collage
- Utile pour transformer des données horizontales en verticales

**Bouton "Ajouter Copie"**
- Ajoute l'opération de copie à la liste

**Bouton "Retirer Sélection"**
- Supprime la copie sélectionnée

---

## 6. Onglet "Options"

Cet onglet regroupe les options globales qui s'appliquent à tous les fichiers traités.

### Options de nommage

**Case à cocher "Incrémenter version dans le nom"**
- Recherche un numéro de version dans le nom du fichier (format vX.Y.Z)
- Incrémente automatiquement la dernière partie
- Exemple : "rapport_v1.2.3.ods" devient "rapport_v1.2.4.ods"
- Si aucune version n'est trouvée, le fichier est simplement copié

**Champ "Remplacer texte entre parenthèses"**
- Remplace le texte entre parenthèses dans le nom du fichier
- Exemple : si le fichier s'appelle "rapport_(brouillon).ods" et que vous tapez "final", il devient "rapport_(final).ods"
- Laissez vide pour ne pas modifier le texte entre parenthèses

### Options de réinitialisation des couleurs

**Case à cocher "Réinitialiser couleurs avant traitement"**
- Efface toutes les couleurs de fond d'une zone avant d'appliquer les opérations
- Utile pour nettoyer un fichier existant avant d'appliquer un nouveau formatage

**Champ "Feuilles concernées"**
- Liste des noms de feuilles séparés par des virgules
- Exemple : "Feuille1, Données, Résultats"
- Seules ces feuilles verront leurs couleurs réinitialisées

**Champ "Ligne de début"**
- Numéro de la première ligne à nettoyer (base 1)
- Par défaut : 10

**Champ "Lignes à exclure"**
- Liste de numéros de lignes à ignorer lors du nettoyage
- Séparés par des virgules
- Exemple : "1, 2, 5" pour préserver les lignes 1, 2 et 5

**Bouton "Couleurs Colonnes (A=...)"**
- Ouvre une fenêtre permettant de définir des couleurs par colonne
- Format : une ligne par colonne
- Syntaxe : `LETTRE=#COULEUR`
- Exemple :
  ```
  A=#FF0000
  B=#00FF00
  C=#0000FF
  ```
- Les colonnes définies ici seront automatiquement colorées après la réinitialisation

---

## Menu principal

### Menu "Fichier"

**"Sauvegarder Config"**
- Enregistre toutes les opérations et options dans un fichier JSON
- Permet de réutiliser une configuration complète ultérieurement
- Sauvegarde :
  - Les actions de structure
  - Les actions de contenu
  - Les styles
  - Les copies
  - Toutes les options globales

**"Charger Config"**
- Charge une configuration précédemment sauvegardée
- Restaure complètement l'état de l'application
- Remplace toutes les opérations actuelles

**"Quitter"**
- Ferme l'application

---

## Bouton de traitement

**Bouton "TRAITER FICHIERS"** (en bas de fenêtre)
- Lance le traitement par lots de tous les fichiers sélectionnés
- Affiche une barre de progression
- Exécute les opérations dans l'ordre :
  1. Structure (insertion de lignes, fusion, effacement)
  2. Contenu (valeurs et grilles)
  3. Style (formatage)
  4. Copie (duplication de cellules)
- Applique les options globales (incrémentation version, réinitialisation couleurs)
- Affiche un message de confirmation à la fin

---

## Ordre d'exécution des opérations

Il est important de comprendre que les opérations sont exécutées dans un ordre précis :

1. **Réinitialisation des couleurs** (si activée)
2. **Structure** : Modifications structurelles de la feuille
3. **Contenu** : Insertion de valeurs et grilles
4. **Style** : Application du formatage
5. **Copie** : Duplication de cellules

Cet ordre garantit que les références de cellules restent cohérentes. Par exemple, si vous insérez 5 lignes à la ligne 10, les opérations suivantes doivent tenir compte de ce décalage.

---

## Astuces d'utilisation

### Gestion des références de cellules

- Toujours utiliser la notation A1 (ex: A1, B12, Z99)
- Les plages s'écrivent avec deux points (ex: A1:C5)
- Les références sont insensibles à la casse

### Workflow recommandé

1. Sélectionner les fichiers à traiter
2. Définir d'abord les opérations structurelles (insertions, fusions)
3. Ajouter ensuite le contenu
4. Appliquer les styles
5. Configurer les copies si nécessaire
6. Vérifier les options globales
7. Sauvegarder la configuration si vous comptez la réutiliser
8. Lancer le traitement

### Sauvegarde des configurations

Pour éviter de reconfigurer l'application à chaque utilisation :
- Créez une configuration pour chaque type de traitement récurrent
- Utilisez des noms de fichiers explicites (ex: "config_rapport_mensuel.json")
- Les configurations sont portables et peuvent être partagées

### Débogage

En cas d'erreur pendant le traitement :
- Les erreurs sont affichées dans la console
- Le fichier source n'est jamais modifié (un nouveau fichier est créé)
- Vérifiez que les noms de feuilles existent dans vos fichiers
- Vérifiez la syntaxe des plages de cellules

---

## Limitations connues

- Les formules complexes dans les cellules copiées peuvent ne pas être préservées correctement
- Le traitement de très gros fichiers (.ods > 50 MB) peut être lent
- Les styles très spécifiques (dégradés, motifs complexes) ne sont pas supportés
- Les graphiques et images ne sont pas modifiés par l'application

---

## Dépannage

**L'application ne se lance pas**
- Vérifiez que toutes les dépendances sont installées
- Assurez-vous d'avoir Python 3.8 ou supérieur

**Les fichiers traités sont corrompus**
- Vérifiez que vos fichiers sources sont bien au format .ods valide
- Essayez de les ouvrir dans LibreOffice avant traitement
- Évitez les caractères spéciaux dans les noms de feuilles

**Les couleurs ne s'appliquent pas**
- Vérifiez le format hexadécimal : #RRGGBB (6 caractères après le #)
- Exemple correct : #FF0000
- Exemple incorrect : FF0000 ou #F00

**Les plages de cellules ne fonctionnent pas**
- Vérifiez la syntaxe : A1:C5 (pas d'espaces)
- Assurez-vous que la plage existe dans la feuille
- Les indices de colonnes vont jusqu'à ZZ

---

## Support

Pour toute question ou problème, référez-vous au code source qui contient des commentaires détaillés sur chaque fonction.
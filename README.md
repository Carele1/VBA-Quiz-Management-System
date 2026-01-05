# VBA-Quiz-Management-System
Système complet de gestion de QCM développé en VBA Excel avec suivi statistique

Projet académique réalisé par **DIFFO CARELE** et **TSAFACK ANGE** pour l'année 2025-2026.

## 📝 Présentation du Projet
Ce projet consiste à développer un système complet de gestion de questionnaires (QCM) utilisant VBA sous Excel. L'application permet de générer, administrer et analyser des questionnaires pédagogiques de manière automatisée.

## 🚀 Architecture de l'Application
L'écosystème du projet est divisé en plusieurs interfaces distinctes pour une meilleure sécurité et organisation:

### 👨‍🏫 Interface Auteur (`QuizAuteur.xlsm`)
* **Accès sécurisé** : Protégé par un mot de passe (par défaut : `1234`) via le formulaire `UF_PassWord`.
* **Gestion des thèmes** : Création de nouveaux thèmes avec gestion optionnelle des niveaux de difficulté (Facile, Moyen, Difficile).
* **Édition de questions** : Formulaire `UF_EntreeQuestion` permettant de saisir une question avec 3 réponses, l'attribution de points et des commentaires associés.
* **Contrôle automatique** : Vérification des doublons et de la cohérence des points lors de l'enregistrement dans `Questionnaires.xlsx`.

### 🎓 Interface Élève (`QuizEleve.xlsm`)
* **Déroulement fluide** : L'élève choisit son thème, le nombre de questions et son nom via le formulaire `UF_DefQuest`.
* **Expérience interactive** : Intégration d'effets sonores de victoire ou de défaite selon la réponse.
* **Système de reprise** : Possibilité de reprendre un quiz déjà effectué en sélectionnant le nom, le thème et la date de la session précédente.
* **Génération PDF** : Exportation automatisée des épreuves sous forme de fichier PDF pour des évaluations papier.

### 📈 Statistiques & Suivi (`QuizStat.xlsm`)
* **Tableaux de bord** : Résumé détaillé des performances par étudiant (score, date, thème) et par thématique (moyenne, nombre de quiz, temps de réponse).
***Visualisation** : Graphiques illustrant les performances individuelles et les notes moyennes par thème sur la feuille "Graphes".

## 📂 Organisation des Fichiers
Le projet utilise une structure multi-fichiers pour séparer les données du code:
* **`src/`** : Contient `QuizAuteur.xlsm`, `QuizEleve.xlsm` et `QuizStat.xlsm`.
* **`resources/`** : Stocke `Questionnaires.xlsx` et `Eleves.xls`.
* **`docs/`** : Contient le rapport technique du projet.

## 🛠️ Installation
1. Téléchargez l'intégralité du dépôt.
2. Assurez-vous que tous les fichiers Excel sont dans le même répertoire racine pour maintenir les liaisons de données.
3. Activez les macros lors de l'ouverture des fichiers `.xlsm`.

**NB**: Lors de l'éxécution les dossiers T, N et PDF, seront creér automatiquement, chacun comportant respetivement, les fichier T.txt, N.txt et PDF générés .

# VBA-Quiz-Management-System
Système complet de gestion de QCM développé en VBA Excel avec suivi statistique

[cite_start]Projet académique réalisé par **DIFFO CARELE** et **TSAFACK ANGE** pour l'année 2025-2026[cite: 1, 2, 5].

## 📝 Présentation du Projet
[cite_start]Ce projet consiste à développer un système complet de gestion de questionnaires (QCM) utilisant VBA sous Excel[cite: 7]. [cite_start]L'application permet de générer, administrer et analyser des questionnaires pédagogiques de manière automatisée[cite: 17].

## 🚀 Architecture de l'Application
[cite_start]L'écosystème du projet est divisé en plusieurs interfaces distinctes pour une meilleure sécurité et organisation[cite: 8]:

### 👨‍🏫 Interface Auteur (`QuizAuteur.xlsm`)
* [cite_start]**Accès sécurisé** : Protégé par un mot de passe (par défaut : `1234`) via le formulaire `UF_PassWord`[cite: 11, 32].
* [cite_start]**Gestion des thèmes** : Création de nouveaux thèmes avec gestion optionnelle des niveaux de difficulté (Facile, Moyen, Difficile)[cite: 12, 33, 36, 37, 38].
* [cite_start]**Édition de questions** : Formulaire `UF_EntreeQuestion` permettant de saisir une question avec 3 réponses, l'attribution de points et des commentaires associés[cite: 41, 43, 44, 46, 47].
* [cite_start]**Contrôle automatique** : Vérification des doublons et de la cohérence des points lors de l'enregistrement dans `Questionnaires.xlsx`[cite: 48, 49].

### 🎓 Interface Élève (`QuizEleve.xlsm`)
* [cite_start]**Déroulement fluide** : L'élève choisit son thème, le nombre de questions et son nom via le formulaire `UF_DefQuest`[cite: 51, 52, 53, 54, 61].
* [cite_start]**Expérience interactive** : Intégration d'effets sonores de victoire ou de défaite selon la réponse[cite: 63, 131].
* [cite_start]**Système de reprise** : Possibilité de reprendre un quiz déjà effectué en sélectionnant le nom, le thème et la date de la session précédente[cite: 67, 68, 69, 70].
* [cite_start]**Génération PDF** : Exportation automatisée des épreuves sous forme de fichier PDF pour des évaluations papier[cite: 76].

### 📈 Statistiques & Suivi (`QuizStat.xlsm`)
* [cite_start]**Tableaux de bord** : Résumé détaillé des performances par étudiant (score, date, thème) et par thématique (moyenne, nombre de quiz, temps de réponse)[cite: 20, 22, 23].
* [cite_start]**Visualisation** : Graphiques illustrant les performances individuelles et les notes moyennes par thème sur la feuille "Graphes"[cite: 24, 81].

## 📂 Organisation des Fichiers
[cite_start]Le projet utilise une structure multi-fichiers pour séparer les données du code[cite: 120, 121]:
* **`src/`** : Contient `QuizAuteur.xlsm`, `QuizEleve.xlsm` et `QuizStat.xlsm`.
* [cite_start]**`resources/`** : Stocke `Questionnaires.xlsx` et `Eleves.xlsx`[cite: 25, 28].
* **`docs/`** : Contient le rapport technique du projet.

## 🛠️ Installation
1. Téléchargez l'intégralité du dépôt.
2. Assurez-vous que tous les fichiers Excel sont dans le même répertoire racine pour maintenir les liaisons de données.
3. Activez les macros lors de l'ouverture des fichiers `.xlsm`.

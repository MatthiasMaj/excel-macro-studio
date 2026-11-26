<div align="center">

# Excel Macro Studio

**Générateur de macros Excel VBA, analyseur de classeurs et cursus d'apprentissage interactif — 100 % local, sans IA externe ni transfert de données.**

[![Démo en ligne](https://img.shields.io/badge/D%C3%A9mo%20en%20ligne-GitHub%20Pages-2563eb?style=for-the-badge&logo=githubpages&logoColor=white)](https://matthiasmaj.github.io/excel-macro-studio/)
[![Version](https://img.shields.io/badge/Version-1.0.0-16a34a?style=for-the-badge)](./package.json)
[![Tests](https://img.shields.io/badge/Tests-10%2F10%20passants-10b981?style=for-the-badge&logo=node.js&logoColor=white)](./tests/studio.test.mjs)
[![Déploiement](https://img.shields.io/badge/CI%2FCD-GitHub%20Actions-2088ff?style=for-the-badge&logo=githubactions&logoColor=white)](./.github/workflows/deploy.yml)
[![100% Local](https://img.shields.io/badge/Confidentialit%C3%A9-100%25%20Local-0f172a?style=for-the-badge)](#confidentialit%C3%A9-et-s%C3%A9curit%C3%A9)

[Accéder à l'application en ligne](https://matthiasmaj.github.io/excel-macro-studio/) • [Fonctionnalités](#-fonctionnalit%C3%A9s-principales) • [Catalogue (129 opérations)](#-catalogue-des-12-cat%C3%A9gories) • [Cursus VBA](#-cursus-complet-dapprentissage-vba) • [Installation locale](#-installation-et-d%C3%A9veloppement-local)

</div>

---

## À propos

**Excel Macro Studio** est une application web conçue pour générer rapidement du code **VBA propre, commenté et prêt à l'emploi** pour Microsoft Excel, analyser la structure de classeurs `.xlsx` / `.xlsm` / `.xls` / `.csv` directement dans le navigateur, et accompagner les utilisateurs pas à pas de la découverte d'Excel jusqu'à l'automatisation professionnelle grâce à un **cursus progressif en 4 niveaux et 12 chapitres**.

Toute l'exécution s'effectue **côté client** dans votre navigateur : aucun fichier Excel, aucune donnée métier et aucun code généré n'est envoyé vers un serveur ou un service externe.

---

## ✨ Fonctionnalités principales

### 1. Trois modes de génération de macros VBA
* **Mode Guidé** : choisissez une catégorie et une opération parmi les **129 opérations** du catalogue, remplissez un formulaire dynamique validé, et obtenez immédiatement le module VBA adapté à vos feuilles et plages.
* **Description libre (Langage naturel local)** : décrivez votre besoin en français courant ; le moteur d'analyse sémantique local (`naturalLanguageEngine.js`) détecte l'intention, extrait vos noms de feuilles, colonnes ou seuils, et génère la macro correspondante sans appel réseau.
* **Compositeur de Workflows** : enchaînez plusieurs opérations dans un pipeline ordonné pour générer une macro orchestrée multi-étapes.

### 2. Analyseur de classeur Excel intégré
* Importez un fichier `.xlsx`, `.xlsm`, `.xls` ou `.csv` directement dans l'interface (parsing local via `xlsx` / SheetJS).
* Inspectez la structure des feuilles, les en-têtes détectés, les plages et les types de données afin de pré-remplir et fiabiliser vos macros.

### 3. Qualité du code VBA et mode débutant
* Code structuré avec `Option Explicit`, gestion d'erreurs (`On Error GoTo ...`), optimisation des performances (`Application.ScreenUpdating`, `Application.Calculation`, `Application.EnableEvents`) et nettoyage des objets.
* **Coloration syntaxique VBA** intégrée (`vbaHighlighter.js`).
* **Mode commentaires débutants** (`vbaBeginnerAnnotator.js`) activable/désactivable pour annoter chaque ligne clé en français clair.
* **Raffinement VBA** (`vbaRefiner.js`), bibliothèque de modèles (`templatesLibrary.js`) et historique local persistant.

### 4. Exports multi-formats
* Copie directe dans le presse-papier pour collage dans l'éditeur VBA d'Excel (`Alt + F11`).
* Export en module VBA standard **`.bas`** prêt à importer dans un classeur.
* Export texte **`.txt`** et génération de classeur d'entraînement **`.xlsx`**.

---

## 📚 Catalogue des 12 catégories

Excel Macro Studio couvre **129 opérations** réparties sur **12 domaines métier** (`A` à `L`) :

| Code | Catégorie | Opérations | Exemples d'usages |
| :---: | :--- | :---: | :--- |
| **A** | Gestion des cellules et des données | 15 | Copie dynamique, nettoyage, détection de doublons, recherche/remplacement |
| **B** | Gestion des feuilles et classeurs | 13 | Création, tri, protection, consolidation et archivage d'onglets |
| **C** | Traitement et analyse des données | 15 | Filtres avancés, tris multi-critères, tableaux croisés, agrégations |
| **D** | Mise en forme et présentation | 10 | Mise en forme conditionnelle, ajustement automatique, styles homogènes |
| **E** | Gestion des dates et du temps | 10 | Calculs d'échéances, jours ouvrés, regroupements mensuels/annuels |
| **F** | Gestion financière et comptable | 11 | Calculs de TVA, marges, lettrage, rapprochements et synthèses |
| **G** | Gestion administrative et RH | 9 | Suivi des congés, ancienneté, plannings et contrôles de dossiers |
| **H** | Gestion commerciale et stocks | 9 | Alertes de seuil de stock, suivi de commandes, CA par client/produit |
| **I** | Génération de documents et exports | 9 | Exports PDF ciblés, publipostage, génération de rapports prêts à diffuser |
| **J** | Automatisation avancée | 13 | Boucles sur classeurs, événements Excel, dictionnaires et tableaux mémoire |
| **K** | Graphiques et tableaux de bord | 7 | Création et mise à jour automatique de graphiques et KPI |
| **L** | Importation et communication | 8 | Import CSV/texte structuré, préparation d'e-mails et logs d'exécution |
| **Total** | **12 catégories** | **129** | **Couverture complète des besoins bureautiques et métier** |

---

## 🎓 Cursus complet d'apprentissage VBA

Pour les personnes souhaitant partir de zéro et devenir autonomes jusqu'au niveau professionnel, l'application intègre un **cursus interactif complet** structuré autour d'un cas fil rouge réaliste :

* **4 niveaux progressifs** et **12 chapitres guidés** (`src/services/curriculumData.js`).
* **Exemples de code intégralement commentés** ligne par ligne pour les débutants.
* **Exercices interactifs** variés à chaque étape (remise en ordre de blocs, code à trous, détection et correction de bugs) avec guidage pédagogique.
* **Projet fil rouge — `Entreprise_NovaTech_FilRouge.xlsx`** généré directement depuis l'outil (`src/services/trainingWorkbook.js`), contenant les onglets :
  * `Ventes_Brutes`
  * `Clients`
  * `Stocks`
  * `Parametres`

---

## 🗂️ Architecture du projet

```text
excel-macro-studio/
├── .github/
│   └── workflows/
│       └── deploy.yml                # Pipeline CI/CD (tests, build et déploiement GitHub Pages)
├── src/
│   ├── catalog/
│   │   ├── categories.js             # Définition des 12 catégories (A à L)
│   │   ├── index.js                  # Agrégation du catalogue (129 opérations)
│   │   └── ops/                      # Définitions par catégorie (catA_cells.js .. catL_imports.js) + _helpers.js
│   ├── core/
│   │   ├── security.js               # Utilitaires de sûreté et échappement
│   │   ├── store.js                  # État applicatif et historique local
│   │   ├── theme.js                  # Gestion du thème clair/sombre
│   │   └── validation.js             # Validation des paramètres formulaires
│   ├── engine/
│   │   ├── vbaBeginnerAnnotator.js   # Annotation pédagogique ligne par ligne
│   │   ├── vbaGenerator.js           # Générateur principal de modules VBA
│   │   ├── vbaHighlighter.js         # Coloration syntaxique VBA
│   │   ├── vbaRefiner.js             # Optimisation et nettoyage du code généré
│   │   ├── vbaUtils.js               # Helpers de génération VBA
│   │   └── workflowComposer.js       # Composition de macros multi-étapes
│   ├── services/
│   │   ├── curriculumData.js         # Contenu du cursus (4 niveaux, 12 chapitres, exercices)
│   │   ├── excelAnalyzer.js          # Analyse locale de fichiers Excel/CSV via SheetJS
│   │   ├── exportService.js          # Exports .bas, .txt et téléchargements
│   │   ├── naturalLanguageEngine.js  # Moteur d'interprétation de requêtes en français
│   │   ├── templatesLibrary.js       # Bibliothèque de modèles prêts à l'emploi
│   │   └── trainingWorkbook.js       # Génération du classeur fil rouge Entreprise_NovaTech_FilRouge.xlsx
│   ├── styles/
│   │   ├── base.css
│   │   ├── components.css
│   │   ├── layout.css
│   │   └── tokens.css
│   ├── ui/
│   │   ├── analyzerView.js           # Vue Analyseur Excel
│   │   ├── catalogView.js            # Navigation catalogue
│   │   ├── formEngine.js             # Génération dynamique des formulaires
│   │   ├── freeView.js               # Mode Description libre
│   │   ├── guidedView.js             # Mode Guidé
│   │   ├── historyView.js            # Historique local des macros
│   │   ├── homeView.js               # Accueil
│   │   ├── libraryView.js            # Bibliothèque de modèles
│   │   ├── resultPanel.js            # Panneau de prévisualisation, annotation et export
│   │   ├── tutorialView.js           # Interface du cursus interactif et exercices
│   │   └── workflowView.js           # Interface du compositeur de workflows
│   └── main.js                       # Point d'entrée de l'application
├── tests/
│   └── studio.test.mjs               # Suite de tests automatisés (node --test)
├── index.html
├── package.json
└── vite.config.js
```

---

## 🚀 Installation et développement local

### Prérequis
* **Node.js** (version récente supportant les modules ESM et `node --test`, ex. Node 20+) et **npm**.

### Commandes

```bash
# 1. Cloner le dépôt
git clone https://github.com/MatthiasMaj/excel-macro-studio.git
cd excel-macro-studio

# 2. Installer les dépendances
npm install

# 3. Lancer le serveur de développement Vite (http://localhost:5173)
npm run dev

# 4. Exécuter les tests automatisés
npm test

# 5. Construire le bundle de production dans dist/
npm run build

# 6. Prévisualiser le build de production localement
npm run preview
```

### Stack technique
* **JavaScript ES Modules** (Vanilla JS modulaire, sans framework lourd)
* **Vite** (`^6.2.0`) pour le développement et le build (`base: './'`)
* **SheetJS / `xlsx`** (`^0.18.5`) pour la lecture et la génération locale de classeurs Excel
* **`node --test`** pour les tests automatisés (`tests/studio.test.mjs`)

---

## 🌐 Déploiement GitHub Pages

Le workflow [`.github/workflows/deploy.yml`](./.github/workflows/deploy.yml) automatise la livraison continue à chaque push sur `main` ou `master` :

1. Installation propre des dépendances (`npm ci`)
2. Exécution de la suite de tests (`npm test`)
3. Build de production (`npm run build`)
4. Déploiement du dossier `dist/` sur GitHub Pages (`actions/configure-pages@v5`, `actions/upload-pages-artifact@v3`, `actions/deploy-pages@v4`)

L'application déployée est accessible à l'adresse :
👉 **https://matthiasmaj.github.io/excel-macro-studio/**

---

## 🔒 Confidentialité et sécurité

* **Zéro transfert de données** : l'analyse des fichiers Excel, l'interprétation en langage naturel et la génération de code VBA s'exécutent intégralement dans votre navigateur.
* **Aucune API d'IA externe** : aucun secret ni clé API requis, fonctionnement hors ligne possible après chargement.
* **Bonnes pratiques VBA intégrées** : génération défensive avec déclarations explicites, gestion d'erreurs et restauration systématique des paramètres Excel en fin d'exécution.

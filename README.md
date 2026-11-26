<div align="center">

# Excel Macro Studio

**Académie d'apprentissage Excel & VBA, Assistant d'automatisation (VBA & Office Scripts), Audit de macros et Analyseur de classeurs — 100 % local, sans IA externe ni transfert de données.**

[![Démo en ligne](https://img.shields.io/badge/D%C3%A9mo%20en%20ligne-GitHub%20Pages-2563eb?style=for-the-badge&logo=githubpages&logoColor=white)](https://matthiasmaj.github.io/excel-macro-studio/)
[![Version](https://img.shields.io/badge/Version-1.0.0-16a34a?style=for-the-badge)](./package.json)
[![Tests](https://img.shields.io/badge/Tests-13%2F13%20passants-10b981?style=for-the-badge&logo=node.js&logoColor=white)](./tests/studio.test.mjs)
[![Déploiement](https://img.shields.io/badge/CI%2FCD-GitHub%20Actions-2088ff?style=for-the-badge&logo=githubactions&logoColor=white)](./.github/workflows/deploy.yml)
[![100% Local](https://img.shields.io/badge/Confidentialit%C3%A9-100%25%20Local-0f172a?style=for-the-badge)](#confidentialit%C3%A9-s%C3%A9curit%C3%A9-et-limites-connues)

[Accéder à l'application en ligne](https://matthiasmaj.github.io/excel-macro-studio/) • [Guide rapide](#-guide-dutilisation-rapide) • [Académie Excel](#-acad%C3%A9mie-excel--vba-parcours-a--b) • [Assistant & Générateur](#-assistant-et-g%C3%A9n%C3%A9rateur-dautomatisations) • [Audit & Analyseur](#-audit-vba-et-analyseur-de-classeurs) • [Architecture](#-architecture-du-projet)

</div>

---

## À propos

**Excel Macro Studio** est un environnement complet de travail et de formation professionnelle pour Microsoft Excel :
- **Apprendre Excel progressivement sans code** puis se spécialiser en **VBA** grâce à l'**Académie Excel** et au projet fil rouge **Entreprise NovaTech**.
- **Automatiser des besoins métier** via un Assistant en français courant, un catalogue de **129 opérations** (12 catégories), un compositeur de workflows et une bibliothèque de scénarios répartie sur **7 domaines professionnels**.
- **Générer le code adapté à l'environnement cible** : **VBA** structuré pour Excel Desktop (Windows / Mac) ou **Office Scripts TypeScript (`.osts`)** pour **Excel Web** (sans jamais mélanger les deux technologies).
- **Auditer et corriger du code VBA existant** (explication, erreurs probables, mauvaises pratiques, risques sécurité, optimisations et diff avant/après).

Toute l'analyse et la génération s'exécutent **100 % côté client dans le navigateur**.

---

## 🧭 Guide d'utilisation rapide

1. **Apprendre pas à pas** : ouvrez **Académie Excel (A & B)**, démarrez par le **Parcours A** si vous débutez sur Excel (aucune programmation requise) ou passez au **Parcours B** pour apprendre le VBA, et téléchargez les classeurs d'entraînement **NovaTech**.
2. **Adapter à votre fichier réel (recommandé)** : ouvrez **Analyseur Excel** et déposez un classeur (`.xlsx`, `.xlsm`, `.xls`, `.csv`) ou chargez un exemple — vos noms de feuilles et colonnes pré-remplissent ensuite automatiquement les formulaires du générateur.
3. **Générer une automatisation** :
   - **Assistant (Langage naturel)** : décrivez votre besoin en français, validez/modifiez la reformulation proposée, répondez aux éventuelles questions ciblées puis choisissez votre plateforme (**Windows**, **Mac** ou **Excel Web**).
   - **Mode rapide (Guidé)** & **Multi-opérations** : sélectionnez directement parmi les **129 opérations** ou enchaînez un pipeline personnalisé.
4. **Auditer une macro existante** : collez votre VBA dans **Audit & Correction VBA** pour inspecter erreurs, mauvaises pratiques, risques et récupérer une version corrigée avec diff Avant/Après.
5. **Installer dans Excel** :
   - **Excel Desktop (Windows / Mac)** : importez le fichier `.bas` dans l'éditeur VBA (`Alt + F11` / `Option + F11` → `Fichier > Importer`) ou collez le code, puis enregistrez le classeur au format **`.xlsm`**.
   - **Excel Web (Microsoft 365)** : sélectionnez **Excel Web (Office Scripts)**, récupérez le script **`.osts`**, puis collez-le dans **Automatisation > Nouveau script**.
6. **Recherche globale** : appuyez sur `/` depuis n'importe quelle vue pour retrouver instantanément une leçon, un scénario métier ou une opération VBA.

---

## 🎓 Académie Excel & VBA (Parcours A & B)

L'Académie distingue clairement savoir utiliser Excel et savoir l'automatiser en VBA :

### Parcours A — Maîtriser Excel (sans VBA · 5 niveaux · 15 chapitres)
* **Niveau 0 — Découverte** : interface et ruban, classeur/feuilles/cellules, saisie et formats de fichiers (`.xlsx`, `.xlsm`, `.csv`, PDF), navigation et raccourcis.
* **Niveau 1 — Fondamentaux** : types de données, formats numériques/dates/textes, références relatives et absolues (`$`), formules (`SOMME`, `MOYENNE`, `MIN`, `MAX`), logique `SI`, `ET`, `OU` et gestion des erreurs (`#DIV/0!`, `#NOM?`, `SIERREUR`).
* **Niveau 2 — Gestion des données** : tableaux structurés, filtres/tris/doublons, validation des données, mise en forme conditionnelle (MFC), `NB.SI`, `SOMME.SI`, `SOMME.SI.ENS`, `RECHERCHEX` + alternatives rétrocompatibles (`INDEX`/`EQUIV`), texte et dates (`NB.JOURS.OUVRES`).
* **Niveau 3 — Analyse** : Tableaux Croisés Dynamiques (TCD), Segments, graphiques, KPI, analyse comparative, tableaux de bord, consolidation multi-sources et introduction à **Power Query**.
* **Niveau 4 — Maîtrise professionnelle** : architecture de classeurs complexes, optimisation des formules (fonctions volatiles), contrôles de cohérence, mise en page de rapports/PDF et maintenance.

### Parcours B — Automatiser Excel avec VBA (4 niveaux · 12 chapitres enrichis)
* Conserve intégralement les **12 chapitres VBA** progressifs (premiers pas, cellules, variables, conditions, boucles, nettoyage/tri, dates/rapprochements, TCD/KPI, analyse, imports/fichiers, événements/PDF, architecture) et les enrichit de **12 modules complémentaires** : débogage/points d'arrêt (`F8`/`F9`), hiérarchie `Workbook`/`Worksheet`/`Range`, portée des variables, `Sub` vs `Function` (`ByVal`/`ByRef`), tableaux `Variant` en mémoire, `ListObject`, `Collection` vs `Scripting.Dictionary`, fonctions personnalisées UDF, optimisation des performances, interaction avec d'autres fichiers, `UserForm` et gestion avancée des erreurs (`Err.Number`, `Resume`).

### Pédagogie unifiée & Projet Fil Rouge NovaTech
* **12 sections homogènes par leçon** : objectif, mise en situation NovaTech, prérequis, explication théorique (tableaux visuels / grilles de cellules), vocabulaire, démonstration commentée, exercice guidé, exercice autonome, application fil rouge, validation des acquis, fiche résumé et tableau de compatibilité (**Windows / Mac / M365 / Excel Web**).
* **9 formats d'exercices interactifs** (`academyExercises.js`) avec indices progressifs et explication des erreurs fréquentes : QCM, Vrai/Faux, à compléter, ordre d'étapes, détection d'erreurs, association formule ↔ résultat, pratique sur tableau, problème métier et question ouverte.
* **Progression adaptative locale** (`academyProgress.js`) : suivi par parcours et niveau, tableau de bord personnel, reprise là où l'on s'est arrêté, maîtrise par compétence et proposition automatique d'exercices de consolidation en cas de difficultés répétées (sans jamais bloquer la navigation).
* **Classeurs NovaTech téléchargeables** : classeur de départ à **4 onglets** (`Entreprise_NovaTech_FilRouge.xlsx`) et classeur complet étendu à **7 onglets** (`Entreprise_NovaTech_Gestion_Complete.xlsx`).

---

## ✨ Assistant et générateur d'automatisations

* **Mode Assistant (Langage naturel local)** : décomposition structurée du besoin (résultat attendu, feuilles/données, opérations, risques, ambiguïtés), **reformulation éditable et validable** avant génération et questions ciblées pour ne jamais inventer vos noms d'onglets ou colonnes.
* **Contexte d'exécution cible** (`executionContext.js`) : Excel Windows Desktop, Excel pour Mac, ou **Excel Web (Office Scripts `.osts` TypeScript)**, volume de données, fréquence et niveau de commentaires configurables (**Débutant / Intermédiaire / Avancé**).

### VBA vs Office Scripts selon votre environnement

| Critère | VBA (`.bas` / `.xlsm`) | Office Scripts (`.osts`) |
| :--- | :--- | :--- |
| **Environnement d'exécution** | Excel Desktop (Windows / Mac) | Excel Web et Excel Desktop M365 (onglet Automatisation) |
| **Langage utilisé** | Visual Basic for Applications (`.bas` / `.xlsm`) | TypeScript / JavaScript (`.osts`) |
| **Accès disque local / Outlook / Word** | Oui sur Windows (`FSO`, `Dir`, COM) | Non : sandbox cloud (nécessite Power Automate pour fichiers/e-mails) |
| **Événements temps réel (`Worksheet_Change`)** | Oui nativement | Non (déclenchement manuel, bouton ou Power Automate) |

* **Mode Rapide (Guidé)** & **Générateur Multi-opérations (Workflows)** couvrant **129 opérations** sur **12 catégories (A à L)** :

| Code | Catégorie | Opérations | Code | Catégorie | Opérations |
| :---: | :--- | :---: | :---: | :--- | :---: |
| **A** | Gestion des cellules et des données | 15 | **G** | Gestion administrative et RH | 9 |
| **B** | Gestion des feuilles et classeurs | 13 | **H** | Gestion commerciale et stocks | 9 |
| **C** | Traitement et analyse des données | 15 | **I** | Génération de documents et exports | 9 |
| **D** | Mise en forme et présentation | 10 | **J** | Automatisation avancée | 13 |
| **E** | Gestion des dates et du temps | 10 | **K** | Graphiques et tableaux de bord | 7 |
| **F** | Gestion financière et comptable | 11 | **L** | Importation et communication | 8 |

* **Bibliothèque de scénarios métier (7 domaines)** : *Comptabilité et gestion*, *Ressources humaines*, *Stocks et logistique*, *Administration*, *Analyse et reporting*, *Maintenance et production*, *Gestion commerciale*.
* **Mes créations & Historique** : favoris, suivi de statut (*Généré → Analysé → À tester → Validé dans Excel*), gestion multi-versions avec comparateur Avant/Après, et exports `.bas`, `.osts`, `.txt`, kit `.xlsx`, documentation imprimable/PDF HTML et JSON.

---

## 🛡️ Audit VBA et Analyseur de classeurs

* **Audit & Correction de macros VBA (`vbaAuditor.js`)** : collez une macro existante pour obtenir son résumé en français, la liste des procédures, les erreurs probables (`Option Explicit` absent, `As Integer` sur lignes, boucle `.Delete` sans `Step -1`), les mauvaises pratiques (`.Select`/`Selection`, `ActiveSheet`), les dépendances, les alertes sécurité et une **version corrigée avec diff Avant/Après** sans supprimer vos fonctionnalités.
* **Analyseur de classeur Excel (`excelAnalyzer.js`)** : lecture locale de `.xlsx`, `.xlsm`, `.xls` et `.csv` via SheetJS sans exécuter aucune macro. Identifie feuilles, dimensions, en-têtes recalculables, colonnes/types, plages nommées, formules et erreurs (`#VALUE!`, `#REF!`...), cellules fusionnées, relations inter-feuilles et incohérences structurelles (sans jamais modifier ni écraser le fichier d'origine).

---

## 🗂️ Architecture du projet

```text
excel-macro-studio/
├── .github/workflows/deploy.yml       # CI/CD : tests, build et déploiement GitHub Pages
├── src/
│   ├── catalog/                       # 12 catégories (A à L) et 129 opérations (ops/catA..catL)
│   ├── core/
│   │   ├── executionContext.js        # Paramétrage plateforme (Win/Mac/Web), volume et commentaires
│   │   ├── security.js                # Sanitisation DOM + analyse statique des risques VBA
│   │   ├── store.js                   # État central, favoris, versions et persistance localStorage
│   │   ├── theme.js                   # Thème clair / sombre / système
│   │   └── validation.js              # Validation des paramètres formulaires
│   ├── engine/
│   │   ├── officeScriptsGenerator.js  # Générateur TypeScript (.osts) dédié à Excel Web
│   │   ├── vbaAuditor.js              # Audit statique, correction non destructive et diff ligne par ligne
│   │   ├── vbaBeginnerAnnotator.js    # Niveaux de commentaires (Débutant / Intermédiaire / Avancé)
│   │   ├── vbaGenerator.js            # Générateur principal + routage Office Scripts
│   │   ├── vbaHighlighter.js          # Coloration syntaxique
│   │   ├── vbaRefiner.js              # Améliorations ciblées (confirmation, sauvegarde, journal, Mac)
│   │   ├── vbaUtils.js                # Helpers VBA échappés
│   │   └── workflowComposer.js        # Composition multi-opérations (VBA ou Office Scripts)
│   ├── services/
│   │   ├── academyExercises.js        # Moteur des 9 formats d'exercices et remédiation
│   │   ├── academyProgress.js         # Progression adaptative persistante et maîtrise par compétence
│   │   ├── curriculumData.js          # Parcours B (12 chapitres VBA d'origine préservés)
│   │   ├── excelAnalyzer.js           # Analyse locale XLSX/XLSM/CSV + incohérences & plages nommées
│   │   ├── excelCurriculumData.js     # Parcours A (5 niveaux, 15 chapitres Maîtriser Excel)
│   │   ├── exportService.js           # Exports .bas, .osts, .txt, kit .xlsx et documentation imprimable
│   │   ├── naturalLanguageEngine.js   # Moteur sémantique local + décomposition & reformulation
│   │   ├── templatesLibrary.js        # Scénarios professionnels classés sur 7 domaines métier
│   │   ├── trainingWorkbook.js        # Classeurs NovaTech (départ 4 onglets + complet 7 onglets)
│   │   └── vbaCurriculumEnrichment.js # Enrichissement 12 sections + 12 modules complémentaires VBA
│   ├── styles/                        # tokens.css, base.css, layout.css, components.css
│   ├── ui/
│   │   ├── analyzerView.js            # Vue Analyseur Excel
│   │   ├── auditorView.js             # Vue Audit & Correction VBA (avec diff)
│   │   ├── catalogView.js             # Navigation du catalogue (12 catégories)
│   │   ├── exerciseWidget.js          # Composant des 9 types d'exercices interactifs
│   │   ├── formEngine.js              # Formulaires dynamiques auto-alimentés par le classeur analysé
│   │   ├── freeView.js                # Mode Assistant (langage naturel + reformulation + contexte)
│   │   ├── guidedView.js              # Mode Rapide (génération guidée par formulaire)
│   │   ├── historyView.js             # Mes créations, favoris, versions et comparateur
│   │   ├── homeView.js                # Accueil et accès rapides
│   │   ├── libraryView.js             # Bibliothèque de scénarios (7 domaines métier)
│   │   ├── resultPanel.js             # Panneau de résultat, cycle de validation, commentaires et exports
│   │   ├── tutorialView.js            # Académie Excel & VBA + Tableau de bord adaptatif
│   │   └── workflowView.js            # Compositeur multi-opérations
│   └── main.js                        # Point d'entrée, navigation et recherche globale unifiée
├── tests/studio.test.mjs              # Suite de tests automatisés (node --test)
├── index.html
├── package.json
└── vite.config.js
```

---

## 🔒 Confidentialité, sécurité et limites connues

* **Confidentialité 100 % locale** : aucun backend ni appel API d'IA externe ; aucune clé d'API n'est stockée ni requise. Aucun fichier Excel ni extrait de code ne quitte votre navigateur.
* **Aucune exécution automatique de VBA** : le code généré ou collé n'est jamais exécuté dans le navigateur. La distinction entre code généré, analysé statiquement et validé par vos soins dans Excel reste explicite.
* **Limites techniques documentées honnêtement** :
  * **Fichiers `.xlsm`** : pour des raisons liées au format binaire propriétaire Microsoft (`vbaProject.bin`), l'application ne compile pas directement un binaire `.xlsm` exécutable côté navigateur et ne lit pas le code VBA compilé à l'intérieur d'un `.xlsm` déposé dans l'analyseur. Elle fournit les modules `.bas` prêts à importer et un kit `.xlsx` pas à pas.
  * **Excel Web** : VBA ne s'exécute pas dans Excel Web ; l'application génère alors un **Office Script TypeScript (`.osts`)** et signale lorsque certaines opérations (accès fichiers locaux, Outlook, PDF disque) requièrent Power Automate ou Excel Desktop.
  * **Excel pour Mac** : les objets COM Windows (`Outlook.Application`, `Word.Application`, `Scripting.FileSystemObject`, `Scripting.Dictionary`, `ADODB`) sont signalés comme incompatibles Mac lors de l'analyse.

---

## 🚀 Installation, tests et déploiement

```bash
# 1. Cloner le dépôt
git clone https://github.com/MatthiasMaj/excel-macro-studio.git
cd excel-macro-studio

# 2. Installer les dépendances
npm install

# 3. Lancer le serveur de développement Vite (http://localhost:5173)
npm run dev

# 4. Exécuter les tests automatisés (13/13)
npm test

# 5. Construire le bundle de production dans dist/
npm run build

# 6. Prévisualiser le build de production localement
npm run preview
```

### Déploiement continu (GitHub Pages)
Le workflow [`.github/workflows/deploy.yml`](./.github/workflows/deploy.yml) exécute automatiquement `npm ci`, `npm test`, `npm run build` et déploie `dist/` sur GitHub Pages à chaque push sur `main` ou `master`.

Application en ligne : **https://matthiasmaj.github.io/excel-macro-studio/**

### Stack technique
* **JavaScript ES Modules** (Vanilla JS modulaire, architecture client sans dépendance serveur)
* **Vite** (`^6.2.0`) avec `base: './'` pour compatibilité directe **GitHub Pages**
* **SheetJS (`xlsx` `^0.18.5`)** pour l'analyse et la création locale de classeurs `.xlsx`
* **`node --test`** pour la suite de tests automatisés

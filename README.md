# 📄 Extracteur PDF vers Excel — Bank Statement Converter

> Application web de conversion automatique de relevés bancaires PDF en fichiers Excel structurés et formatés.

[![Streamlit App](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://mani26368-extracteur-pdf-excel-app-xxxxx.streamlit.app)
![Python](https://img.shields.io/badge/Python-3.10+-blue.svg)
![Streamlit](https://img.shields.io/badge/Streamlit-1.x-red.svg)
![License](https://img.shields.io/badge/License-MIT-green.svg)

---

## 📋 Table des matières

- [Contexte et problématique](#-contexte-et-problématique)
- [Solution développée](#-solution-développée)
- [Démo](#-démo)
- [Fonctionnalités](#-fonctionnalités)
- [Architecture du projet](#-architecture-du-projet)
- [Formats supportés](#-formats-supportés)
- [Format de sortie Excel](#-format-de-sortie-excel)
- [Technologies utilisées](#-technologies-utilisées)
- [Installation locale](#-installation-locale)
- [Déploiement](#-déploiement)
- [Exemples de résultats](#-exemples-de-résultats)
- [Auteur](#-auteur)

---

## 🎯 Contexte et problématique

Dans le cadre du stage au sein du département comptabilité de **Djibouti Ports Corridor Road S.A (DPCR)**, l'équipe traitait manuellement des relevés bancaires reçus en format PDF provenant de plusieurs banques :

- **Banque Nationale de Djibouti (BND)**
- **Exim Bank Djibouti**

### Le problème

Le processus manuel présentait plusieurs inconvénients majeurs :

| Problème | Impact |
|----------|--------|
| Recopie manuelle des transactions | Perte de temps considérable |
| Risque d'erreurs de saisie | Données comptables incorrectes |
| Formats différents selon les banques | Traitement non standardisé |
| Aucune distinction Débit / Crédit automatique | Analyse financière difficile |
| Pas de mise en forme standardisée | Rapports non uniformes |

### L'objectif

Développer une application capable de :
1. **Lire automatiquement** n'importe quel relevé bancaire PDF
2. **Extraire intelligemment** toutes les transactions
3. **Générer un fichier Excel** propre, formaté et standardisé
4. **Être accessible** facilement par toute l'équipe

---

## 💡 Solution développée

L'application utilise une approche en **3 phases** pour traiter les PDFs :

```
PDF Source
    │
    ▼
┌─────────────────────────────────────┐
│  PHASE 1 : Détection automatique    │
│  - Identification de la banque      │
│  - Extraction de l'en-tête          │
│  - Détection du format de date      │
└─────────────────────────────────────┘
    │
    ▼
┌─────────────────────────────────────┐
│  PHASE 2 : Extraction des données   │
│  - Méthode 1 : Extraction tableau   │
│  - Méthode 2 : Extraction mot-à-mot │
│  - Calcul Débit / Crédit auto       │
└─────────────────────────────────────┘
    │
    ▼
┌─────────────────────────────────────┐
│  PHASE 3 : Génération Excel         │
│  - En-tête formaté avec couleurs    │
│  - Colonnes standardisées           │
│  - Mise en forme professionnelle    │
└─────────────────────────────────────┘
    │
    ▼
Excel Structuré ✅
```

---

## 🌐 Démo

🔗 **Application en ligne** : [https://mani26368-extracteur-pdf-excel-app.streamlit.app](https://mani26368-extracteur-pdf-excel-app.streamlit.app)

![Interface de l'application](https://i.imgur.com/placeholder.png)

---

## ✨ Fonctionnalités

### Détection automatique de la banque
L'application identifie automatiquement la banque émettrice du relevé sans configuration manuelle :
- ✅ Banque Nationale de Djibouti
- ✅ Exim Bank Djibouti
- ✅ Toute autre banque avec un PDF électronique

### Extraction intelligente
- ✅ **Deux stratégies d'extraction** : par tableau (PDFs avec bordures) et par analyse de mots (PDFs sans bordures)
- ✅ **Multi-pages** : traite automatiquement tous les pages du document
- ✅ **Détection des dates** : supporte les formats DD-MM-YYYY, DD-MMM-YYYY, YYYY-MM-DD
- ✅ **Extraction de l'en-tête** : nom du client, numéro de compte, période, devise, etc.

### Calcul automatique Débit / Crédit
L'application détermine intelligemment si chaque transaction est un débit ou un crédit en comparant les soldes consécutifs :
- Solde **augmente** → **Crédit** (entrée d'argent)
- Solde **diminue** → **Débit** (sortie d'argent)

### Excel formaté et professionnel
- ✅ En-tête coloré avec les informations du compte
- ✅ Colonnes aux largeurs optimisées
- ✅ Alternance de couleurs sur les lignes
- ✅ Format standardisé identique pour toutes les banques

---

## 📁 Architecture du projet

```
extracteur-pdf-excel/
│
├── app.py                 # Application Streamlit (interface web)
├── requirements.txt       # Dépendances Python
└── README.md              # Documentation du projet
```

### Description des fichiers

| Fichier | Rôle |
|---------|------|
| `app.py` | Point d'entrée de l'application. Contient l'interface Streamlit ET le moteur d'extraction PDF |
| `requirements.txt` | Liste des bibliothèques Python nécessaires |
| `README.md` | Documentation complète du projet |

---

## 🏦 Formats supportés

### Format 1 — Banque Nationale de Djibouti
```
Date         | Description          | Check No | Branch      | Debit     | Credit    | Balance
01-12-2025   | DPCS Transfer...     |          | Main Branch |           | 3,600.00  | 48,920.50
02-12-2025   | Wire Transfer...     | 1045     | Main Branch | 12,500.00 |           | 40,020.50
```
- Format de date : `DD-MM-YYYY`
- Colonne Check No pour les chèques
- Colonne Branch (Main Branch)

### Format 2 — Exim Bank Djibouti
```
Txn. Date    | Value Date  | Description          | Ex-ref no | Txn. Ref No | Debit | Credit   | Balance
Apr 29, 2026 | Apr 29, 2026| DPCS Refno:PMT-...   |           | 1/190/2     |       | 3,600.00 |
```
- Format de date : `DD-MMM-YYYY`
- Colonne Value Date
- Colonne Txn. Ref No

### Format générique
Tout PDF bancaire électronique avec des tableaux structurés est supporté grâce à la détection automatique des en-têtes de colonnes.

---

## 📊 Format de sortie Excel

Quel que soit le PDF d'entrée, l'Excel généré contient toujours :

### Section En-tête (si présente dans le PDF)
| Champ | Exemple |
|-------|---------|
| Nom de la banque | BANQUE NATIONALE DE DJIBOUTI |
| Customer Name | Messrs DJIBOUTI PORTS CORRIDOR ROAD S.A |
| Période | 01/12/2025 to 15/12/2025 |
| Account No | 01003597737 |
| Currency | USD |

### Section Transactions
| Txn. Date | Value Date | Description | Ex-ref no | Txn. Ref No | Debit | Credit | Balance |
|-----------|------------|-------------|-----------|-------------|-------|--------|---------|
| 01-12-2025 | | Monthly Charge Fee | | | | 700.00 | 45,320.50 |
| 02-12-2025 | | Wire Transfer SWIFT... | | 1045 | 12,500.00 | | 40,020.50 |

---

## 🛠 Technologies utilisées

| Technologie | Version | Utilisation |
|-------------|---------|-------------|
| **Python** | 3.10+ | Langage principal |
| **Streamlit** | 1.x | Interface web |
| **pdfplumber** | 0.x | Extraction de texte depuis les PDFs |
| **openpyxl** | 3.x | Création et formatage des fichiers Excel |
| **pandas** | 2.x | Manipulation des données tabulaires |
| **re** | stdlib | Expressions régulières pour la détection des dates et montants |

### Pourquoi ces choix ?

- **pdfplumber** : meilleure bibliothèque Python pour extraire du texte structuré depuis des PDFs, avec support des tableaux et des coordonnées de mots
- **Streamlit** : framework web Python simple à déployer, idéal pour les applications data
- **openpyxl** : permet un contrôle total sur la mise en forme Excel (couleurs, polices, bordures)

---

## 💻 Installation locale

### Prérequis
- Python 3.10 ou supérieur
- pip

### Étapes

```bash
# 1. Cloner le repository
git clone https://github.com/Mani26368/extracteur-pdf-excel.git
cd extracteur-pdf-excel

# 2. Créer un environnement virtuel (recommandé)
python -m venv venv
source venv/bin/activate  # Linux/Mac
venv\Scripts\activate     # Windows

# 3. Installer les dépendances
pip install -r requirements.txt

# 4. Lancer l'application
streamlit run app.py
```

L'application s'ouvre automatiquement sur `http://localhost:8501`

---

## 🚀 Déploiement

L'application est déployée sur **Streamlit Cloud** (gratuit) :

1. Fork ou clone ce repository sur votre GitHub
2. Connectez-vous sur [share.streamlit.io](https://share.streamlit.io)
3. Cliquez **New app** → sélectionnez votre repo
4. Main file : `app.py` → **Deploy**

---

## 📸 Exemples de résultats

### Avant (PDF brut)
```
01-12-2025  DPCS - From DPCS- Djibouti - DPCS Transfer Ref: TRF/2025/DEC/001234
            Beneficiary: DPCR SA Corridor Road Project Fund     Main Branch  3,600.00  48,920.50
```

### Après (Excel généré)
| Txn. Date | Description | Txn. Ref No | Debit | Credit | Balance |
|-----------|-------------|-------------|-------|--------|---------|
| 01-12-2025 | DPCS - From DPCS- Djibouti - DPCS Transfer Ref: TRF/2025/DEC/001234 Beneficiary: DPCR SA... | | | 3,600.00 | 48,920.50 |

---

## 👤 Auteur

**Développé dans le cadre d'un stage à DPCR — Djibouti Ports Corridor Road S.A**

- GitHub : [@Mani26368](https://github.com/Mani26368)

---

## 📄 Licence

Ce projet est sous licence MIT — voir le fichier [LICENSE](LICENSE) pour plus de détails.

---

*Développé avec ❤️ pour automatiser le traitement comptable des relevés bancaires à Djibouti*

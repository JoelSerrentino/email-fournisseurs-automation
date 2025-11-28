# Email Fournisseurs Automation

## Description
Application d'automatisation pour le traitement des emails fournisseurs, conçue pour le **Service des Finances**. Elle permet de filtrer les emails en fonction de mots clés, de les déplacer vers un dossier Outlook, de fusionner le contenu et les pièces jointes en fichiers PDF, et d'appliquer une catégorie Outlook après traitement.

## 🚀 Installation rapide (Exécutable)

**Aucune installation Python requise !**

1. Téléchargez `Email-Fournisseurs-Automation.exe` depuis le dossier `dist/`
2. Double-cliquez pour lancer l'application

### Prérequis sur le poste cible
- ✅ Windows 10/11
- ✅ Microsoft Outlook installé et configuré
- ⚪ Microsoft Word/Excel (optionnel, pour conversion des pièces jointes Office)

## Fonctionnalités
- 📬 **Sélection de la boîte aux lettres** Outlook via interface graphique
- 🔍 **Filtrage des emails** par mots clés dans l'objet
- 📅 **Filtrage par date** avec sélecteur de calendrier (période Du/Au)
- 📁 **Déplacement automatique** des emails vers un dossier Outlook choisi
- 📄 **Fusion en PDF** : emails et pièces jointes combinés en un seul fichier
- 🏷️ **Catégorisation automatique** avec couleurs (vert = succès, rouge = erreur)
- 💾 **Sauvegarde des paramètres** pour une réutilisation rapide
- 📋 **Journal d'activité** en temps réel avec causes d'erreurs détaillées
- 📊 **Barre de progression** avec statistiques (succès/échecs)
- ⏹️ **Arrêt du traitement** à tout moment
- 🔄 **Traitement asynchrone** (interface non bloquée)

### Types de pièces jointes supportés
| Type | Extensions | Méthode de conversion |
|------|------------|----------------------|
| PDF | `.pdf` | Fusion directe |
| Images | `.png`, `.jpg`, `.jpeg`, `.gif`, `.bmp`, `.tiff`, `.webp` | Pillow |
| Word | `.doc`, `.docx` | Microsoft Word (COM) |
| Excel | `.xls`, `.xlsx` | Microsoft Excel (COM) |
| Texte | `.txt`, `.csv`, `.log` | ReportLab |

## Interface graphique
L'application dispose d'une interface moderne **style Windows 11**, adaptée pour un environnement professionnel :
- Thème clair avec accents bleu Microsoft (#0078d4)
- Cartes avec barres d'accent colorées (or, vert, bleu finance)
- Effets de survol et focus
- Design responsive (s'adapte à toutes les tailles de fenêtre)
- Barre de progression animée avec statistiques en temps réel

## Structure du projet
```
email-fournisseurs-automation/
├── src/
│   ├── main.py                # Point d'entrée de l'application
│   ├── email_processor.py     # Orchestration du traitement des emails
│   ├── pdf_generator.py       # Génération et fusion de fichiers PDF
│   ├── outlook_handler.py     # Gestion des interactions avec Outlook (COM)
│   ├── gui/
│   │   ├── __init__.py
│   │   └── main_window.py     # Interface graphique Windows 11 (Tkinter)
│   └── utils/
│       ├── __init__.py
│       ├── sanitize.py        # Nettoyage de texte et noms de fichiers
│       └── logger.py          # Journalisation avec niveaux et callbacks
├── config/
│   └── gui_settings.json      # Paramètres sauvegardés de l'interface
├── logs/                      # Fichiers de log générés
├── tests/
│   ├── __init__.py
│   ├── test_email_processor.py
│   └── test_pdf_generator.py
├── dist/                      # Exécutable généré
│   └── Email-Fournisseurs-Automation.exe
├── build_installer.ps1        # Script de build PowerShell
├── Email-Fournisseurs-Automation.spec  # Configuration PyInstaller
├── requirements.txt           # Dépendances Python
└── README.md
```

## Prérequis (pour le développement)
- Python 3.10 ou supérieur
- Microsoft Outlook installé et configuré
- Windows 10/11

## Installation (pour le développement)

1. **Cloner le dépôt**
   ```bash
   git clone <url_du_dépôt>
   cd email-fournisseurs-automation
   ```

2. **Créer un environnement virtuel** (recommandé)
   ```bash
   python -m venv venv
   .\venv\Scripts\Activate.ps1  # Windows PowerShell
   ```

3. **Installer les dépendances**
   ```bash
   pip install -r requirements.txt
   ```

## Utilisation

1. **Lancer l'application**
   ```bash
   python src/main.py
   ```

2. **Configurer les paramètres** via l'interface graphique :
   - Sélectionner la boîte aux lettres Outlook
   - Choisir le dossier de destination Outlook
   - Définir la catégorie à appliquer après traitement
   - Saisir les mots clés de filtrage (séparés par des virgules)
   - Sélectionner une période de dates (optionnel) : cliquez sur ▼ pour ouvrir le calendrier
   - Sélectionner le dossier de sortie pour les PDF

3. **Sauvegarder les paramètres** (optionnel) pour les réutiliser ultérieurement

4. **Lancer le traitement** en cliquant sur le bouton "🚀 Lancer le traitement"

5. **Suivre la progression** via la barre de progression et les statistiques en temps réel

6. **Arrêter le traitement** si nécessaire avec le bouton "⏹ Arrêter"

## Dépendances principales
| Package | Version | Description |
|---------|---------|-------------|
| `pywin32` | ≥306 | Interaction avec Microsoft Outlook, Word, Excel via COM |
| `reportlab` | ≥4.0.0 | Génération de PDF depuis le contenu des emails |
| `PyPDF2` | ≥3.0.0 | Manipulation et fusion de fichiers PDF |
| `Pillow` | ≥10.0.0 | Conversion d'images en PDF |
| `tkcalendar` | ≥1.6.1 | Sélecteur de date avec calendrier intégré |

## 📦 Créer l'exécutable

### Méthode rapide (PowerShell)
```powershell
.\build_installer.ps1
```

### Méthode manuelle
```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name "Email-Fournisseurs-Automation" --add-data "config;config" src/main.py
```

L'exécutable sera créé dans le dossier `dist/`.

## Architecture technique

### Modules principaux

- **`email_processor.py`** : Orchestrateur principal avec callbacks de progression, gestion des statistiques et support d'arrêt gracieux
- **`outlook_handler.py`** : Wrapper COM pour Outlook avec classe `EmailItem` et `OutlookHandler`
- **`pdf_generator.py`** : Génération de PDF avec `reportlab`, fusion avec `PyPDF2`, conversion d'images avec `Pillow`
- **`logger.py`** : Système de log avec niveaux (DEBUG, INFO, WARNING, ERROR, SUCCESS), écriture fichier et callback GUI

### Traitement asynchrone
Le traitement des emails s'exécute dans un thread séparé pour ne pas bloquer l'interface graphique. Les mises à jour de progression sont transmises via callbacks thread-safe.

### Catégories Outlook
L'application crée automatiquement les catégories avec les couleurs appropriées :
- **Succès** : Catégorie verte (configurable dans l'interface)
- **Erreur** : Catégorie rouge "Erreur traitement"

## 🐛 Dépannage

| Problème | Solution |
|----------|----------|
| "pywin32 n'est pas installé" | `pip install pywin32` |
| "Connexion Outlook échouée" | Vérifier qu'Outlook est ouvert et configuré |
| "Dossier introuvable" | Vérifier le chemin du dossier Outlook |
| Conversion Word/Excel échoue | Vérifier que Microsoft Office est installé |
| L'exécutable ne démarre pas | Exécuter en tant qu'administrateur |

## Licence
Ce projet est sous licence MIT.
# Script de build complet pour Email Fournisseurs Automation
# Crée l'exécutable ET l'installateur

param(
    [switch]$SkipInstaller,
    [switch]$Clean
)

$ErrorActionPreference = "Continue"

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host " Build Installateur Complet" -ForegroundColor Cyan
Write-Host " Email Fournisseurs Automation" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Fonction pour afficher les étapes
function Write-Step {
    param($Step, $Total, $Message)
    Write-Host "[$Step/$Total] $Message" -ForegroundColor Yellow
}

function Write-OK {
    param($Message)
    Write-Host "       $Message" -ForegroundColor Green
}

function Write-Err {
    param($Message)
    Write-Host "       $Message" -ForegroundColor Red
}

# Étape 1 : Nettoyage si demandé
if ($Clean) {
    Write-Step 1 4 "Nettoyage complet..."
    Remove-Item -Recurse -Force -ErrorAction SilentlyContinue "build"
    Remove-Item -Recurse -Force -ErrorAction SilentlyContinue "dist"
    Remove-Item -Recurse -Force -ErrorAction SilentlyContinue "installer_output"
    Remove-Item -Force -ErrorAction SilentlyContinue "*.spec"
    Write-OK "OK"
    Write-Host ""
}

# Étape 2 : Vérifier PyInstaller
Write-Step 1 4 "Vérification de PyInstaller..."
$pyinstaller = pip show pyinstaller 2>$null
if (-not $pyinstaller) {
    Write-Host "       Installation de PyInstaller..." -ForegroundColor Gray
    pip install pyinstaller
}
Write-OK "OK"
Write-Host ""

# Étape 3 : Créer le dossier assets
if (-not (Test-Path "assets")) {
    New-Item -ItemType Directory -Path "assets" -Force | Out-Null
}

# Étape 4 : Créer l'exécutable
Write-Step 2 4 "Création de l'exécutable (cela peut prendre quelques minutes)..."
Remove-Item -Recurse -Force -ErrorAction SilentlyContinue "build"
Remove-Item -Recurse -Force -ErrorAction SilentlyContinue "dist"

# Exécuter PyInstaller
pyinstaller --onefile --windowed `
    --name "Email-Fournisseurs-Automation" `
    --add-data "config;config" `
    --hidden-import win32com.client `
    --hidden-import win32com.server `
    --hidden-import pythoncom `
    --hidden-import pywintypes `
    --hidden-import tkinter `
    --hidden-import tkinter.ttk `
    --hidden-import tkinter.filedialog `
    --hidden-import tkinter.messagebox `
    --hidden-import PIL `
    --hidden-import PIL.Image `
    --hidden-import PIL.ImageDraw `
    --hidden-import reportlab `
    --hidden-import reportlab.lib `
    --hidden-import reportlab.lib.pagesizes `
    --hidden-import reportlab.lib.styles `
    --hidden-import reportlab.platypus `
    --hidden-import reportlab.pdfgen `
    --hidden-import PyPDF2 `
    src/main.py

if (-not (Test-Path "dist\Email-Fournisseurs-Automation.exe")) {
    Write-Err "La création de l'exécutable a échoué !"
    Write-Host ""
    Read-Host "Appuyez sur Entrée pour fermer"
    exit 1
}

$exeSize = [math]::Round((Get-Item "dist\Email-Fournisseurs-Automation.exe").Length / 1MB, 2)
Write-OK "OK ($exeSize Mo)"
Write-Host ""

# Étape 5 : Vérifier Inno Setup
if (-not $SkipInstaller) {
    Write-Step 3 4 "Recherche d'Inno Setup..."
    
    $innoPath = $null
    $innoPaths = @(
        "C:\Program Files (x86)\Inno Setup 6\ISCC.exe",
        "C:\Program Files\Inno Setup 6\ISCC.exe",
        "C:\Program Files (x86)\Inno Setup 5\ISCC.exe"
    )
    
    foreach ($path in $innoPaths) {
        if (Test-Path $path) {
            $innoPath = $path
            break
        }
    }
    
    if (-not $innoPath) {
        Write-Host ""
        Write-Host "========================================" -ForegroundColor Yellow
        Write-Host " Inno Setup non trouvé" -ForegroundColor Yellow
        Write-Host "========================================" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "L'exécutable a été créé avec succès :" -ForegroundColor White
        Write-Host "  dist\Email-Fournisseurs-Automation.exe" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "Pour créer l'installateur :" -ForegroundColor White
        Write-Host "  1. Téléchargez Inno Setup : https://jrsoftware.org/isdl.php" -ForegroundColor Gray
        Write-Host "  2. Installez-le" -ForegroundColor Gray
        Write-Host "  3. Relancez ce script" -ForegroundColor Gray
        Write-Host ""
    } else {
        Write-OK "Trouvé: $innoPath"
        Write-Host ""
        
        Write-Step 4 4 "Création de l'installateur..."
        
        # Créer le dossier de sortie
        if (-not (Test-Path "installer_output")) {
            New-Item -ItemType Directory -Path "installer_output" -Force | Out-Null
        }
        
        # Compiler si installer.iss existe
        if (Test-Path "installer.iss") {
            & $innoPath "installer.iss"
            
            $installerPath = Get-ChildItem "installer_output\*.exe" -ErrorAction SilentlyContinue | Select-Object -First 1
            
            if ($installerPath) {
                $installerSize = [math]::Round($installerPath.Length / 1MB, 2)
                Write-OK "OK ($installerSize Mo)"
            } else {
                Write-Err "La création de l'installateur a échoué"
            }
        } else {
            Write-Host "       Fichier installer.iss non trouvé" -ForegroundColor Gray
        }
    }
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host " BUILD TERMINÉ !" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""
Write-Host "Exécutable : dist\Email-Fournisseurs-Automation.exe" -ForegroundColor Cyan
Write-Host ""

Read-Host "Appuyez sur Entrée pour fermer"

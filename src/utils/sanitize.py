"""
Module de nettoyage et validation de texte pour l'application Email Fournisseurs Automation.
"""

import re
import unicodedata
from typing import Optional
from datetime import datetime


# Caractères interdits dans les noms de fichiers Windows
FORBIDDEN_CHARS = ['<', '>', ':', '"', '/', '\\', '|', '?', '*']

# Noms réservés Windows
RESERVED_NAMES = [
    'CON', 'PRN', 'AUX', 'NUL',
    'COM1', 'COM2', 'COM3', 'COM4', 'COM5', 'COM6', 'COM7', 'COM8', 'COM9',
    'LPT1', 'LPT2', 'LPT3', 'LPT4', 'LPT5', 'LPT6', 'LPT7', 'LPT8', 'LPT9'
]


def sanitize_text(text: str, replacement: str = '_') -> str:
    """
    Nettoie le texte pour une utilisation sûre dans les noms de fichiers.
    
    Args:
        text: Texte à nettoyer
        replacement: Caractère de remplacement pour les caractères interdits
    
    Returns:
        Texte nettoyé
    """
    if not text:
        return ''
    
    # Remplacer les caractères interdits
    for char in FORBIDDEN_CHARS:
        text = text.replace(char, replacement)
    
    # Supprimer les caractères de contrôle
    text = ''.join(char for char in text if unicodedata.category(char) != 'Cc')
    
    # Supprimer les espaces multiples
    text = re.sub(r'\s+', ' ', text)
    
    # Supprimer les points/espaces en début et fin
    text = text.strip('. ')
    
    return text


def sanitize_filename(filename: str, max_length: int = 200) -> str:
    """
    Nettoie un nom de fichier pour le système de fichiers Windows.
    
    Args:
        filename: Nom de fichier à nettoyer
        max_length: Longueur maximale du nom de fichier
    
    Returns:
        Nom de fichier nettoyé et sécurisé
    """
    if not filename:
        return f"fichier_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    
    # Séparer le nom et l'extension
    if '.' in filename:
        name, ext = filename.rsplit('.', 1)
        ext = '.' + sanitize_text(ext)
    else:
        name = filename
        ext = ''
    
    # Nettoyer le nom
    name = sanitize_text(name)
    
    # Vérifier les noms réservés Windows
    if name.upper() in RESERVED_NAMES:
        name = f"_{name}_"
    
    # Tronquer si nécessaire
    max_name_length = max_length - len(ext)
    if len(name) > max_name_length:
        name = name[:max_name_length]
    
    # S'assurer qu'il y a un nom
    if not name:
        name = f"fichier_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    
    return name + ext


def extract_company_name(email_address: str) -> str:
    """
    Extrait le nom de l'entreprise à partir d'une adresse email.
    
    Args:
        email_address: Adresse email du type "nom@entreprise.com"
    
    Returns:
        Nom de l'entreprise nettoyé
    """
    if not email_address or '@' not in email_address:
        return "Inconnu"
    
    # Extraire le domaine
    domain = email_address.split('@')[1]
    
    # Enlever l'extension (.com, .fr, etc.)
    company = domain.rsplit('.', 1)[0]
    
    # Mettre en majuscule la première lettre
    company = company.capitalize()
    
    return sanitize_text(company)


def extract_sender_name(sender: str) -> str:
    """
    Extrait le nom de l'expéditeur depuis le champ "From".
    
    Args:
        sender: Champ expéditeur (peut être "Nom <email@domain.com>" ou juste l'email)
    
    Returns:
        Nom de l'expéditeur nettoyé
    """
    if not sender:
        return "Inconnu"
    
    # Si format "Nom <email@domain.com>"
    if '<' in sender and '>' in sender:
        name = sender.split('<')[0].strip()
        if name:
            return sanitize_text(name)
        # Sinon utiliser l'email
        email = sender.split('<')[1].split('>')[0]
        return extract_company_name(email)
    
    # Si juste une adresse email
    if '@' in sender:
        return extract_company_name(sender)
    
    return sanitize_text(sender)


def format_date_for_filename(date_obj: Optional[datetime] = None) -> str:
    """
    Formate une date pour utilisation dans un nom de fichier.
    
    Args:
        date_obj: Objet datetime (utilise la date actuelle si None)
    
    Returns:
        Date formatée (YYYYMMDD)
    """
    if date_obj is None:
        date_obj = datetime.now()
    
    return date_obj.strftime('%Y%m%d')


def generate_pdf_filename(sender: str, date_obj: Optional[datetime] = None, 
                          subject: str = "") -> str:
    """
    Génère un nom de fichier PDF basé sur l'expéditeur et la date.
    
    Args:
        sender: Expéditeur de l'email
        date_obj: Date de l'email
        subject: Sujet de l'email (optionnel)
    
    Returns:
        Nom de fichier PDF formaté
    """
    company = extract_sender_name(sender)
    date_str = format_date_for_filename(date_obj)
    
    if subject:
        subject_clean = sanitize_text(subject)[:50]  # Limiter la longueur
        filename = f"{company}_{date_str}_{subject_clean}.pdf"
    else:
        filename = f"{company}_{date_str}.pdf"
    
    return sanitize_filename(filename)


def validate_keywords(keywords_str: str) -> list:
    """
    Valide et parse une chaîne de mots clés séparés par des virgules.
    
    Args:
        keywords_str: Chaîne de mots clés séparés par des virgules
    
    Returns:
        Liste de mots clés nettoyés
    """
    if not keywords_str:
        return []
    
    keywords = [kw.strip() for kw in keywords_str.split(',')]
    keywords = [kw for kw in keywords if kw]  # Supprimer les vides
    
    return keywords


def validate_path(path: str) -> bool:
    """
    Vérifie si un chemin est valide.
    
    Args:
        path: Chemin à vérifier
    
    Returns:
        True si le chemin est valide
    """
    import os
    
    if not path:
        return False
    
    # Vérifier si c'est un chemin absolu valide
    try:
        return os.path.isabs(path) or os.path.exists(path)
    except (OSError, ValueError):
        return False
"""
Module de gestion des interactions avec Microsoft Outlook.
Utilise l'API COM Windows via pywin32.
"""

import os
import tempfile
from typing import List, Optional, Callable, Any
from datetime import datetime

try:
    import win32com.client
    from pywintypes import com_error
    OUTLOOK_AVAILABLE = True
except ImportError:
    OUTLOOK_AVAILABLE = False
    com_error = Exception

from utils.logger import logger


class OutlookError(Exception):
    """Exception personnalisée pour les erreurs Outlook"""
    pass


class EmailItem:
    """Représente un email Outlook avec ses métadonnées"""
    
    def __init__(self, mail_item):
        self._mail = mail_item
        self._attachments_saved = []
    
    @property
    def subject(self) -> str:
        """Sujet de l'email"""
        try:
            return self._mail.Subject or ""
        except com_error:
            return ""
    
    @property
    def sender(self) -> str:
        """Expéditeur de l'email"""
        try:
            return self._mail.SenderEmailAddress or self._mail.SenderName or ""
        except com_error:
            return ""
    
    @property
    def sender_name(self) -> str:
        """Nom de l'expéditeur"""
        try:
            return self._mail.SenderName or ""
        except com_error:
            return ""
    
    @property
    def received_time(self) -> Optional[datetime]:
        """Date/heure de réception"""
        try:
            return self._mail.ReceivedTime
        except com_error:
            return None
    
    @property
    def body(self) -> str:
        """Corps de l'email (texte brut)"""
        try:
            return self._mail.Body or ""
        except com_error:
            return ""
    
    @property
    def html_body(self) -> str:
        """Corps de l'email (HTML)"""
        try:
            return self._mail.HTMLBody or ""
        except com_error:
            return ""
    
    @property
    def has_attachments(self) -> bool:
        """Vérifie si l'email a des pièces jointes"""
        try:
            return self._mail.Attachments.Count > 0
        except com_error:
            return False
    
    @property
    def attachment_count(self) -> int:
        """Nombre de pièces jointes"""
        try:
            return self._mail.Attachments.Count
        except com_error:
            return 0
    
    @property
    def is_unread(self) -> bool:
        """Vérifie si l'email est non lu"""
        try:
            return self._mail.UnRead
        except com_error:
            return False
    
    def get_attachments_info(self) -> List[dict]:
        """Récupère les informations sur les pièces jointes"""
        attachments_info = []
        try:
            for i in range(1, self._mail.Attachments.Count + 1):
                att = self._mail.Attachments.Item(i)
                attachments_info.append({
                    'index': i,
                    'filename': att.FileName,
                    'size': att.Size,
                    'type': att.Type
                })
        except com_error as e:
            logger.error(f"Erreur lecture pièces jointes: {e}")
        return attachments_info
    
    def save_attachments(self, folder: str) -> List[str]:
        """
        Sauvegarde les pièces jointes dans un dossier.
        
        Args:
            folder: Chemin du dossier de destination
        
        Returns:
            Liste des chemins des fichiers sauvegardés
        """
        saved_files = []
        os.makedirs(folder, exist_ok=True)
        
        try:
            for i in range(1, self._mail.Attachments.Count + 1):
                att = self._mail.Attachments.Item(i)
                filename = att.FileName
                
                # Éviter les doublons
                filepath = os.path.join(folder, filename)
                counter = 1
                base, ext = os.path.splitext(filename)
                while os.path.exists(filepath):
                    filepath = os.path.join(folder, f"{base}_{counter}{ext}")
                    counter += 1
                
                att.SaveAsFile(filepath)
                saved_files.append(filepath)
                logger.debug(f"Pièce jointe sauvegardée: {filepath}")
                
        except com_error as e:
            logger.error(f"Erreur sauvegarde pièces jointes: {e}")
        
        self._attachments_saved = saved_files
        return saved_files
    
    def set_category(self, category: str):
        """Définit la catégorie de l'email"""
        try:
            self._mail.Categories = category
            self._mail.Save()
            logger.debug(f"Catégorie '{category}' appliquée à: {self.subject[:50]}")
        except com_error as e:
            logger.error(f"Erreur définition catégorie: {e}")
            raise OutlookError(f"Impossible de définir la catégorie: {e}")
    
    def mark_as_read(self):
        """Marque l'email comme lu"""
        try:
            self._mail.UnRead = False
            self._mail.Save()
        except com_error as e:
            logger.error(f"Erreur marquage comme lu: {e}")
    
    def move_to(self, target_folder) -> bool:
        """
        Déplace l'email vers un dossier.
        
        Args:
            target_folder: Objet dossier Outlook destination
        
        Returns:
            True si succès, False sinon
        """
        try:
            self._mail.Move(target_folder)
            logger.debug(f"Email déplacé: {self.subject[:50]}")
            return True
        except com_error as e:
            logger.error(f"Erreur déplacement email: {e}")
            return False
    
    def to_dict(self) -> dict:
        """Convertit l'email en dictionnaire"""
        return {
            'subject': self.subject,
            'sender': self.sender,
            'sender_name': self.sender_name,
            'received_time': self.received_time,
            'has_attachments': self.has_attachments,
            'attachment_count': self.attachment_count,
            'is_unread': self.is_unread
        }


class OutlookHandler:
    """Gestionnaire des interactions avec Outlook"""
    
    def __init__(self):
        if not OUTLOOK_AVAILABLE:
            raise OutlookError("pywin32 n'est pas installé. Installez-le avec: pip install pywin32")
        
        self._outlook = None
        self._namespace = None
        self._connected = False
    
    def connect(self) -> bool:
        """
        Établit la connexion avec Outlook.
        
        Returns:
            True si la connexion est établie
        """
        try:
            self._outlook = win32com.client.Dispatch("Outlook.Application")
            self._namespace = self._outlook.GetNamespace("MAPI")
            self._connected = True
            logger.success("Connexion à Outlook établie")
            return True
        except com_error as e:
            logger.error(f"Impossible de se connecter à Outlook: {e}")
            self._connected = False
            raise OutlookError(f"Connexion Outlook échouée: {e}")
    
    @property
    def is_connected(self) -> bool:
        """Vérifie si la connexion est active"""
        return self._connected and self._outlook is not None
    
    def get_mailboxes(self) -> List[str]:
        """
        Récupère la liste des boîtes aux lettres disponibles.
        
        Returns:
            Liste des noms de boîtes aux lettres
        """
        if not self.is_connected:
            self.connect()
        
        mailboxes = []
        try:
            for folder in self._namespace.Folders:
                mailboxes.append(folder.Name)
        except com_error as e:
            logger.error(f"Erreur récupération boîtes aux lettres: {e}")
        
        return mailboxes
    
    def get_mailbox(self, mailbox_name: str):
        """
        Récupère une boîte aux lettres par son nom.
        
        Args:
            mailbox_name: Nom de la boîte aux lettres
        
        Returns:
            Objet boîte aux lettres Outlook
        """
        if not self.is_connected:
            self.connect()
        
        try:
            return self._namespace.Folders[mailbox_name]
        except com_error as e:
            logger.error(f"Boîte aux lettres introuvable: {mailbox_name}")
            raise OutlookError(f"Boîte aux lettres '{mailbox_name}' introuvable")
    
    def get_inbox(self, mailbox_name: str):
        """
        Récupère la boîte de réception d'une boîte aux lettres.
        
        Args:
            mailbox_name: Nom de la boîte aux lettres
        
        Returns:
            Dossier Inbox
        """
        mailbox = self.get_mailbox(mailbox_name)
        try:
            # Constante olFolderInbox = 6
            return mailbox.Folders["Boîte de réception"]
        except com_error:
            try:
                return mailbox.Folders["Inbox"]
            except com_error:
                # Parcourir pour trouver la boîte de réception
                for folder in mailbox.Folders:
                    if "inbox" in folder.Name.lower() or "réception" in folder.Name.lower():
                        return folder
                raise OutlookError("Boîte de réception introuvable")
    
    def get_folder_by_path(self, folder_path: str):
        """
        Récupère un dossier par son chemin complet.
        
        Args:
            folder_path: Chemin du dossier (ex: "\\\\Mailbox\\Inbox\\Subfolder")
        
        Returns:
            Objet dossier Outlook
        """
        if not self.is_connected:
            self.connect()
        
        try:
            # Nettoyer le chemin
            path_parts = [p for p in folder_path.split('\\') if p]
            
            if not path_parts:
                raise OutlookError("Chemin de dossier vide")
            
            # Premier élément = boîte aux lettres
            current = self._namespace.Folders[path_parts[0]]
            
            # Parcourir les sous-dossiers
            for part in path_parts[1:]:
                current = current.Folders[part]
            
            return current
            
        except com_error as e:
            logger.error(f"Dossier introuvable: {folder_path}")
            raise OutlookError(f"Dossier '{folder_path}' introuvable")
    
    def pick_folder(self):
        """
        Ouvre le sélecteur de dossier Outlook natif.
        
        Returns:
            Objet dossier sélectionné ou None
        """
        if not self.is_connected:
            self.connect()
        
        try:
            return self._namespace.PickFolder()
        except com_error:
            return None
    
    def filter_emails(self, mailbox_name: str, keywords: List[str], 
                      unread_only: bool = False,
                      folder_name: str = None) -> List[EmailItem]:
        """
        Filtre les emails par mots clés dans l'objet.
        
        Args:
            mailbox_name: Nom de la boîte aux lettres
            keywords: Liste de mots clés à rechercher
            unread_only: Filtrer uniquement les non lus
            folder_name: Nom du dossier spécifique (sinon Inbox)
        
        Returns:
            Liste d'objets EmailItem correspondants
        """
        if not self.is_connected:
            self.connect()
        
        filtered_emails = []
        
        try:
            if folder_name:
                folder = self.get_folder_by_path(folder_name)
            else:
                folder = self.get_inbox(mailbox_name)
            
            items = folder.Items
            items.Sort("[ReceivedTime]", True)  # Plus récents d'abord
            
            for item in items:
                try:
                    # Vérifier si c'est un email
                    if item.Class != 43:  # olMail = 43
                        continue
                    
                    # Filtre non lu si demandé
                    if unread_only and not item.UnRead:
                        continue
                    
                    # Vérifier les mots clés dans l'objet
                    subject = item.Subject or ""
                    subject_lower = subject.lower()
                    
                    if any(kw.lower() in subject_lower for kw in keywords):
                        filtered_emails.append(EmailItem(item))
                        logger.debug(f"Email correspondant trouvé: {subject[:50]}")
                        
                except com_error:
                    continue
            
            logger.info(f"{len(filtered_emails)} email(s) correspondant(s) trouvé(s)")
            
        except com_error as e:
            logger.error(f"Erreur filtrage emails: {e}")
        
        return filtered_emails
    
    def get_all_emails(self, mailbox_name: str, limit: int = 100) -> List[EmailItem]:
        """
        Récupère tous les emails (avec limite).
        
        Args:
            mailbox_name: Nom de la boîte aux lettres
            limit: Nombre maximum d'emails à récupérer
        
        Returns:
            Liste d'objets EmailItem
        """
        if not self.is_connected:
            self.connect()
        
        emails = []
        
        try:
            inbox = self.get_inbox(mailbox_name)
            items = inbox.Items
            items.Sort("[ReceivedTime]", True)
            
            count = 0
            for item in items:
                if count >= limit:
                    break
                try:
                    if item.Class == 43:  # olMail
                        emails.append(EmailItem(item))
                        count += 1
                except com_error:
                    continue
                    
        except com_error as e:
            logger.error(f"Erreur récupération emails: {e}")
        
        return emails
    
    def move_email(self, email: EmailItem, target_folder) -> bool:
        """
        Déplace un email vers un dossier.
        
        Args:
            email: EmailItem à déplacer
            target_folder: Dossier destination
        
        Returns:
            True si succès
        """
        return email.move_to(target_folder)
    
    def categorize_email(self, email: EmailItem, category: str):
        """
        Applique une catégorie à un email.
        
        Args:
            email: EmailItem à catégoriser
            category: Nom de la catégorie
        """
        email.set_category(category)
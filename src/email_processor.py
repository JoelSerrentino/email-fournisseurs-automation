"""
Module principal de traitement des emails pour l'application Email Fournisseurs Automation.
Coordonne le filtrage, le déplacement, la génération PDF et la catégorisation.
"""

import os
import tempfile
import shutil
from typing import List, Optional, Callable, Dict, Any
from dataclasses import dataclass
from datetime import datetime
from enum import Enum

from outlook_handler import OutlookHandler, EmailItem, OutlookError
from pdf_generator import PDFGenerator, PDFGeneratorError
from utils.logger import logger
from utils.sanitize import validate_keywords


class ProcessingStatus(Enum):
    """États possibles du traitement"""
    PENDING = "pending"
    IN_PROGRESS = "in_progress"
    SUCCESS = "success"
    FAILED = "failed"
    SKIPPED = "skipped"


@dataclass
class ProcessingResult:
    """Résultat du traitement d'un email"""
    email_subject: str
    email_sender: str
    status: ProcessingStatus
    pdf_path: Optional[str] = None
    error_message: Optional[str] = None
    
    def to_dict(self) -> Dict[str, Any]:
        return {
            'subject': self.email_subject,
            'sender': self.email_sender,
            'status': self.status.value,
            'pdf_path': self.pdf_path,
            'error': self.error_message
        }


@dataclass
class ProcessingStats:
    """Statistiques de traitement"""
    total: int = 0
    processed: int = 0
    success: int = 0
    failed: int = 0
    skipped: int = 0
    
    @property
    def progress_percent(self) -> float:
        if self.total == 0:
            return 0.0
        return (self.processed / self.total) * 100
    
    def to_dict(self) -> Dict[str, Any]:
        return {
            'total': self.total,
            'processed': self.processed,
            'success': self.success,
            'failed': self.failed,
            'skipped': self.skipped,
            'progress': self.progress_percent
        }


class EmailProcessor:
    """
    Processeur principal pour le traitement des emails fournisseurs.
    
    Gère le workflow complet:
    1. Filtrage des emails par mots clés
    2. Extraction des pièces jointes
    3. Génération du PDF (email + pièces jointes)
    4. Déplacement vers le dossier cible
    5. Application de la catégorie Outlook
    """
    
    def __init__(self, output_folder: str, 
                 progress_callback: Optional[Callable[[int, int, str], None]] = None,
                 log_callback: Optional[Callable[[str, str], None]] = None):
        """
        Initialise le processeur d'emails.
        
        Args:
            output_folder: Dossier de sortie pour les PDF
            progress_callback: Callback de progression (current, total, message)
            log_callback: Callback de log (message, level)
        """
        self.output_folder = output_folder
        self.outlook_handler = OutlookHandler()
        self.pdf_generator = PDFGenerator(output_folder)
        
        # Callbacks pour l'interface
        self._progress_callback: Optional[Callable[[int, int, str], None]] = progress_callback
        self._status_callback: Optional[Callable[[str], None]] = None
        self._log_callback: Optional[Callable[[str, str], None]] = log_callback
        
        # Statistiques
        self.stats = ProcessingStats()
        self.results: List[ProcessingResult] = []
        
        # État
        self._is_running = False
        self._should_stop = False
        
        # Dossier temporaire pour les pièces jointes
        self._temp_dir = None
    
    def set_progress_callback(self, callback: Callable[[int, int, str], None]):
        """
        Définit le callback de progression.
        
        Args:
            callback: Fonction(current, total, message) appelée à chaque étape
        """
        self._progress_callback = callback
    
    def set_status_callback(self, callback: Callable[[str], None]):
        """
        Définit le callback de statut.
        
        Args:
            callback: Fonction(status) appelée pour les changements d'état
        """
        self._status_callback = callback
    
    def _report_progress(self, current: int, total: int, message: str):
        """Rapporte la progression"""
        if self._progress_callback:
            self._progress_callback(current, total, message)
    
    def _report_status(self, status: str):
        """Rapporte le statut"""
        if self._status_callback:
            self._status_callback(status)
    
    def _report_log(self, message: str, level: str = "info"):
        """Rapporte un message de log"""
        if self._log_callback:
            self._log_callback(message, level)
    
    def stop(self):
        """Demande l'arrêt du traitement"""
        self._should_stop = True
        logger.warning("Arrêt du traitement demandé...")
        if self._log_callback:
            self._log_callback("Arrêt du traitement demandé...", "warning")
    
    def _check_stop(self) -> bool:
        """Vérifie si l'arrêt a été demandé"""
        return self._should_stop
    
    @property
    def is_running(self) -> bool:
        """Vérifie si le traitement est en cours"""
        return self._is_running
    
    def connect_outlook(self) -> bool:
        """
        Établit la connexion avec Outlook.
        
        Returns:
            True si la connexion est établie
        """
        try:
            self.outlook_handler.connect()
            return True
        except OutlookError as e:
            logger.error(f"Connexion Outlook échouée: {e}")
            return False
    
    def get_matching_emails(self, mailbox_name: str, keywords: List[str],
                           unread_only: bool = False,
                           date_from: str = None, date_to: str = None) -> List[EmailItem]:
        """
        Récupère les emails correspondant aux critères.
        
        Args:
            mailbox_name: Nom de la boîte aux lettres
            keywords: Liste des mots clés à rechercher
            unread_only: Filtrer uniquement les non lus
            date_from: Date de début (format JJ/MM/AAAA)
            date_to: Date de fin (format JJ/MM/AAAA)
        
        Returns:
            Liste des emails correspondants
        """
        return self.outlook_handler.filter_emails(
            mailbox_name, keywords, unread_only=unread_only,
            date_from=date_from, date_to=date_to
        )
    
    def preview_emails(self, mailbox_name: str, keywords_str: str,
                       unread_only: bool = False) -> List[Dict[str, Any]]:
        """
        Prévisualise les emails qui seraient traités.
        
        Args:
            mailbox_name: Nom de la boîte aux lettres
            keywords_str: Mots clés séparés par des virgules
            unread_only: Filtrer uniquement les non lus
        
        Returns:
            Liste de dictionnaires avec les infos des emails
        """
        keywords = validate_keywords(keywords_str)
        if not keywords:
            return []
        
        if not self.outlook_handler.is_connected:
            self.connect_outlook()
        
        emails = self.get_matching_emails(mailbox_name, keywords, unread_only)
        
        return [email.to_dict() for email in emails]
    
    def process_emails(self, mailbox_name: str, keywords_str: str,
                       target_folder_path: str, category: str,
                       unread_only: bool = False,
                       date_from: str = None, date_to: str = None) -> ProcessingStats:
        """
        Traite tous les emails correspondant aux critères.
        
        Args:
            mailbox_name: Nom de la boîte aux lettres
            keywords_str: Mots clés séparés par des virgules
            target_folder_path: Chemin du dossier Outlook destination
            category: Catégorie à appliquer après traitement (succès)
            unread_only: Filtrer uniquement les non lus
            date_from: Date de début (format JJ/MM/AAAA)
            date_to: Date de fin (format JJ/MM/AAAA)
        
        Returns:
            Statistiques de traitement
        """
        # Réinitialiser l'état
        self._is_running = True
        self._should_stop = False
        self.stats = ProcessingStats()
        self.results = []
        
        # Catégorie pour les erreurs
        error_category = "Erreur traitement"
        
        # Créer le dossier temporaire
        self._temp_dir = tempfile.mkdtemp(prefix='email_fournisseurs_')
        
        try:
            self._report_status("Connexion à Outlook...")
            
            # Connexion
            if not self.outlook_handler.is_connected:
                if not self.connect_outlook():
                    raise OutlookError("Impossible de se connecter à Outlook")
            
            # Créer les catégories avec les bonnes couleurs
            if category:
                self.outlook_handler.ensure_category_exists(category, 'green')
            self.outlook_handler.ensure_category_exists(error_category, 'red')
            
            # Valider les mots clés
            keywords = validate_keywords(keywords_str)
            if not keywords:
                logger.warning("Aucun mot clé défini")
                return self.stats
            
            logger.info(f"Mots clés: {', '.join(keywords)}")
            
            # Récupérer le dossier cible
            target_folder = None
            if target_folder_path:
                try:
                    target_folder = self.outlook_handler.get_folder_by_path(target_folder_path)
                    logger.info(f"Dossier cible: {target_folder_path}")
                except OutlookError as e:
                    logger.warning(f"Dossier cible introuvable, les emails ne seront pas déplacés: {e}")
            
            # Rechercher les emails
            self._report_status("Recherche des emails...")
            emails = self.get_matching_emails(mailbox_name, keywords, unread_only, date_from, date_to)
            
            self.stats.total = len(emails)
            
            if self.stats.total == 0:
                logger.info("Aucun email correspondant trouvé")
                self._report_status("Aucun email trouvé")
                return self.stats
            
            logger.info(f"{self.stats.total} email(s) à traiter")
            
            # Traiter chaque email
            for i, email in enumerate(emails):
                if self._should_stop:
                    logger.warning("Traitement interrompu par l'utilisateur")
                    break
                
                self._report_progress(i + 1, self.stats.total, 
                                     f"Traitement: {email.subject[:40]}...")
                
                result = self._process_single_email(email, target_folder, category, error_category)
                self.results.append(result)
                
                # Mettre à jour les statistiques
                self.stats.processed += 1
                if result.status == ProcessingStatus.SUCCESS:
                    self.stats.success += 1
                elif result.status == ProcessingStatus.FAILED:
                    self.stats.failed += 1
                else:
                    self.stats.skipped += 1
            
            # Rapport final
            logger.success(f"Traitement terminé: {self.stats.success}/{self.stats.total} succès, "
                          f"{self.stats.failed} échec(s)")
            
            self._report_status("Traitement terminé")
            
        except Exception as e:
            logger.error(f"Erreur lors du traitement: {e}")
            self._report_status(f"Erreur: {e}")
            
        finally:
            # Nettoyer le dossier temporaire
            if self._temp_dir and os.path.exists(self._temp_dir):
                try:
                    shutil.rmtree(self._temp_dir)
                except OSError:
                    pass
            
            self._is_running = False
        
        return self.stats
    
    def _process_single_email(self, email: EmailItem, target_folder,
                               category: str, error_category: str = "Erreur traitement") -> ProcessingResult:
        """
        Traite un seul email.
        
        Args:
            email: EmailItem à traiter
            target_folder: Dossier Outlook destination (ou None)
            category: Catégorie à appliquer en cas de succès (vert)
            error_category: Catégorie à appliquer en cas d'erreur (rouge)
        
        Returns:
            Résultat du traitement
        """
        result = ProcessingResult(
            email_subject=email.subject,
            email_sender=email.sender,
            status=ProcessingStatus.PENDING
        )
        
        # Vérifier si l'arrêt est demandé avant de commencer
        if self._check_stop():
            result.status = ProcessingStatus.SKIPPED
            result.error_message = "Traitement interrompu par l'utilisateur"
            return result
        
        try:
            logger.info(f"Traitement: {email.subject[:50]}")
            result.status = ProcessingStatus.IN_PROGRESS
            
            # 1. Sauvegarder les pièces jointes
            attachment_paths = []
            if email.has_attachments:
                if self._check_stop():
                    result.status = ProcessingStatus.SKIPPED
                    return result
                email_temp_dir = os.path.join(self._temp_dir, 
                                              f"email_{datetime.now().strftime('%Y%m%d%H%M%S%f')}")
                attachment_paths = email.save_attachments(email_temp_dir)
                logger.debug(f"{len(attachment_paths)} pièce(s) jointe(s) sauvegardée(s)")
            
            # 2. Générer le PDF (vérifier l'arrêt avant)
            if self._check_stop():
                result.status = ProcessingStatus.SKIPPED
                return result
            
            # Récupérer received_time avec protection contre les erreurs win32timezone
            try:
                received_time = email.received_time
            except Exception:
                received_time = None
                
            pdf_path = self.pdf_generator.generate_email_pdf(
                sender=email.sender,
                sender_name=email.sender_name,
                subject=email.subject,
                body=email.body,
                received_time=received_time,
                attachment_paths=attachment_paths
            )
            result.pdf_path = pdf_path
            
            # 3. Appliquer la catégorie SUCCÈS (vert) AVANT de déplacer
            if category:
                email.set_category(category)
            
            # 4. Marquer comme lu AVANT de déplacer
            email.mark_as_read()
            
            # 5. Déplacer l'email si un dossier cible est défini (en dernier!)
            if self._check_stop():
                result.status = ProcessingStatus.SKIPPED
                return result
                
            if target_folder:
                email.move_to(target_folder)
            
            result.status = ProcessingStatus.SUCCESS
            logger.success(f"Email traité avec succès: {email.subject[:40]}")
            
        except PDFGeneratorError as e:
            result.status = ProcessingStatus.FAILED
            result.error_message = f"Erreur génération PDF: {e}"
            logger.error(result.error_message)
            # Appliquer la catégorie d'erreur (rouge)
            try:
                email.set_category(error_category)
            except:
                pass
            
        except OutlookError as e:
            result.status = ProcessingStatus.FAILED
            result.error_message = f"Erreur Outlook: {e}"
            logger.error(result.error_message)
            # Appliquer la catégorie d'erreur (rouge)
            try:
                email.set_category(error_category)
            except:
                pass
            
        except Exception as e:
            result.status = ProcessingStatus.FAILED
            result.error_message = f"Erreur inattendue: {e}"
            logger.error(result.error_message)
            # Appliquer la catégorie d'erreur (rouge)
            try:
                email.set_category(error_category)
            except Exception:
                # Ignorer les erreurs lors de la définition de la catégorie
                pass
        
        return result
    
    def get_results_summary(self) -> Dict[str, Any]:
        """
        Récupère un résumé des résultats de traitement.
        
        Returns:
            Dictionnaire avec les statistiques et résultats
        """
        return {
            'stats': self.stats.to_dict(),
            'results': [r.to_dict() for r in self.results]
        }


# Fonction de compatibilité avec l'ancien code
def process_emails(outlook_handler, pdf_generator, output_folder,
                   keywords, target_folder):
    """Fonction de compatibilité avec l'ancien code"""
    processor = EmailProcessor(output_folder)
    processor.outlook_handler = outlook_handler
    processor.pdf_generator = pdf_generator
    
    # Cette fonction est dépréciée, utiliser EmailProcessor.process_emails() directement
    logger.warning("process_emails() est déprécié, utilisez EmailProcessor.process_emails()")
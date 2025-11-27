"""
Module de journalisation pour l'application Email Fournisseurs Automation.
Gère les logs dans les fichiers et les callbacks pour l'interface graphique.
"""

import os
import logging
from datetime import datetime
from typing import Callable, Optional
from enum import Enum


class LogLevel(Enum):
    """Niveaux de log"""
    DEBUG = "debug"
    INFO = "info"
    SUCCESS = "success"
    WARNING = "warning"
    ERROR = "error"


class Logger:
    """Gestionnaire de logs centralisé"""
    
    _instance = None
    _gui_callback: Optional[Callable[[str, str], None]] = None
    
    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
            cls._instance._initialized = False
        return cls._instance
    
    def __init__(self, log_file: str = 'logs/email_processing.log'):
        if self._initialized:
            return
            
        self.log_file = log_file
        self._ensure_log_directory()
        self._setup_file_logger()
        self._initialized = True
    
    def _ensure_log_directory(self):
        """Crée le dossier de logs si nécessaire"""
        log_dir = os.path.dirname(self.log_file)
        if log_dir:
            os.makedirs(log_dir, exist_ok=True)
    
    def _setup_file_logger(self):
        """Configure le logger fichier"""
        self.file_logger = logging.getLogger('email_fournisseurs')
        self.file_logger.setLevel(logging.DEBUG)
        
        # Handler pour fichier
        file_handler = logging.FileHandler(self.log_file, encoding='utf-8')
        file_handler.setLevel(logging.DEBUG)
        
        # Format
        formatter = logging.Formatter(
            '%(asctime)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        file_handler.setFormatter(formatter)
        
        # Éviter les doublons
        if not self.file_logger.handlers:
            self.file_logger.addHandler(file_handler)
    
    @classmethod
    def set_gui_callback(cls, callback: Callable[[str, str], None]):
        """Définit le callback pour l'interface graphique"""
        cls._gui_callback = callback
    
    def log(self, message: str, level: LogLevel = LogLevel.INFO):
        """Log un message avec le niveau spécifié"""
        # Log dans le fichier
        if level == LogLevel.DEBUG:
            self.file_logger.debug(message)
        elif level == LogLevel.INFO or level == LogLevel.SUCCESS:
            self.file_logger.info(message)
        elif level == LogLevel.WARNING:
            self.file_logger.warning(message)
        elif level == LogLevel.ERROR:
            self.file_logger.error(message)
        
        # Callback GUI si disponible
        if self._gui_callback:
            self._gui_callback(message, level.value)
    
    def debug(self, message: str):
        self.log(message, LogLevel.DEBUG)
    
    def info(self, message: str):
        self.log(message, LogLevel.INFO)
    
    def success(self, message: str):
        self.log(message, LogLevel.SUCCESS)
    
    def warning(self, message: str):
        self.log(message, LogLevel.WARNING)
    
    def error(self, message: str):
        self.log(message, LogLevel.ERROR)
    
    def get_log_content(self, lines: int = 100) -> str:
        """Récupère les dernières lignes du fichier de log"""
        if not os.path.exists(self.log_file):
            return ""
        
        with open(self.log_file, 'r', encoding='utf-8') as f:
            all_lines = f.readlines()
            return ''.join(all_lines[-lines:])
    
    def clear_log_file(self):
        """Vide le fichier de log"""
        with open(self.log_file, 'w', encoding='utf-8') as f:
            f.write('')
        self.info("Fichier de log effacé")
    
    def export_logs(self, export_path: str) -> bool:
        """Exporte les logs vers un fichier"""
        try:
            import shutil
            shutil.copy2(self.log_file, export_path)
            self.success(f"Logs exportés vers: {export_path}")
            return True
        except Exception as e:
            self.error(f"Erreur lors de l'export des logs: {e}")
            return False


# Instance globale
logger = Logger()


# Fonction de compatibilité avec l'ancien code
def log_message(message: str, log_file: str = 'logs/email_processing.log'):
    """Fonction de compatibilité pour l'ancien code"""
    logger.info(message)
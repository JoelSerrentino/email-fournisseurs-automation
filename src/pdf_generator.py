"""
Module de génération de fichiers PDF pour l'application Email Fournisseurs Automation.
Fusionne les emails et leurs pièces jointes en un seul fichier PDF.
"""

import os
import tempfile
from typing import List, Optional
from datetime import datetime

from utils.logger import logger
from utils.sanitize import generate_pdf_filename, sanitize_filename


class PDFGeneratorError(Exception):
    """Exception personnalisée pour les erreurs de génération PDF"""
    pass


class PDFGenerator:
    """Générateur de fichiers PDF à partir d'emails et pièces jointes"""
    
    # Extensions supportées pour les pièces jointes
    SUPPORTED_IMAGE_EXTENSIONS = ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.webp']
    SUPPORTED_PDF_EXTENSION = '.pdf'
    SUPPORTED_TEXT_EXTENSIONS = ['.txt', '.csv', '.log']
    SUPPORTED_WORD_EXTENSIONS = ['.doc', '.docx']
    SUPPORTED_EXCEL_EXTENSIONS = ['.xls', '.xlsx']
    
    def __init__(self, output_dir: str):
        """
        Initialise le générateur PDF.
        
        Args:
            output_dir: Dossier de sortie pour les PDF générés
        """
        self.output_dir = output_dir
        os.makedirs(output_dir, exist_ok=True)
        
        # Vérifier les dépendances
        self._check_dependencies()
    
    def _check_dependencies(self):
        """Vérifie que les dépendances sont disponibles"""
        try:
            from reportlab.lib.pagesizes import A4
            from reportlab.platypus import SimpleDocTemplate
            self._has_reportlab = True
        except ImportError:
            self._has_reportlab = False
            logger.warning("reportlab non installé - génération PDF limitée")
        
        try:
            import PyPDF2
            self._has_pypdf2 = True
        except ImportError:
            self._has_pypdf2 = False
            logger.warning("PyPDF2 non installé - fusion PDF limitée")
        
        try:
            from PIL import Image
            self._has_pillow = True
        except ImportError:
            self._has_pillow = False
            logger.warning("Pillow non installé - images non supportées")
        
        # Vérifier si on peut convertir les documents Word (via COM Windows)
        try:
            import win32com.client
            self._has_win32com = True
        except ImportError:
            self._has_win32com = False
            logger.warning("pywin32 non installé - conversion Word/Excel limitée")
    
    def generate_email_pdf(self, sender: str, sender_name: str, subject: str,
                           body: str, received_time: Optional[datetime],
                           attachment_paths: List[str] = None, recipient: str = None) -> str:
        """
        Génère un PDF à partir d'un email et ses pièces jointes.
        
        Args:
            sender: Adresse email de l'expéditeur
            sender_name: Nom de l'expéditeur
            subject: Sujet de l'email
            body: Corps de l'email
            received_time: Date/heure de réception
            attachment_paths: Liste des chemins des pièces jointes
            recipient: Adresse email du destinataire
        
        Returns:
            Chemin du fichier PDF généré
        """
        if not self._has_reportlab:
            raise PDFGeneratorError("reportlab est requis pour générer des PDF")
        
        # Logging de la date reçue
        logger.debug(f"received_time type: {type(received_time)}, value: {received_time}")
        
        # Générer le nom de fichier
        filename = generate_pdf_filename(sender_name or sender, received_time, subject)
        output_path = os.path.join(self.output_dir, filename)
        
        # Éviter les doublons
        base, ext = os.path.splitext(output_path)
        counter = 1
        while os.path.exists(output_path):
            output_path = f"{base}_{counter}{ext}"
            counter += 1
        
        try:
            # Créer le PDF de l'email
            email_pdf_path = self._create_email_pdf(
                sender, sender_name, subject, body, received_time, output_path, recipient
            )
            
            # Fusionner avec les pièces jointes si présentes
            if attachment_paths:
                merged_path = self._merge_with_attachments(email_pdf_path, attachment_paths)
                if merged_path != email_pdf_path:
                    # Supprimer le PDF intermédiaire
                    try:
                        os.remove(email_pdf_path)
                    except OSError:
                        pass
                    return merged_path
            
            logger.success(f"PDF généré: {os.path.basename(output_path)}")
            return email_pdf_path
            
        except Exception as e:
            logger.error(f"Erreur génération PDF: {e}")
            raise PDFGeneratorError(f"Impossible de générer le PDF: {e}")
    
    def _create_email_pdf(self, sender: str, sender_name: str, subject: str,
                          body: str, received_time: Optional[datetime],
                          output_path: str, recipient: str = None) -> str:
        """Crée un PDF à partir du contenu de l'email"""
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.units import cm
        from reportlab.lib.colors import HexColor
        from reportlab.platypus import (
            SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, 
            PageBreak, HRFlowable
        )
        from reportlab.lib.enums import TA_LEFT, TA_CENTER
        
        # Créer le document
        doc = SimpleDocTemplate(
            output_path,
            pagesize=A4,
            rightMargin=2*cm,
            leftMargin=2*cm,
            topMargin=2*cm,
            bottomMargin=2*cm
        )
        
        # Styles
        styles = getSampleStyleSheet()
        
        # Style personnalisé pour le titre
        title_style = ParagraphStyle(
            'CustomTitle',
            parent=styles['Heading1'],
            fontSize=14,
            spaceAfter=12,
            textColor=HexColor('#0078d4')
        )
        
        # Style pour les métadonnées
        meta_style = ParagraphStyle(
            'Meta',
            parent=styles['Normal'],
            fontSize=10,
            textColor=HexColor('#5c5c5c'),
            spaceAfter=4
        )
        
        # Style pour le corps
        body_style = ParagraphStyle(
            'Body',
            parent=styles['Normal'],
            fontSize=11,
            spaceAfter=6,
            leading=14
        )
        
        # Contenu
        elements = []
        
        # En-tête avec bandeau
        elements.append(Paragraph("📧 Email Fournisseur", title_style))
        elements.append(HRFlowable(width="100%", thickness=2, color=HexColor('#0078d4')))
        elements.append(Spacer(1, 12))
        
        # Métadonnées de l'email
        logger.debug(f"Formatage de received_time: {type(received_time)} = {received_time}")
        if received_time:
            try:
                date_str = received_time.strftime('%d/%m/%Y à %H:%M')
                logger.debug(f"Date formatée: {date_str}")
            except Exception as e:
                logger.debug(f"Erreur formatage date: {e}")
                date_str = "Date inconnue"
        else:
            logger.debug("received_time est None")
            date_str = "Date inconnue"
        
        meta_data = [
            ['De:', sender_name or sender],
            ['À :', recipient or "(Non disponible)"],
            ['Sujet:', subject or "(Sans objet)"],
            ['Date:', date_str]
        ]
        
        meta_table = Table(meta_data, colWidths=[3*cm, 13*cm])
        meta_table.setStyle(TableStyle([
            ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
            ('FONTNAME', (1, 0), (1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('TEXTCOLOR', (0, 0), (0, -1), HexColor('#5c5c5c')),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ]))
        elements.append(meta_table)
        elements.append(Spacer(1, 20))
        
        # Séparateur
        elements.append(HRFlowable(width="100%", thickness=1, color=HexColor('#e0e0e0')))
        elements.append(Spacer(1, 12))
        
        # Corps de l'email
        elements.append(Paragraph("<b>Contenu de l'email:</b>", meta_style))
        elements.append(Spacer(1, 8))
        
        # Nettoyer et formater le corps
        if body:
            # Échapper les caractères spéciaux pour ReportLab
            body_clean = body.replace('&', '&amp;')
            body_clean = body_clean.replace('<', '&lt;')
            body_clean = body_clean.replace('>', '&gt;')
            
            # Convertir les sauts de ligne
            paragraphs = body_clean.split('\n')
            for para in paragraphs:
                if para.strip():
                    elements.append(Paragraph(para, body_style))
                else:
                    elements.append(Spacer(1, 6))
        else:
            elements.append(Paragraph("<i>(Email sans contenu texte)</i>", meta_style))
        
        # Pied de page
        elements.append(Spacer(1, 30))
        elements.append(HRFlowable(width="100%", thickness=1, color=HexColor('#e0e0e0')))
        elements.append(Spacer(1, 8))
        
        footer_text = f"Généré le {datetime.now().strftime('%d/%m/%Y à %H:%M')} - Email Fournisseurs Automation"
        footer_style = ParagraphStyle(
            'Footer',
            parent=styles['Normal'],
            fontSize=8,
            textColor=HexColor('#a0a0a0'),
            alignment=TA_CENTER
        )
        elements.append(Paragraph(footer_text, footer_style))
        
        # Générer le PDF
        doc.build(elements)
        
        return output_path
    
    def _merge_with_attachments(self, email_pdf_path: str, 
                                 attachment_paths: List[str]) -> str:
        """
        Fusionne le PDF de l'email avec les pièces jointes.
        
        Args:
            email_pdf_path: Chemin du PDF de l'email
            attachment_paths: Liste des chemins des pièces jointes
        
        Returns:
            Chemin du PDF fusionné
        """
        if not self._has_pypdf2:
            logger.warning("PyPDF2 non disponible - pièces jointes non fusionnées")
            return email_pdf_path
        
        import PyPDF2
        
        # Créer le merger
        merger = PyPDF2.PdfMerger()

        # Traiter chaque pièce jointe et conserver les PDFs à ajouter
        temp_files = []
        processed_att_paths: List[str] = []

        for att_path in attachment_paths:
            if not os.path.exists(att_path):
                logger.warning(f"Pièce jointe introuvable: {att_path}")
                continue
            
            ext = os.path.splitext(att_path)[1].lower()
            
            try:
                if ext == self.SUPPORTED_PDF_EXTENSION:
                    # Marquer pour ajout (les pièces jointes seront ajoutées avant l'email)
                    processed_att_paths.append(att_path)
                    logger.debug(f"PDF ajouté à la liste: {os.path.basename(att_path)}")
                    
                elif ext in self.SUPPORTED_IMAGE_EXTENSIONS:
                    # Ne pas intégrer les images
                    logger.debug(f"Image ignorée (non intégrée): {os.path.basename(att_path)}")
                
                elif ext in self.SUPPORTED_WORD_EXTENSIONS and self._has_win32com:
                    # Convertir le document Word en PDF
                    temp_pdf = self._word_to_pdf(att_path)
                    if temp_pdf:
                        processed_att_paths.append(temp_pdf)
                        temp_files.append(temp_pdf)
                        logger.debug(f"Document Word converti: {os.path.basename(att_path)}")
                
                elif ext in self.SUPPORTED_EXCEL_EXTENSIONS and self._has_win32com:
                    # Convertir le document Excel en PDF
                    temp_pdf = self._excel_to_pdf(att_path)
                    if temp_pdf:
                        processed_att_paths.append(temp_pdf)
                        temp_files.append(temp_pdf)
                        logger.debug(f"Document Excel converti: {os.path.basename(att_path)}")
                        
                elif ext in self.SUPPORTED_TEXT_EXTENSIONS:
                    # Convertir le texte en PDF
                    temp_pdf = self._text_to_pdf(att_path)
                    if temp_pdf:
                        processed_att_paths.append(temp_pdf)
                        temp_files.append(temp_pdf)
                        logger.debug(f"Texte converti: {os.path.basename(att_path)}")
                else:
                    logger.warning(f"Type de fichier non supporté: {ext}")
                    
            except Exception as e:
                logger.error(f"Erreur traitement pièce jointe {att_path}: {e}")
        
        # Ajouter d'abord les pièces jointes traitées, puis le PDF de l'email
        for p in processed_att_paths:
            try:
                merger.append(p)
            except Exception as e:
                logger.error(f"Erreur ajout pièce jointe au merger: {p} => {e}")

        # Enfin ajouter le PDF de l'email
        merger.append(email_pdf_path)

        # Sauvegarder le PDF fusionné
        output_path = email_pdf_path.replace('.pdf', '_complet.pdf')

        try:
            with open(output_path, 'wb') as output_file:
                merger.write(output_file)
            merger.close()

            # Nettoyer les fichiers temporaires
            for temp_file in temp_files:
                try:
                    os.remove(temp_file)
                except OSError:
                    pass

            return output_path

        except Exception as e:
            merger.close()
            logger.error(f"Erreur fusion PDF: {e}")
            return email_pdf_path
    
    def _image_to_pdf(self, image_path: str) -> Optional[str]:
        """Convertit une image en PDF"""
        if not self._has_pillow:
            return None
        
        from PIL import Image
        from reportlab.lib.pagesizes import A4
        from reportlab.pdfgen import canvas
        
        try:
            # Créer un fichier temporaire
            temp_fd, temp_path = tempfile.mkstemp(suffix='.pdf')
            os.close(temp_fd)
            
            # Ouvrir l'image
            img = Image.open(image_path)
            
            # Convertir en RGB si nécessaire
            if img.mode in ('RGBA', 'LA', 'P'):
                background = Image.new('RGB', img.size, (255, 255, 255))
                if img.mode == 'P':
                    img = img.convert('RGBA')
                background.paste(img, mask=img.split()[-1] if img.mode == 'RGBA' else None)
                img = background
            elif img.mode != 'RGB':
                img = img.convert('RGB')
            
            # Calculer les dimensions pour tenir sur une page A4
            page_width, page_height = A4
            margin = 50
            max_width = page_width - 2 * margin
            max_height = page_height - 2 * margin
            
            img_width, img_height = img.size
            ratio = min(max_width / img_width, max_height / img_height)
            
            new_width = int(img_width * ratio)
            new_height = int(img_height * ratio)
            
            # Créer le PDF
            c = canvas.Canvas(temp_path, pagesize=A4)
            
            # Centrer l'image
            x = (page_width - new_width) / 2
            y = (page_height - new_height) / 2
            
            # Sauvegarder temporairement l'image
            temp_img_fd, temp_img_path = tempfile.mkstemp(suffix='.jpg')
            os.close(temp_img_fd)
            img.save(temp_img_path, 'JPEG', quality=95)
            
            c.drawImage(temp_img_path, x, y, width=new_width, height=new_height)
            c.save()
            
            # Nettoyer
            os.remove(temp_img_path)
            
            return temp_path
            
        except Exception as e:
            logger.error(f"Erreur conversion image: {e}")
            return None
    
    def _text_to_pdf(self, text_path: str) -> Optional[str]:
        """Convertit un fichier texte en PDF"""
        if not self._has_reportlab:
            return None
        
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.styles import getSampleStyleSheet
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
        from reportlab.lib.units import cm
        
        try:
            # Lire le contenu
            with open(text_path, 'r', encoding='utf-8', errors='replace') as f:
                content = f.read()
            
            # Créer un fichier temporaire
            temp_fd, temp_path = tempfile.mkstemp(suffix='.pdf')
            os.close(temp_fd)
            
            # Créer le PDF
            doc = SimpleDocTemplate(temp_path, pagesize=A4,
                                   rightMargin=2*cm, leftMargin=2*cm,
                                   topMargin=2*cm, bottomMargin=2*cm)
            
            styles = getSampleStyleSheet()
            elements = []
            
            # Titre du fichier
            elements.append(Paragraph(f"<b>Pièce jointe: {os.path.basename(text_path)}</b>",
                                     styles['Heading2']))
            elements.append(Spacer(1, 12))
            
            # Contenu
            for line in content.split('\n'):
                if line.strip():
                    # Échapper les caractères spéciaux
                    line = line.replace('&', '&amp;')
                    line = line.replace('<', '&lt;')
                    line = line.replace('>', '&gt;')
                    elements.append(Paragraph(line, styles['Code']))
                else:
                    elements.append(Spacer(1, 6))
            
            doc.build(elements)
            return temp_path
            
        except Exception as e:
            logger.error(f"Erreur conversion texte: {e}")
            return None
    
    def _word_to_pdf(self, word_path: str) -> Optional[str]:
        """Convertit un document Word (.doc, .docx) en PDF via Microsoft Word"""
        if not self._has_win32com:
            logger.warning("pywin32 requis pour convertir les documents Word")
            return None
        
        import win32com.client
        from pywintypes import com_error
        
        word = None
        doc = None
        
        try:
            # Créer un fichier temporaire pour le PDF
            temp_fd, temp_path = tempfile.mkstemp(suffix='.pdf')
            os.close(temp_fd)
            
            # Convertir le chemin en absolu
            word_path_abs = os.path.abspath(word_path)
            temp_path_abs = os.path.abspath(temp_path)
            
            # Ouvrir Word
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            word.DisplayAlerts = False
            
            # Ouvrir le document
            doc = word.Documents.Open(word_path_abs, ReadOnly=True)
            
            # Exporter en PDF (17 = wdExportFormatPDF)
            doc.ExportAsFixedFormat(
                temp_path_abs,
                17,  # wdExportFormatPDF
                False,  # OpenAfterExport
                0,  # wdExportOptimizeForPrint
                0,  # Range = wdExportAllDocument
                1,  # From
                1,  # To
                0,  # Item = wdExportDocumentContent
                True,  # IncludeDocProps
                True,  # KeepIRM
                0  # CreateBookmarks = wdExportCreateNoBookmarks
            )
            
            logger.debug(f"Word converti en PDF: {os.path.basename(word_path)}")
            return temp_path
            
        except com_error as e:
            logger.error(f"Erreur COM lors de la conversion Word: {e}")
            # Nettoyer le fichier temporaire si créé
            if os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except OSError:
                    pass
            return None
            
        except Exception as e:
            logger.error(f"Erreur conversion Word: {e}")
            return None
            
        finally:
            # Fermer proprement
            try:
                if doc:
                    doc.Close(False)
            except:
                pass
            try:
                if word:
                    word.Quit()
            except:
                pass
    
    def _excel_to_pdf(self, excel_path: str) -> Optional[str]:
        """Convertit un document Excel (.xls, .xlsx) en PDF via Microsoft Excel"""
        if not self._has_win32com:
            logger.warning("pywin32 requis pour convertir les documents Excel")
            return None
        
        import win32com.client
        from pywintypes import com_error
        
        excel = None
        workbook = None
        temp_path = None
        
        try:
            # Créer un fichier temporaire pour le PDF
            temp_fd, temp_path = tempfile.mkstemp(suffix='.pdf')
            os.close(temp_fd)
            
            # Convertir le chemin en absolu (utiliser des chemins normalisés pour Windows)
            excel_path_abs = os.path.normpath(os.path.abspath(excel_path))
            temp_path_abs = os.path.normpath(os.path.abspath(temp_path))
            
            # Ouvrir Excel
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            
            # Ouvrir le classeur
            workbook = excel.Workbooks.Open(excel_path_abs, ReadOnly=True)
            
            # Exporter en PDF - syntaxe simplifiée
            # Type=0 (xlTypePDF), Filename, Quality=0 (xlQualityStandard)
            workbook.ExportAsFixedFormat(
                Type=0,  # xlTypePDF
                Filename=temp_path_abs,
                Quality=0  # xlQualityStandard
            )
            
            logger.debug(f"Excel converti en PDF: {os.path.basename(excel_path)}")
            return temp_path
            
        except com_error as e:
            logger.error(f"Erreur COM lors de la conversion Excel: {e}")
            # Nettoyer le fichier temporaire si créé
            if temp_path and os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except OSError:
                    pass
            return None
            
        except Exception as e:
            logger.error(f"Erreur conversion Excel: {e}")
            return None
            
        finally:
            # Fermer proprement
            try:
                if workbook:
                    workbook.Close(False)
            except:
                pass
            try:
                if excel:
                    excel.Quit()
            except:
                pass

    def generate_pdf(self, company_name: str, date: str, 
                     email_contents: List[str], attachments: List[str]) -> str:
        """
        Méthode de compatibilité avec l'ancien code.
        
        Args:
            company_name: Nom de l'entreprise
            date: Date au format string
            email_contents: Liste des contenus d'email
            attachments: Liste des chemins de pièces jointes
        
        Returns:
            Chemin du fichier PDF généré
        """
        body = '\n\n'.join(email_contents)
        
        try:
            date_obj = datetime.strptime(date, '%Y%m%d')
        except ValueError:
            date_obj = datetime.now()
        
        return self.generate_email_pdf(
            sender=company_name,
            sender_name=company_name,
            subject="",
            body=body,
            received_time=date_obj,
            attachment_paths=attachments
        )
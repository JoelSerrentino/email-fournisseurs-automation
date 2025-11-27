import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import sys
import json
import threading
from datetime import datetime

# Ajouter le répertoire src au path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

class MainWindow:
    # Couleurs du thème Windows 11 - Service des Finances
    COLORS = {
        # Fonds - Style Windows 11 clair/moderne
        'bg_dark': '#f3f3f3',           # Fond principal gris très clair
        'bg_medium': '#ffffff',          # Cartes en blanc pur
        'bg_light': '#e8e8e8',           # Fond boutons secondaires
        
        # Accent - Bleu professionnel finances
        'accent': '#0078d4',             # Bleu Windows 11
        'accent_hover': '#106ebe',       # Bleu foncé au survol
        'accent_light': '#e5f1fb',       # Bleu très clair pour highlights
        
        # Texte
        'text': '#1a1a1a',               # Texte principal noir/gris foncé
        'text_secondary': '#5c5c5c',     # Texte secondaire gris
        'text_on_accent': '#ffffff',     # Texte sur boutons accent
        
        # États
        'success': '#107c10',            # Vert Microsoft
        'warning': '#ca5010',            # Orange Microsoft
        'error': '#c42b1c',              # Rouge Microsoft
        
        # Champs de saisie
        'entry_bg': '#ffffff',           # Fond champs blanc
        'entry_border': '#d1d1d1',       # Bordure champs gris clair
        'entry_focus': '#0078d4',        # Bordure focus bleu
        
        # Bordures et séparateurs
        'border': '#e0e0e0',             # Bordure cartes gris clair
        'divider': '#edebe9',            # Séparateurs
        
        # Finance - Accents dorés/verts pour les indicateurs
        'finance_gold': '#c19c00',       # Or pour highlights financiers
        'finance_green': '#0e7a0d',      # Vert positif (gains)
        'finance_red': '#bc2f32',        # Rouge négatif (pertes)
    }
    
    # Breakpoints pour le responsive
    BREAKPOINTS = {
        'small': 600,
        'medium': 800,
        'large': 1000
    }

    def __init__(self, master):
        self.master = master
        self.master.title("📧 Email Fournisseurs Automation")
        self.master.geometry("900x700")
        self.master.minsize(500, 450)
        self.master.configure(bg=self.COLORS['bg_dark'])
        
        # Variables
        self.mailbox_var = tk.StringVar()
        self.keywords_var = tk.StringVar()
        self.output_folder_var = tk.StringVar()
        self.outlook_folder_var = tk.StringVar()
        self.category_var = tk.StringVar(value="Traité")
        
        # Variables de progression
        self.progress_var = tk.DoubleVar(value=0)
        self.progress_text_var = tk.StringVar(value="")
        
        # État du traitement
        self.is_processing = False
        self.email_processor = None
        
        # Configuration du layout responsive
        self.master.grid_rowconfigure(0, weight=0)  # Header
        self.master.grid_rowconfigure(1, weight=1)  # Main content
        self.master.grid_rowconfigure(2, weight=0)  # Footer
        self.master.grid_columnconfigure(0, weight=1)
        
        # Configuration du style
        self.setup_styles()
        
        # Création de l'interface
        self.create_header()
        self.create_main_content()
        self.create_footer()
        
        # Chargement des paramètres
        self.load_settings()
        
        # Bind pour le responsive
        self.master.bind('<Configure>', self.on_resize)
        self.current_layout = None

    def setup_styles(self):
        """Configure les styles ttk personnalisés"""
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        # Style du frame principal
        self.style.configure('Main.TFrame', background=self.COLORS['bg_dark'])
        self.style.configure('Card.TFrame', background=self.COLORS['bg_medium'])
        
        # Style des labels
        self.style.configure('Title.TLabel', 
                            background=self.COLORS['bg_dark'],
                            foreground=self.COLORS['text'],
                            font=('Segoe UI', 20, 'bold'))
        
        self.style.configure('Subtitle.TLabel',
                            background=self.COLORS['bg_dark'],
                            foreground=self.COLORS['text_secondary'],
                            font=('Segoe UI', 10))
        
        self.style.configure('Card.TLabel',
                            background=self.COLORS['bg_medium'],
                            foreground=self.COLORS['text'],
                            font=('Segoe UI', 10))
        
        self.style.configure('CardTitle.TLabel',
                            background=self.COLORS['bg_medium'],
                            foreground=self.COLORS['accent'],
                            font=('Segoe UI', 11, 'bold'))
        
        # Style des LabelFrames
        self.style.configure('Card.TLabelframe',
                            background=self.COLORS['bg_medium'],
                            foreground=self.COLORS['text'])
        self.style.configure('Card.TLabelframe.Label',
                            background=self.COLORS['bg_medium'],
                            foreground=self.COLORS['accent'],
                            font=('Segoe UI', 11, 'bold'))
        
        # Style des boutons
        self.style.configure('Accent.TButton',
                            background=self.COLORS['accent'],
                            foreground='white',
                            font=('Segoe UI', 10, 'bold'),
                            padding=(20, 10))
        
        self.style.map('Accent.TButton',
                      background=[('active', self.COLORS['accent_hover'])])
        
        self.style.configure('Secondary.TButton',
                            background=self.COLORS['bg_light'],
                            foreground=self.COLORS['text'],
                            font=('Segoe UI', 9),
                            padding=(10, 5))
        
        self.style.map('Secondary.TButton',
                      background=[('active', self.COLORS['border'])])
        
        # Style des Entry
        self.style.configure('Modern.TEntry',
                            fieldbackground=self.COLORS['entry_bg'],
                            foreground=self.COLORS['text'],
                            insertcolor=self.COLORS['text'],
                            padding=8)

    def on_resize(self, event=None):
        """Gère le redimensionnement de la fenêtre"""
        if event and event.widget == self.master:
            width = event.width
            
            # Déterminer le layout en fonction de la largeur
            if width < self.BREAKPOINTS['small']:
                new_layout = 'small'
            elif width < self.BREAKPOINTS['medium']:
                new_layout = 'medium'
            else:
                new_layout = 'large'
            
            # Mettre à jour le layout si nécessaire
            if new_layout != self.current_layout:
                self.current_layout = new_layout
                self.update_responsive_layout(new_layout)

    def update_responsive_layout(self, layout):
        """Met à jour les éléments en fonction du layout"""
        # Ajuster les paddings selon la taille
        if layout == 'small':
            self.main_padding = 10
            self.card_padding = 8
            self.label_width = 18
        elif layout == 'medium':
            self.main_padding = 20
            self.card_padding = 12
            self.label_width = 20
        else:
            self.main_padding = 30
            self.card_padding = 15
            self.label_width = 22
        
        # Mettre à jour le padding du container principal
        if hasattr(self, 'main_container'):
            self.main_container.configure(padx=self.main_padding, pady=10)

    def create_header(self):
        """Crée l'en-tête de l'application"""
        self.header_frame = tk.Frame(self.master, bg=self.COLORS['bg_dark'])
        self.header_frame.grid(row=0, column=0, sticky='ew', padx=20, pady=(15, 10))
        self.header_frame.grid_columnconfigure(0, weight=1)
        self.header_frame.grid_columnconfigure(1, weight=0)
        
        # Icône et titre
        title_frame = tk.Frame(self.header_frame, bg=self.COLORS['bg_dark'])
        title_frame.grid(row=0, column=0, sticky='w')
        
        self.title_label = tk.Label(title_frame, 
                               text="📧 Email Fournisseurs",
                               font=('Segoe UI', 20, 'bold'),
                               bg=self.COLORS['bg_dark'],
                               fg=self.COLORS['text'])
        self.title_label.pack(anchor=tk.W)
        
        self.subtitle_label = tk.Label(title_frame,
                                  text="Automatisation du traitement des emails",
                                  font=('Segoe UI', 10),
                                  bg=self.COLORS['bg_dark'],
                                  fg=self.COLORS['text_secondary'])
        self.subtitle_label.pack(anchor=tk.W)
        
        # Badge de statut
        status_frame = tk.Frame(self.header_frame, bg=self.COLORS['bg_dark'])
        status_frame.grid(row=0, column=1, sticky='e')
        
        self.status_indicator = tk.Label(status_frame,
                                         text="● Prêt",
                                         font=('Segoe UI', 10),
                                         bg=self.COLORS['bg_dark'],
                                         fg=self.COLORS['success'])
        self.status_indicator.pack()

    def create_main_content(self):
        """Crée le contenu principal avec scroll"""
        # Canvas pour le scroll
        self.canvas = tk.Canvas(self.master, bg=self.COLORS['bg_dark'], 
                                highlightthickness=0)
        self.canvas.grid(row=1, column=0, sticky='nsew', padx=20)
        
        # Scrollbar
        self.scrollbar = ttk.Scrollbar(self.master, orient='vertical', 
                                       command=self.canvas.yview)
        self.scrollbar.grid(row=1, column=1, sticky='ns')
        
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        # Frame intérieur pour le contenu
        self.main_container = tk.Frame(self.canvas, bg=self.COLORS['bg_dark'])
        self.canvas_frame = self.canvas.create_window((0, 0), window=self.main_container, 
                                                       anchor='nw')
        
        # Configuration pour le responsive
        self.main_container.grid_columnconfigure(0, weight=1)
        
        # Sections
        self.create_outlook_section(self.main_container)
        self.create_filter_section(self.main_container)
        self.create_output_section(self.main_container)
        self.create_progress_section(self.main_container)
        self.create_log_section(self.main_container)
        
        # Bindings pour le scroll et le resize
        self.main_container.bind('<Configure>', self.on_frame_configure)
        self.canvas.bind('<Configure>', self.on_canvas_configure)
        
        # Scroll avec la molette
        self.canvas.bind_all('<MouseWheel>', self.on_mousewheel)

    def on_frame_configure(self, event=None):
        """Ajuste la région de scroll"""
        self.canvas.configure(scrollregion=self.canvas.bbox('all'))

    def on_canvas_configure(self, event):
        """Ajuste la largeur du contenu au canvas"""
        self.canvas.itemconfig(self.canvas_frame, width=event.width)

    def on_mousewheel(self, event):
        """Gère le scroll avec la molette"""
        self.canvas.yview_scroll(int(-1*(event.delta/120)), 'units')

    def create_card_frame(self, parent, title, icon="", row=0):
        """Crée un cadre stylisé en forme de carte Windows 11"""
        # Conteneur externe pour l'ombre
        shadow_frame = tk.Frame(parent, bg=self.COLORS['border'])
        shadow_frame.grid(row=row, column=0, sticky='ew', pady=8, padx=5)
        shadow_frame.grid_columnconfigure(0, weight=1)
        
        # Carte principale
        card = tk.Frame(shadow_frame, bg=self.COLORS['bg_medium'], 
                        highlightbackground=self.COLORS['border'],
                        highlightthickness=1)
        card.grid(row=0, column=0, sticky='ew', padx=1, pady=1)
        card.grid_columnconfigure(0, weight=1)
        
        # Barre d'accent en haut de la carte
        accent_bar = tk.Frame(card, bg=self.COLORS['accent'], height=3)
        accent_bar.grid(row=0, column=0, sticky='ew')
        
        # En-tête de la carte
        header = tk.Frame(card, bg=self.COLORS['bg_medium'])
        header.grid(row=1, column=0, sticky='ew', padx=18, pady=(12, 8))
        
        title_label = tk.Label(header,
                               text=f"{icon} {title}",
                               font=('Segoe UI', 11, 'bold'),
                               bg=self.COLORS['bg_medium'],
                               fg=self.COLORS['text'])
        title_label.pack(side=tk.LEFT)
        
        # Contenu de la carte
        content = tk.Frame(card, bg=self.COLORS['bg_medium'])
        content.grid(row=2, column=0, sticky='ew', padx=18, pady=(0, 18))
        content.grid_columnconfigure(1, weight=1)
        
        return content

    def create_form_row(self, parent, label_text, variable, row, has_button=False, 
                        button_text="", button_command=None):
        """Crée une ligne de formulaire responsive"""
        # Label
        label = tk.Label(parent, text=label_text, anchor=tk.W,
                        font=('Segoe UI', 10), bg=self.COLORS['bg_medium'],
                        fg=self.COLORS['text'])
        label.grid(row=row, column=0, sticky='w', pady=5, padx=(0, 10))
        
        # Entry container
        entry_frame = tk.Frame(parent, bg=self.COLORS['bg_medium'])
        entry_frame.grid(row=row, column=1, sticky='ew', pady=5)
        entry_frame.grid_columnconfigure(0, weight=1)
        
        # Conteneur avec bordure pour l'entry
        entry_border = tk.Frame(entry_frame, bg=self.COLORS['entry_border'])
        entry_border.grid(row=0, column=0, sticky='ew')
        entry_border.grid_columnconfigure(0, weight=1)
        
        entry = tk.Entry(entry_border, textvariable=variable,
                        font=('Segoe UI', 10),
                        bg=self.COLORS['entry_bg'],
                        fg=self.COLORS['text'],
                        insertbackground=self.COLORS['accent'],
                        relief=tk.FLAT,
                        highlightthickness=0)
        entry.grid(row=0, column=0, sticky='ew', ipady=8, ipadx=10, padx=1, pady=1)
        
        # Effet focus
        def on_focus_in(e):
            entry_border.configure(bg=self.COLORS['accent'])
        def on_focus_out(e):
            entry_border.configure(bg=self.COLORS['entry_border'])
        entry.bind('<FocusIn>', on_focus_in)
        entry.bind('<FocusOut>', on_focus_out)
        
        if has_button:
            btn = tk.Button(entry_frame, text=button_text,
                           command=button_command,
                           font=('Segoe UI', 9),
                           bg=self.COLORS['accent'],
                           fg=self.COLORS['text_on_accent'],
                           activebackground=self.COLORS['accent_hover'],
                           activeforeground=self.COLORS['text_on_accent'],
                           relief=tk.FLAT, padx=16, pady=6, cursor='hand2',
                           borderwidth=0)
            btn.grid(row=0, column=1, sticky='e', padx=(10, 0))
            
            # Effet hover
            def on_enter(e, b=btn):
                b.configure(bg=self.COLORS['accent_hover'])
            def on_leave(e, b=btn):
                b.configure(bg=self.COLORS['accent'])
            btn.bind('<Enter>', on_enter)
            btn.bind('<Leave>', on_leave)
        
        return entry

    def create_outlook_section(self, parent):
        """Section de configuration Outlook"""
        content = self.create_card_frame(parent, "Configuration Outlook", "📬", row=0)
        
        # Boîte aux lettres
        self.mailbox_entry = self.create_form_row(
            content, "Boîte aux lettres :", self.mailbox_var, 0,
            has_button=True, button_text="Sélectionner", button_command=self.select_mailbox
        )
        
        # Dossier destination Outlook
        self.outlook_folder_entry = self.create_form_row(
            content, "Dossier destination :", self.outlook_folder_var, 1,
            has_button=True, button_text="Sélectionner", button_command=self.select_outlook_folder
        )
        
        # Catégorie
        self.category_entry = self.create_form_row(
            content, "Catégorie après traitement :", self.category_var, 2
        )

    def create_filter_section(self, parent):
        """Section de filtrage"""
        content = self.create_card_frame(parent, "Filtrage des emails", "🔍", row=1)
        
        self.keywords_entry = self.create_form_row(
            content, "Mots clés (séparés par ,) :", self.keywords_var, 0
        )

    def create_output_section(self, parent):
        """Section de sortie PDF"""
        content = self.create_card_frame(parent, "Dossier de sortie PDF", "📁", row=2)
        
        self.output_entry = self.create_form_row(
            content, "Dossier de sortie :", self.output_folder_var, 0,
            has_button=True, button_text="Parcourir", button_command=self.select_output_folder
        )

    def create_progress_section(self, parent):
        """Section de progression et statistiques"""
        # Conteneur avec ombre
        shadow_frame = tk.Frame(parent, bg=self.COLORS['border'])
        shadow_frame.grid(row=3, column=0, sticky='ew', pady=8, padx=5)
        shadow_frame.grid_columnconfigure(0, weight=1)
        
        card = tk.Frame(shadow_frame, bg=self.COLORS['bg_medium'],
                       highlightbackground=self.COLORS['border'],
                       highlightthickness=1)
        card.grid(row=0, column=0, sticky='ew', padx=1, pady=1)
        card.grid_columnconfigure(0, weight=1)
        
        # Barre d'accent verte (finance)
        accent_bar = tk.Frame(card, bg=self.COLORS['finance_green'], height=3)
        accent_bar.grid(row=0, column=0, sticky='ew')
        
        # Contenu
        content = tk.Frame(card, bg=self.COLORS['bg_medium'])
        content.grid(row=1, column=0, sticky='ew', padx=18, pady=15)
        content.grid_columnconfigure(0, weight=1)
        
        # Titre et statut
        header_frame = tk.Frame(content, bg=self.COLORS['bg_medium'])
        header_frame.grid(row=0, column=0, sticky='ew')
        header_frame.grid_columnconfigure(1, weight=1)
        
        tk.Label(header_frame, text="📊 Progression",
                font=('Segoe UI', 11, 'bold'),
                bg=self.COLORS['bg_medium'],
                fg=self.COLORS['text']).grid(row=0, column=0, sticky='w')
        
        self.progress_label = tk.Label(header_frame, textvariable=self.progress_text_var,
                                       font=('Segoe UI', 9),
                                       bg=self.COLORS['bg_medium'],
                                       fg=self.COLORS['text_secondary'])
        self.progress_label.grid(row=0, column=1, sticky='e')
        
        # Barre de progression
        progress_frame = tk.Frame(content, bg=self.COLORS['entry_border'], height=8)
        progress_frame.grid(row=1, column=0, sticky='ew', pady=(10, 8))
        progress_frame.grid_columnconfigure(0, weight=1)
        progress_frame.grid_propagate(False)
        
        self.progress_bar_inner = tk.Frame(progress_frame, bg=self.COLORS['accent'], height=6)
        self.progress_bar_inner.place(x=1, y=1, relwidth=0, height=6)
        
        # Statistiques
        stats_frame = tk.Frame(content, bg=self.COLORS['bg_medium'])
        stats_frame.grid(row=2, column=0, sticky='ew', pady=(5, 0))
        
        # Variables pour les statistiques
        self.stat_total_var = tk.StringVar(value="0")
        self.stat_success_var = tk.StringVar(value="0")
        self.stat_failed_var = tk.StringVar(value="0")
        
        # Total
        tk.Label(stats_frame, text="Total: ",
                font=('Segoe UI', 9),
                bg=self.COLORS['bg_medium'],
                fg=self.COLORS['text_secondary']).pack(side=tk.LEFT)
        tk.Label(stats_frame, textvariable=self.stat_total_var,
                font=('Segoe UI', 9, 'bold'),
                bg=self.COLORS['bg_medium'],
                fg=self.COLORS['text']).pack(side=tk.LEFT, padx=(0, 20))
        
        # Succès
        tk.Label(stats_frame, text="✅ Succès: ",
                font=('Segoe UI', 9),
                bg=self.COLORS['bg_medium'],
                fg=self.COLORS['finance_green']).pack(side=tk.LEFT)
        tk.Label(stats_frame, textvariable=self.stat_success_var,
                font=('Segoe UI', 9, 'bold'),
                bg=self.COLORS['bg_medium'],
                fg=self.COLORS['finance_green']).pack(side=tk.LEFT, padx=(0, 20))
        
        # Échecs
        tk.Label(stats_frame, text="❌ Échecs: ",
                font=('Segoe UI', 9),
                bg=self.COLORS['bg_medium'],
                fg=self.COLORS['finance_red']).pack(side=tk.LEFT)
        tk.Label(stats_frame, textvariable=self.stat_failed_var,
                font=('Segoe UI', 9, 'bold'),
                bg=self.COLORS['bg_medium'],
                fg=self.COLORS['finance_red']).pack(side=tk.LEFT)

    def create_log_section(self, parent):
        """Section du journal - Style Windows 11"""
        # Conteneur avec ombre
        shadow_frame = tk.Frame(parent, bg=self.COLORS['border'])
        shadow_frame.grid(row=4, column=0, sticky='nsew', pady=8, padx=5)
        shadow_frame.grid_columnconfigure(0, weight=1)
        shadow_frame.grid_rowconfigure(0, weight=1)
        
        card = tk.Frame(shadow_frame, bg=self.COLORS['bg_medium'],
                       highlightbackground=self.COLORS['border'],
                       highlightthickness=1)
        card.grid(row=0, column=0, sticky='nsew', padx=1, pady=1)
        card.grid_columnconfigure(0, weight=1)
        card.grid_rowconfigure(2, weight=1)
        
        # Barre d'accent
        accent_bar = tk.Frame(card, bg=self.COLORS['finance_gold'], height=3)
        accent_bar.grid(row=0, column=0, sticky='ew')
        
        # En-tête
        header = tk.Frame(card, bg=self.COLORS['bg_medium'])
        header.grid(row=1, column=0, sticky='ew', padx=18, pady=(12, 8))
        header.grid_columnconfigure(0, weight=1)
        
        tk.Label(header, text="📋 Journal d'activité",
                font=('Segoe UI', 11, 'bold'),
                bg=self.COLORS['bg_medium'],
                fg=self.COLORS['text']).grid(row=0, column=0, sticky='w')
        
        btn_clear = tk.Button(header, text="Effacer",
                             command=self.clear_log,
                             font=('Segoe UI', 8),
                             bg=self.COLORS['bg_light'],
                             fg=self.COLORS['text_secondary'],
                             activebackground=self.COLORS['border'],
                             activeforeground=self.COLORS['text'],
                             relief=tk.FLAT, padx=10, pady=2, cursor='hand2')
        btn_clear.grid(row=0, column=1, sticky='e')
        
        # Zone de texte avec bordure
        log_container = tk.Frame(card, bg=self.COLORS['bg_medium'])
        log_container.grid(row=2, column=0, sticky='nsew', padx=18, pady=(0, 18))
        log_container.grid_columnconfigure(0, weight=1)
        log_container.grid_rowconfigure(0, weight=1)
        
        log_border = tk.Frame(log_container, bg=self.COLORS['entry_border'])
        log_border.grid(row=0, column=0, sticky='nsew')
        log_border.grid_columnconfigure(0, weight=1)
        log_border.grid_rowconfigure(0, weight=1)
        
        self.log_text = tk.Text(log_border, height=8,
                                font=('Cascadia Code', 9),
                                bg=self.COLORS['entry_bg'],
                                fg=self.COLORS['text_secondary'],
                                insertbackground=self.COLORS['accent'],
                                relief=tk.FLAT,
                                wrap=tk.WORD,
                                padx=10, pady=8,
                                state=tk.DISABLED)
        self.log_text.grid(row=0, column=0, sticky='nsew', padx=1, pady=1)
        
        log_scrollbar = tk.Scrollbar(log_border, command=self.log_text.yview)
        log_scrollbar.grid(row=0, column=1, sticky='ns', padx=(0, 1), pady=1)
        self.log_text.configure(yscrollcommand=log_scrollbar.set)

    def create_footer(self):
        """Crée le pied de page avec les boutons d'action"""
        self.footer = tk.Frame(self.master, bg=self.COLORS['bg_dark'])
        self.footer.grid(row=2, column=0, columnspan=2, sticky='ew', padx=20, pady=15)
        self.footer.grid_columnconfigure(0, weight=1)
        
        # Boutons centrés
        btn_frame = tk.Frame(self.footer, bg=self.COLORS['bg_dark'])
        btn_frame.grid(row=0, column=0)
        
        # Bouton Sauvegarder - Style secondaire Windows 11
        self.btn_save = tk.Button(btn_frame, text="💾 Sauvegarder",
                            command=self.save_settings,
                            font=('Segoe UI', 10),
                            bg=self.COLORS['bg_medium'],
                            fg=self.COLORS['text'],
                            activebackground=self.COLORS['bg_light'],
                            activeforeground=self.COLORS['text'],
                            relief=tk.SOLID, borderwidth=1,
                            padx=20, pady=10, cursor='hand2')
        self.btn_save.pack(side=tk.LEFT, padx=8)
        
        # Effet hover bouton sauvegarder
        def on_enter_save(e):
            if self.btn_save['state'] != 'disabled':
                self.btn_save.configure(bg=self.COLORS['bg_light'])
        def on_leave_save(e):
            if self.btn_save['state'] != 'disabled':
                self.btn_save.configure(bg=self.COLORS['bg_medium'])
        self.btn_save.bind('<Enter>', on_enter_save)
        self.btn_save.bind('<Leave>', on_leave_save)
        
        # Bouton Lancer - Style accent Windows 11
        self.btn_start = tk.Button(btn_frame, text="🚀 Lancer le traitement",
                             command=self.start_processing,
                             font=('Segoe UI', 11, 'bold'),
                             bg=self.COLORS['accent'],
                             fg=self.COLORS['text_on_accent'],
                             activebackground=self.COLORS['accent_hover'],
                             activeforeground=self.COLORS['text_on_accent'],
                             relief=tk.FLAT, borderwidth=0,
                             padx=28, pady=12, cursor='hand2')
        self.btn_start.pack(side=tk.LEFT, padx=8)
        
        # Effet hover bouton lancer
        def on_enter_start(e):
            if self.btn_start['state'] != 'disabled':
                self.btn_start.configure(bg=self.COLORS['accent_hover'])
        def on_leave_start(e):
            if self.btn_start['state'] != 'disabled':
                self.btn_start.configure(bg=self.COLORS['accent'])
        self.btn_start.bind('<Enter>', on_enter_start)
        self.btn_start.bind('<Leave>', on_leave_start)
        
        # Bouton Arrêter - Style danger (initialement caché)
        self.btn_stop = tk.Button(btn_frame, text="⏹ Arrêter",
                             command=self.stop_processing,
                             font=('Segoe UI', 10, 'bold'),
                             bg=self.COLORS['error'],
                             fg='#ffffff',
                             activebackground='#c42b1c',
                             activeforeground='#ffffff',
                             relief=tk.FLAT, borderwidth=0,
                             padx=20, pady=10, cursor='hand2')
        # Le bouton est créé mais pas affiché par défaut
        
        # Effet hover bouton arrêter
        def on_enter_stop(e):
            self.btn_stop.configure(bg='#c42b1c')
        def on_leave_stop(e):
            self.btn_stop.configure(bg=self.COLORS['error'])
        self.btn_stop.bind('<Enter>', on_enter_stop)
        self.btn_stop.bind('<Leave>', on_leave_stop)

    def log(self, message, level="info"):
        """Ajoute un message au journal"""
        self.log_text.configure(state=tk.NORMAL)
        timestamp = datetime.now().strftime("%H:%M:%S")
        
        prefix = "ℹ️"
        if level == "success":
            prefix = "✅"
        elif level == "error":
            prefix = "❌"
        elif level == "warning":
            prefix = "⚠️"
        
        self.log_text.insert(tk.END, f"[{timestamp}] {prefix} {message}\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def clear_log(self):
        """Efface le journal"""
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def select_mailbox(self):
        """Sélectionne une boîte aux lettres Outlook"""
        try:
            import win32com.client
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            
            mailboxes = [folder.Name for folder in outlook.Folders]
            
            if mailboxes:
                self.show_selection_dialog("Sélectionner une boîte aux lettres", 
                                          mailboxes, self.mailbox_var)
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible de se connecter à Outlook : {e}")
            self.log(f"Erreur connexion Outlook: {e}", "error")

    def select_outlook_folder(self):
        """Sélectionne un dossier Outlook"""
        try:
            import win32com.client
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            folder = outlook.PickFolder()
            if folder:
                self.outlook_folder_var.set(folder.FolderPath)
                self.log(f"Dossier Outlook sélectionné: {folder.FolderPath}", "success")
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible de sélectionner le dossier : {e}")
            self.log(f"Erreur sélection dossier: {e}", "error")

    def show_selection_dialog(self, title, items, target_var):
        """Affiche une boîte de dialogue de sélection personnalisée"""
        dialog = tk.Toplevel(self.master)
        dialog.title(title)
        dialog.configure(bg=self.COLORS['bg_dark'])
        dialog.transient(self.master)
        dialog.grab_set()
        
        # Taille responsive du dialog
        dialog_width = min(450, self.master.winfo_width() - 50)
        dialog_height = min(400, self.master.winfo_height() - 100)
        dialog.geometry(f"{dialog_width}x{dialog_height}")
        dialog.minsize(300, 250)
        
        # Centrer la fenêtre
        x = self.master.winfo_x() + (self.master.winfo_width() - dialog_width) // 2
        y = self.master.winfo_y() + (self.master.winfo_height() - dialog_height) // 2
        dialog.geometry(f"+{x}+{y}")
        
        # Configuration responsive
        dialog.grid_columnconfigure(0, weight=1)
        dialog.grid_rowconfigure(1, weight=1)
        
        tk.Label(dialog, text=title,
                font=('Segoe UI', 12, 'bold'),
                bg=self.COLORS['bg_dark'],
                fg=self.COLORS['text']).grid(row=0, column=0, pady=15, padx=20, sticky='w')
        
        listbox_frame = tk.Frame(dialog, bg=self.COLORS['bg_medium'])
        listbox_frame.grid(row=1, column=0, sticky='nsew', padx=20, pady=(0, 10))
        listbox_frame.grid_columnconfigure(0, weight=1)
        listbox_frame.grid_rowconfigure(0, weight=1)
        
        listbox = tk.Listbox(listbox_frame,
                            font=('Segoe UI', 10),
                            bg=self.COLORS['entry_bg'],
                            fg=self.COLORS['text'],
                            selectbackground=self.COLORS['accent'],
                            selectforeground='white',
                            relief=tk.FLAT,
                            highlightthickness=0)
        listbox.grid(row=0, column=0, sticky='nsew', padx=2, pady=2)
        
        scrollbar = tk.Scrollbar(listbox_frame, command=listbox.yview)
        scrollbar.grid(row=0, column=1, sticky='ns')
        listbox.configure(yscrollcommand=scrollbar.set)
        
        for item in items:
            listbox.insert(tk.END, item)
        
        def on_select():
            selection = listbox.curselection()
            if selection:
                target_var.set(items[selection[0]])
                self.log(f"Sélectionné: {items[selection[0]]}", "success")
                dialog.destroy()
        
        def on_double_click(event):
            on_select()
        
        listbox.bind('<Double-Button-1>', on_double_click)
        
        btn_select = tk.Button(dialog, text="Sélectionner",
                              command=on_select,
                              font=('Segoe UI', 10, 'bold'),
                              bg=self.COLORS['accent'],
                              fg='white',
                              activebackground=self.COLORS['accent_hover'],
                              activeforeground='white',
                              relief=tk.FLAT, padx=25, pady=8, cursor='hand2')
        btn_select.grid(row=2, column=0, pady=15)

    def select_output_folder(self):
        """Sélectionne le dossier de sortie"""
        folder = filedialog.askdirectory(title="Sélectionner le dossier de sortie")
        if folder:
            self.output_folder_var.set(folder)
            self.log(f"Dossier de sortie: {folder}", "success")

    def save_settings(self):
        """Sauvegarde les paramètres"""
        settings = {
            "mailbox": self.mailbox_var.get(),
            "keywords": self.keywords_var.get(),
            "output_folder": self.output_folder_var.get(),
            "outlook_folder": self.outlook_folder_var.get(),
            "category": self.category_var.get()
        }
        
        config_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), "config")
        os.makedirs(config_dir, exist_ok=True)
        
        config_path = os.path.join(config_dir, "gui_settings.json")
        with open(config_path, "w", encoding="utf-8") as f:
            json.dump(settings, f, indent=4, ensure_ascii=False)
        
        self.log("Paramètres sauvegardés avec succès", "success")
        messagebox.showinfo("Succès", "Paramètres sauvegardés !")

    def load_settings(self):
        """Charge les paramètres sauvegardés"""
        config_path = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 
                                   "config", "gui_settings.json")
        if os.path.exists(config_path):
            try:
                with open(config_path, "r", encoding="utf-8") as f:
                    settings = json.load(f)
                    self.mailbox_var.set(settings.get("mailbox", ""))
                    self.keywords_var.set(settings.get("keywords", ""))
                    self.output_folder_var.set(settings.get("output_folder", ""))
                    self.outlook_folder_var.set(settings.get("outlook_folder", ""))
                    self.category_var.set(settings.get("category", "Traité"))
                self.log("Paramètres chargés", "info")
            except Exception as e:
                self.log(f"Erreur chargement paramètres: {e}", "error")

    def start_processing(self):
        """Démarre le traitement des emails"""
        if not self.mailbox_var.get():
            messagebox.showwarning("Attention", "Veuillez sélectionner une boîte aux lettres.")
            return
        if not self.output_folder_var.get():
            messagebox.showwarning("Attention", "Veuillez sélectionner un dossier de sortie.")
            return
        if not self.keywords_var.get().strip():
            messagebox.showwarning("Attention", "Veuillez entrer au moins un mot clé.")
            return
        
        # Réinitialiser les indicateurs
        self.is_processing = True
        self.progress_var.set(0)
        self.progress_text_var.set("Initialisation...")
        self.stat_total_var.set("0")
        self.stat_success_var.set("0")
        self.stat_failed_var.set("0")
        self.clear_log()
        
        # Mettre à jour l'interface
        self.status_indicator.configure(text="● En cours...", fg=self.COLORS['warning'])
        self.btn_start.pack_forget()
        self.btn_stop.pack(side=tk.LEFT, padx=8)
        self.btn_save.configure(state='disabled')
        
        # Réinitialiser la barre de progression
        if hasattr(self, 'progress_bar_inner'):
            self.progress_bar_inner.place(relx=0, rely=0, relheight=1, relwidth=0)
        
        self.log("Démarrage du traitement...", "info")
        self.log(f"Boîte aux lettres: {self.mailbox_var.get()}", "info")
        self.log(f"Mots clés: {self.keywords_var.get()}", "info")
        self.log(f"Dossier de sortie: {self.output_folder_var.get()}", "info")
        
        # Lancer le traitement dans un thread séparé
        self.processing_thread = threading.Thread(target=self._run_processing, daemon=True)
        self.processing_thread.start()
    
    def _run_processing(self):
        """Exécute le traitement dans un thread séparé"""
        try:
            from email_processor import EmailProcessor
            
            # Créer le processeur avec les callbacks
            self.email_processor = EmailProcessor(
                output_folder=self.output_folder_var.get(),
                progress_callback=self._on_progress,
                log_callback=self._on_log
            )
            
            # Récupérer les paramètres
            keywords = [k.strip() for k in self.keywords_var.get().split(',') if k.strip()]
            mailbox_name = self.mailbox_var.get()
            destination_folder = self.outlook_folder_var.get() if self.outlook_folder_var.get() else None
            category = self.category_var.get() if self.category_var.get() else None
            
            # Lancer le traitement
            stats = self.email_processor.process_emails(
                mailbox_name=mailbox_name,
                keywords=keywords,
                destination_folder=destination_folder,
                category=category
            )
            
            # Traitement terminé
            self.master.after(0, lambda: self._on_processing_complete(stats))
            
        except Exception as e:
            self.master.after(0, lambda: self._on_processing_error(str(e)))
    
    def _on_progress(self, current: int, total: int, message: str):
        """Callback de progression (appelé depuis le thread de traitement)"""
        def update():
            if total > 0:
                progress = int((current / total) * 100)
                self.progress_var.set(progress)
                self.progress_text_var.set(f"{message} ({current}/{total})")
                
                # Mettre à jour la barre visuelle
                if hasattr(self, 'progress_bar_inner'):
                    self.progress_bar_inner.place(relx=0, rely=0, relheight=1, relwidth=progress/100)
            else:
                self.progress_text_var.set(message)
        
        self.master.after(0, update)
    
    def _on_log(self, message: str, level: str = "info"):
        """Callback de log (appelé depuis le thread de traitement)"""
        self.master.after(0, lambda: self.log(message, level))
    
    def _on_processing_complete(self, stats):
        """Appelé quand le traitement est terminé"""
        self.is_processing = False
        
        # Mettre à jour les statistiques
        if stats:
            self.stat_total_var.set(str(stats.total))
            self.stat_success_var.set(str(stats.success))
            self.stat_failed_var.set(str(stats.failed))
        
        # Mettre à jour l'interface
        self.progress_var.set(100)
        self.progress_text_var.set("Traitement terminé !")
        if hasattr(self, 'progress_bar_inner'):
            self.progress_bar_inner.place(relx=0, rely=0, relheight=1, relwidth=1)
        
        self.status_indicator.configure(text="● Terminé", fg=self.COLORS['success'])
        self.btn_stop.pack_forget()
        self.btn_start.pack(side=tk.LEFT, padx=8)
        self.btn_save.configure(state='normal')
        
        self.log("=" * 50, "info")
        self.log("TRAITEMENT TERMINÉ", "success")
        if stats:
            self.log(f"Total: {stats.total} | Succès: {stats.success} | Échecs: {stats.failed}", "info")
        self.log("=" * 50, "info")
        
        # Message de confirmation
        if stats and stats.failed == 0:
            messagebox.showinfo("Succès", f"Traitement terminé !\n{stats.success} email(s) traité(s) avec succès.")
        elif stats:
            messagebox.showwarning("Terminé avec erreurs", 
                                  f"Traitement terminé.\n{stats.success} succès, {stats.failed} échec(s).")
    
    def _on_processing_error(self, error: str):
        """Appelé en cas d'erreur fatale"""
        self.is_processing = False
        
        self.status_indicator.configure(text="● Erreur", fg=self.COLORS['error'])
        self.progress_text_var.set("Erreur !")
        self.btn_stop.pack_forget()
        self.btn_start.pack(side=tk.LEFT, padx=8)
        self.btn_save.configure(state='normal')
        
        self.log(f"ERREUR FATALE: {error}", "error")
        messagebox.showerror("Erreur", f"Une erreur est survenue:\n{error}")
    
    def stop_processing(self):
        """Arrête le traitement en cours"""
        if self.is_processing and self.email_processor:
            self.email_processor.stop()
            self.log("Arrêt demandé...", "warning")
            self.progress_text_var.set("Arrêt en cours...")


if __name__ == "__main__":
    root = tk.Tk()
    app = MainWindow(root)
    root.mainloop()
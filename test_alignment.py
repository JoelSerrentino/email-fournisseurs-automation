"""Script de test pour vérifier l'alignement visuel"""
import tkinter as tk

# Couleurs
COLORS = {
    'bg_dark': '#f3f3f3',
    'bg_medium': '#ffffff',
    'accent': '#0078d4',
    'border': '#e0e0e0',
    'text': '#1a1a1a',
}

root = tk.Tk()
root.title("Test Alignement")
root.geometry("900x400")
root.configure(bg=COLORS['bg_dark'])

# Container principal
main = tk.Frame(root, bg=COLORS['bg_dark'])
main.pack(fill='both', expand=True)

# HEADER avec bordure rouge pour debug
header_frame = tk.Frame(main, bg=COLORS['bg_dark'], highlightbackground='red', highlightthickness=2)
header_frame.grid(row=0, column=0, sticky='ew', padx=20, pady=(15, 10))
header_frame.grid_columnconfigure(0, weight=1)

# Titre
tk.Label(header_frame, text="📧 Email Fournisseurs", 
         font=('Segoe UI', 20, 'bold'),
         bg=COLORS['bg_dark']).grid(row=0, column=0, sticky='w')

# Right frame avec bordure verte pour debug
right_frame = tk.Frame(header_frame, bg='lightgreen', highlightbackground='green', highlightthickness=2)
right_frame.grid(row=0, column=1, columnspan=2, sticky='e', padx=(10, 4))

tk.Label(right_frame, text="● Prêt", bg='lightgreen').pack(side='left', padx=(0, 10))
tk.Button(right_frame, text="❓ Aide", bg=COLORS['accent'], fg='white', padx=15, pady=6).pack(side='left')

# CARTE avec bordure bleue pour debug
shadow_frame = tk.Frame(main, bg=COLORS['border'], highlightbackground='blue', highlightthickness=2)
shadow_frame.grid(row=1, column=0, sticky='ew', pady=8, padx=5)
shadow_frame.grid_columnconfigure(0, weight=1)

card = tk.Frame(shadow_frame, bg=COLORS['bg_medium'])
card.grid(row=0, column=0, sticky='ew', padx=1, pady=1)

content = tk.Frame(card, bg=COLORS['bg_medium'], highlightbackground='orange', highlightthickness=2)
content.grid(row=0, column=0, sticky='ew', padx=18, pady=12)

tk.Label(content, text="📬 Configuration Outlook",
         font=('Segoe UI', 11, 'bold'),
         bg=COLORS['bg_medium']).pack(anchor='w')

# Informations de padding
info_frame = tk.Frame(main, bg=COLORS['bg_dark'])
info_frame.grid(row=2, column=0, sticky='ew', padx=20, pady=10)

info_text = """
PADDING ANALYSIS:
- Rouge : header_frame (padx=20)
- Vert : right_frame (padx=(10, 4)) → 20 + 4 = 24px du bord droit
- Bleu : shadow_frame (padx=5)
- Orange : content (padx=18) → 5 + 1 + 18 = 24px du bord droit

Le bord droit du bouton Aide (vert) devrait être aligné avec le bord droit du contenu (orange).
"""

tk.Label(info_frame, text=info_text, bg=COLORS['bg_dark'], 
         font=('Courier', 9), justify='left', fg=COLORS['text']).pack(anchor='w')

root.mainloop()

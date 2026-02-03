"""
Module pour charger correctement les binaires win32 sur Python 3.13.
Contourne les problèmes d'importation de win32api en préchargeant les DLLs.
"""

import sys
import os
import ctypes
from pathlib import Path

def load_win32_dlls():
    """Précharge les DLLs win32 nécessaires avant d'importer win32com"""
    try:
        # Ajouter le chemin du virtualenv aux chemins de recherche des DLLs
        venv_path = Path(sys.executable).parent
        
        # Essayer de charger les DLLs de base
        dll_names = ['pythoncom313', 'pywintypes313', 'win32api']
        
        for dll_name in dll_names:
            try:
                # Essayer d'abord depuis le virtualenv
                dll_path = str(venv_path / f'{dll_name}.dll')
                if os.path.exists(dll_path):
                    ctypes.CDLL(dll_path)
                else:
                    # Sinon laisser Windows le chercher
                    ctypes.CDLL(dll_name)
            except OSError as e:
                pass  # Silencieusement continuer si le DLL n'est pas trouvé
    except Exception as e:
        pass  # Silencieusement échouer - win32com handle sera désactivé

if __name__ == "__main__":
    load_win32_dlls()

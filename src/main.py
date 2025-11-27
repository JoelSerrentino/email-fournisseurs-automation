from tkinter import Tk
from gui.main_window import MainWindow
from utils.logger import log_message

def main():
    log_message("Application started.")
    root = Tk()
    app = MainWindow(root)
    root.mainloop()

if __name__ == "__main__":
    main()
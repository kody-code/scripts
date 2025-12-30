import tkinter as tk
from gui.main_window import InvoiceApp

if __name__ == "__main__":
    root = tk.Tk()
    app = InvoiceApp(root)
    root.mainloop()
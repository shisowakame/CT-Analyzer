import sys
from gui.main_window import MainWindow
import tkinter as tk


def main():
    root = tk.Tk()
    app = MainWindow(root)
    root.mainloop()


if __name__ == "__main__":
    main() 
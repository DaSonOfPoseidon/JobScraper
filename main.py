# main.py
from gui import CalendarBuddyGUI
import tkinter as tk
import traceback
from tkinterdnd2 import TkinterDnD

if __name__ == "__main__":
    try:
        root = TkinterDnD.Tk()
        gui = CalendarBuddyGUI(root)
        root.mainloop()
    except Exception as e:
        print(f"❌ Unhandled crash: {e}")
        traceback.print_exc()

# main.py
import traceback
from tkinterdnd2 import TkinterDnD
from gui import CalendarBuddyGUI
import sys, os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "../embedded_python/lib")))

if __name__ == "__main__":
    try:
        root = TkinterDnD.Tk()
        gui = CalendarBuddyGUI(root)
        root.mainloop()
    except Exception as e:
        print(f"❌ Unhandled crash: {e}")
        traceback.print_exc()

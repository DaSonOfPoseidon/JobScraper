# main.py
import traceback
from tkinterdnd2 import TkinterDnD
from gui import CalendarBuddyGUI
import sys
import argparse
from utils import ensure_playwright, BROWSERS, __version__, check_for_update
import os

if __name__ == "__main__":
    try:
        parser = argparse.ArgumentParser()
        parser.add_argument(
            '--update',
            action='store_true',
            help="Check for a new version and apply it"
        )
        parser.add_argument(
            '--version',
            action='store_true',
            help="Print current version and exit"
        )
        args, remaining = parser.parse_known_args()

        if args.version:
            print(__version__)
            sys.exit(0)

        if args.update:
            check_for_update()
            sys.exit(0)
            
        os.environ["PLAYWRIGHT_BROWSERS_PATH"] = BROWSERS
        print(f"PLAYWRIGHT_BROWSERS_PATH set to {BROWSERS}")
        ensure_playwright()
        root = TkinterDnD.Tk()
        gui = CalendarBuddyGUI(root)
        root.mainloop()

    except Exception as e:
        print(f"❌ Unhandled crash: {e}")
        traceback.print_exc()

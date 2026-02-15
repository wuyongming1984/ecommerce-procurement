import sys
import os

# Ensure src is importable
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.gui.app import App

if __name__ == "__main__":
    app = App()
    app.mainloop()

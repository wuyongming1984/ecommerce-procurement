import sys
import os

# Add the project root to sys.path to allow imports from src
current_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(current_dir)
sys.path.insert(0, project_root)

# Now we can import from src
try:
    from src.gui.app import App
except ImportError as e:
    # Fallback if run from root as "python src/main.py" without src as package
    # But sys.path insert above should handle it.
    print(f"Import Error: {e}")
    print("Please run this script from the project root using: python src/main.py")
    sys.exit(1)

if __name__ == "__main__":
    app = App()
    app.mainloop()

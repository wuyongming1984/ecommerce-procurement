import PyInstaller.__main__
import os
import shutil
import customtkinter

# Get customtkinter path to include its assets
ctk_path = os.path.dirname(customtkinter.__file__)

# Define paths
current_dir = os.path.dirname(os.path.abspath(__file__))
src_main = os.path.join(current_dir, "src", "main.py")
dist_dir = os.path.join(current_dir, "dist")
build_dir = os.path.join(current_dir, "build")

# Clean previous builds
if os.path.exists(dist_dir):
    shutil.rmtree(dist_dir)
if os.path.exists(build_dir):
    shutil.rmtree(build_dir)

print(f"Building from: {src_main}")
print(f"CustomTkinter path: {ctk_path}")

# Run PyInstaller
PyInstaller.__main__.run([
    src_main,
    '--name=ProcurementGen',
    '--noconfirm',
    '--windowed',  # Hide console for GUI app
    '--onefile',   # Create single exe
    f'--add-data={ctk_path}{os.pathsep}customtkinter', # Add CustomTkinter assets
    # Add other data if needed, e.g. templates? User selects them at runtime, so no need to embed unless they want defaults.
    # The 'stencil' folder is user data, better keep outside.
])

print("Build complete. Executable is in 'dist' folder.")

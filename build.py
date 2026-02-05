import PyInstaller.__main__
import customtkinter
import os

# Get the location of the customtkinter library files
ctk_path = os.path.dirname(customtkinter.__file__)

# Define the build command arguments
args = [
    'gui_app_themed.py',                            # Your main script
    '--name=Neotracker',      # New name to distinguish it
    '--onefile',                          # Create a single file
    
    # --- CHANGE: ENABLE CONSOLE FOR DEBUGGING ---
    '--noconsole',                          # SHOW the black window so we can see errors
    # ------------------------------------------
    
    '--icon=NONE',                        
    
    # Force include theme files
    f'--add-data={ctk_path};customtkinter', 
    
    # Exclude conflicting libraries (Safety measure)
    '--exclude-module=PyQt6',
    '--exclude-module=PySide6',
    '--exclude-module=PyQt5',
    '--exclude-module=PySide2',
    
    '--clean',                            
]

print("🚀 Building DEBUG .exe... (Console Enabled)...")
PyInstaller.__main__.run(args)
print("✅ Build Complete! Check the 'dist' folder for 'PlacementWatcher_Debug.exe'.")
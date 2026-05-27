"""
Test plugin structure and validate basic functionality
"""

import ast
import sys
import os

def check_python_syntax(filepath):
    """Check if a Python file has valid syntax"""
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            content = f.read()
        
        ast.parse(content)
        return True, "OK"
    except SyntaxError as e:
        return False, f"Syntax error: {e}"
    except Exception as e:
        return False, f"Error: {e}"

def main():
    plugin_dir = r"c:\Users\stall\Geotech\Integrated-Geospatial-Analysis-main"
    
    print("QGIS Plugin Validation Report")
    print("=" * 50)
    
    # Check essential files
    essential_files = [
        "__init__.py",
        "integrated_geospatial_analysis.py", 
        "main_dialog.py",
        "metadata.txt",
        "icon.png"
    ]
    
    print("\n1. Essential Files Check:")
    for filename in essential_files:
        filepath = os.path.join(plugin_dir, filename)
        if os.path.exists(filepath):
            print(f"✓ {filename}")
        else:
            print(f"✗ {filename} - MISSING")
    
    # Check Python syntax
    python_files = [
        "__init__.py",
        "integrated_geospatial_analysis.py",
        "main_dialog.py",
        "core/__init__.py",
        "core/map_tools.py", 
        "core/excel_utils.py",
        "gui/__init__.py",
        "gui/professional_theme.py"
    ]
    
    print("\n2. Python Syntax Check:")
    syntax_ok = True
    for filename in python_files:
        filepath = os.path.join(plugin_dir, filename)
        if os.path.exists(filepath):
            is_valid, message = check_python_syntax(filepath)
            if is_valid:
                print(f"✓ {filename}")
            else:
                print(f"✗ {filename} - {message}")
                syntax_ok = False
        else:
            print(f"? {filename} - File not found")
    
    # Check metadata.txt
    print("\n3. Metadata Check:")
    metadata_path = os.path.join(plugin_dir, "metadata.txt")
    if os.path.exists(metadata_path):
        required_fields = ['name', 'qgisMinimumVersion', 'description', 'version', 'author']
        with open(metadata_path, 'r') as f:
            metadata_content = f.read()
        
        for field in required_fields:
            if f"{field}=" in metadata_content:
                print(f"✓ {field} field present")
            else:
                print(f"✗ {field} field missing")
    
    print("\n" + "=" * 50)
    if syntax_ok:
        print("✓ Plugin structure appears valid!")
        print("\nNext steps:")
        print("1. Install the plugin in QGIS plugins folder")
        print("2. Enable it in QGIS Plugin Manager") 
        print("3. Test functionality within QGIS")
    else:
        print("✗ Issues found that need to be fixed")

if __name__ == "__main__":
    main()

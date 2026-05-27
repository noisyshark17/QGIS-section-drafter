"""
Integrated Geospatial Analysis Plugin for QGIS
A comprehensive suite of tools for geotechnical slope analysis including:
- Interactive polyline creation with raster draping
- Elevation profile extraction
- Intersection point analysis
- Excel data processing and merge workflows
- Advanced matplotlib plotting capabilities

Author: Your Name
Email: your.email@example.com
Version: 1.0.0
"""

def classFactory(iface):
    """Load the IntegratedGeospatialTools class from the main plugin file.
    
    Args:
        iface: A reference to the QgisInterface object
        
    Returns:
        IntegratedGeospatialTools: The plugin class instance
    """
    # Import the plugin class from the main plugin file
    from .integrated_geospatial_analysis import IntegratedGeospatialTools
    return IntegratedGeospatialTools(iface)

def name():
    """Return the plugin name"""
    return "Integrated Geospatial Analysis"

def description():
    """Return the plugin description"""
    return ("Comprehensive geotechnical analysis suite including polyline creation with raster draping, "
            "elevation profile extraction, intersection point analysis, Excel merge workflows, and "
            "matplotlib plotting capabilities for professional geotechnical engineering applications.")

def version():
    """Return the plugin version"""
    return "1.0.0"

def qgisMinimumVersion():
    """Return the minimum QGIS version required"""
    return "3.16"

def author():
    """Return the plugin author"""
    return "Your Name"

def email():
    """Return the author email"""
    return "your.email@example.com"

def icon():
    """Return the plugin icon filename"""
    return "icon.png"

def experimental():
    """Return whether this is an experimental plugin"""
    return False

def deprecated():
    """Return whether this plugin is deprecated"""
    return False

def homepage():
    """Return the plugin homepage URL"""
    return "https://github.com/yourusername/integrated-geospatial-analysis"

def repository():
    """Return the plugin repository URL"""
    return "https://github.com/yourusername/integrated-geospatial-analysis"

def tracker():
    """Return the bug tracker URL"""
    return "https://github.com/yourusername/integrated-geospatial-analysis/issues"

def category():
    """Return the plugin category"""
    return "Analysis"

def tags():
    """Return plugin tags"""
    return "geotechnical,slope analysis,profiles,intersections,excel,plotting"

def changelog():
    """Return the changelog"""
    return """
    Version 1.0.0:
    - Initial release
    - Interactive polyline creation with raster draping
    - Elevation profile extraction with configurable sampling
    - Intersection point analysis with multiple geometry types
    - Excel data processing and merge workflows
    - Advanced matplotlib plotting with customizable appearance
    - Support for multiple data formats (Excel, CSV, GeoPackage, Shapefile)
    - Professional UI with tabbed interface
    - Comprehensive error handling and logging
    """

def about():
    """Return detailed about information"""
    return """
    Integrated Geospatial Analysis Plugin

    This plugin provides a comprehensive suite of tools specifically designed for 
    geotechnical engineering and slope analysis applications. It combines multiple 
    geospatial analysis capabilities into a single, integrated workflow.

    Key Features:
    • Interactive polyline creation with automatic raster elevation draping
    • Flexible elevation profile extraction with customizable sampling intervals
    • Advanced intersection point analysis supporting multiple geometry types
    • Excel data processing workflows with duplicate prevention
    • Publication-ready matplotlib plotting with extensive customization options
    • Support for multiple output formats (Excel, CSV, GeoPackage, Shapefile)
    • Professional tabbed interface with context-sensitive help
    • Robust error handling with detailed logging

    Technical Requirements:
    • QGIS 3.16 or higher
    • Python packages: pandas, matplotlib, xlsxwriter, openpyxl (optional but recommended)
    • Sufficient system memory for processing large datasets

    Workflow:
    1. Create polylines interactively on the map
    2. Extract elevation profiles from line features
    3. Find intersection points with other geometric features
    4. Process and merge data using Excel workflows
    5. Generate publication-ready plots with custom styling

    The plugin is designed to handle real-world geotechnical data and provides 
    professional-grade output suitable for engineering reports and analysis.
    """

# Plugin metadata for internal use
__version__ = version()
__author__ = author()
__email__ = email()
__license__ = "GPL v3"
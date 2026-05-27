"""
Main plugin class for Integrated Geospatial Analysis
"""

from PyQt5.QtCore import QTranslator, QCoreApplication
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QAction
from qgis.core import QgsApplication
import os.path

from .main_dialog import CombinedGeospatialToolDialog


class IntegratedGeospatialTools:
    """QGIS Plugin Implementation for Integrated Geospatial Analysis"""

    def __init__(self, iface):
        """Constructor.

        Args:
            iface (QgsInterface): A reference to the QgisInterface object
        """
        # Save reference to the QGIS interface
        self.iface = iface
        
        # Initialize plugin directory
        self.plugin_dir = os.path.dirname(__file__)
        
        # Initialize locale
        locale = QgsApplication.locale()[:2]
        locale_path = os.path.join(
            self.plugin_dir,
            'i18n',
            'IntegratedGeospatialTools_{}.qm'.format(locale))

        if os.path.exists(locale_path):
            self.translator = QTranslator()
            self.translator.load(locale_path)
            QCoreApplication.installTranslator(self.translator)

        # Declare instance attributes
        self.actions = []
        self.menu = self.tr('&Integrated Geospatial Analysis')
        
        # Check if plugin was started the first time in current QGIS session
        # Must be set in initGui() to survive plugin reloads
        self.first_start = None
        
        # Store dialog reference
        self.dialog = None

    def tr(self, message):
        """Get the translation for a string using Qt translation API.

        Args:
            message (str): String for translation

        Returns:
            str: Translated string
        """
        return QCoreApplication.translate('IntegratedGeospatialTools', message)

    def add_action(
        self,
        icon_path,
        text,
        callback,
        enabled_flag=True,
        add_to_menu=True,
        add_to_toolbar=True,
        status_tip=None,
        whats_this=None,
        parent=None):
        """Add a toolbar icon to the toolbar.

        Args:
            icon_path (str): Path to the icon for this action
            text (str): Text that should be shown in menu items for this action
            callback (function): Function to be called when the action is triggered
            enabled_flag (bool): A flag indicating if the action should be enabled by default
            add_to_menu (bool): Flag indicating whether the action should also be added to the menu
            add_to_toolbar (bool): Flag indicating whether the action should also be added to the toolbar
            status_tip (str): Optional text to show in a popup when mouse pointer hovers over the action
            whats_this (str): Optional text to show in the status bar when the mouse pointer hovers over the action
            parent: Parent widget for the new action

        Returns:
            QAction: The action that was created
        """
        icon = QIcon(icon_path)
        action = QAction(icon, text, parent)
        action.triggered.connect(callback)
        action.setEnabled(enabled_flag)

        if status_tip is not None:
            action.setStatusTip(status_tip)

        if whats_this is not None:
            action.setWhatsThis(whats_this)

        if add_to_toolbar:
            # Adds plugin icon to Plugins toolbar
            self.iface.addToolBarIcon(action)

        if add_to_menu:
            self.iface.addPluginToMenu(
                self.menu,
                action)

        self.actions.append(action)

        return action

    def initGui(self):
        """Create the menu entries and toolbar icons inside the QGIS GUI."""
        
        icon_path = os.path.join(self.plugin_dir, 'icon.png')
        self.add_action(
            icon_path,
            text=self.tr('Integrated Geospatial Analysis'),
            callback=self.run,
            parent=self.iface.mainWindow())

        # Will be set False in run()
        self.first_start = True

    def unload(self):
        """Removes the plugin menu item and icon from QGIS GUI."""
        for action in self.actions:
            self.iface.removePluginMenu(
                self.tr('&Integrated Geospatial Analysis'),
                action)
            self.iface.removeToolBarIcon(action)
            
        # Clean up dialog
        if self.dialog:
            self.dialog.close()
            self.dialog = None

    def run(self):
        """Run method that performs all the real work"""
        
        # Create the dialog with elements (after translation) and keep reference
        # Only create new dialog if one doesn't exist
        if self.dialog is None:
            self.dialog = CombinedGeospatialToolDialog(self.iface)
            
        # Show the dialog
        self.dialog.show()
        # Run the dialog event loop - this allows the dialog to stay open
        # but doesn't block QGIS like exec_() would
        self.dialog.raise_()
        self.dialog.activateWindow()

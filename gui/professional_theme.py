"""
Professional Theme System for Integrated Geospatial Tools
Provides consistent styling and modern appearance
"""

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont, QPalette, QColor
from PyQt5.QtWidgets import QApplication


class ProfessionalTheme:
    """Professional color scheme and styling constants"""
    
    # Color Palette - Modern Professional
    COLORS = {
        # Primary Colors
        'primary': '#2E7D32',           # Dark Green
        'primary_light': '#4CAF50',     # Light Green  
        'primary_dark': '#1B5E20',      # Darker Green
        
        # Secondary Colors
        'secondary': '#1976D2',         # Blue
        'secondary_light': '#42A5F5',   # Light Blue
        'secondary_dark': '#0D47A1',    # Dark Blue
        
        # Accent Colors
        'accent': '#FF9800',            # Orange
        'accent_light': '#FFB74D',      # Light Orange
        'accent_dark': '#F57C00',       # Dark Orange
        
        # Neutral Colors
        'background': '#FAFAFA',        # Very Light Gray
        'surface': '#FFFFFF',           # White
        'surface_variant': '#F5F5F5',   # Light Gray
        'outline': '#E0E0E0',           # Border Gray
        
        # Text Colors
        'on_primary': '#FFFFFF',        # White text on primary
        'on_secondary': '#FFFFFF',      # White text on secondary
        'on_surface': '#212121',        # Dark text on surface
        'on_surface_variant': '#757575', # Gray text
        
        # Status Colors
        'success': '#4CAF50',           # Green
        'warning': '#FF9800',           # Orange
        'error': '#F44336',             # Red
        'info': '#2196F3',              # Blue
        
        # Status Backgrounds
        'success_bg': '#E8F5E8',        # Light Green
        'warning_bg': '#FFF3E0',        # Light Orange
        'error_bg': '#FFEBEE',          # Light Red
        'info_bg': '#E3F2FD',           # Light Blue
    }
    
    # Typography Scale
    FONTS = {
        'display_large': {'size': 24, 'weight': 'bold'},
        'display_medium': {'size': 20, 'weight': 'bold'},
        'display_small': {'size': 18, 'weight': 'bold'},
        'headline_large': {'size': 16, 'weight': '600'},
        'headline_medium': {'size': 14, 'weight': '600'},
        'headline_small': {'size': 13, 'weight': '600'},
        'title_large': {'size': 14, 'weight': '500'},
        'title_medium': {'size': 13, 'weight': '500'},
        'title_small': {'size': 12, 'weight': '500'},
        'body_large': {'size': 12, 'weight': 'normal'},
        'body_medium': {'size': 11, 'weight': 'normal'},
        'body_small': {'size': 10, 'weight': 'normal'},
        'label_large': {'size': 11, 'weight': '500'},
        'label_medium': {'size': 10, 'weight': '500'},
        'label_small': {'size': 9, 'weight': '500'},
    }
    
    # Spacing Scale (in pixels)
    SPACING = {
        'xs': 4,    # Extra small
        'sm': 8,    # Small  
        'md': 12,   # Medium
        'lg': 16,   # Large
        'xl': 20,   # Extra large
        'xxl': 24,  # Extra extra large
        'xxxl': 32, # Triple extra large
    }
    
    # Border Radius
    RADIUS = {
        'sm': 4,    # Small radius
        'md': 6,    # Medium radius  
        'lg': 8,    # Large radius
        'xl': 12,   # Extra large radius
        'round': 50, # Fully rounded
    }
    
    # Shadows
    SHADOWS = {
        'none': 'none',
        'sm': '0px 1px 3px rgba(0, 0, 0, 0.12)',
        'md': '0px 4px 6px rgba(0, 0, 0, 0.1)',
        'lg': '0px 8px 15px rgba(0, 0, 0, 0.1)',
        'xl': '0px 16px 24px rgba(0, 0, 0, 0.1)',
    }

    @classmethod
    def get_button_style(cls, variant='primary', size='medium'):
        """Get professional button styling"""
        base_style = f"""
            QPushButton {{
                border: none;
                border-radius: {cls.RADIUS['md']}px;
                font-weight: 500;
                text-align: center;
                cursor: pointer;
                transition: all 0.2s ease;
            }}
        """
        
        # Size variants
        if size == 'small':
            size_style = f"padding: {cls.SPACING['xs']}px {cls.SPACING['md']}px; font-size: {cls.FONTS['label_small']['size']}px;"
        elif size == 'large':
            size_style = f"padding: {cls.SPACING['md']}px {cls.SPACING['xl']}px; font-size: {cls.FONTS['title_medium']['size']}px;"
        else:  # medium
            size_style = f"padding: {cls.SPACING['sm']}px {cls.SPACING['lg']}px; font-size: {cls.FONTS['body_large']['size']}px;"
        
        # Color variants
        if variant == 'primary':
            color_style = f"""
                background-color: {cls.COLORS['primary']};
                color: {cls.COLORS['on_primary']};
            }}
            QPushButton:hover {{
                background-color: {cls.COLORS['primary_light']};
            }}
            QPushButton:pressed {{
                background-color: {cls.COLORS['primary_dark']};
            }}
            QPushButton:disabled {{
                background-color: {cls.COLORS['outline']};
                color: {cls.COLORS['on_surface_variant']};
            """
        elif variant == 'secondary':
            color_style = f"""
                background-color: transparent;
                color: {cls.COLORS['primary']};
                border: 2px solid {cls.COLORS['primary']};
            }}
            QPushButton:hover {{
                background-color: {cls.COLORS['success_bg']};
            }}
            QPushButton:pressed {{
                background-color: {cls.COLORS['primary']};
                color: {cls.COLORS['on_primary']};
            }}
            QPushButton:disabled {{
                border-color: {cls.COLORS['outline']};
                color: {cls.COLORS['on_surface_variant']};
            """
        elif variant == 'danger':
            color_style = f"""
                background-color: {cls.COLORS['error']};
                color: white;
            }}
            QPushButton:hover {{
                background-color: #E53935;
            }}
            QPushButton:pressed {{
                background-color: #C62828;
            }}
            QPushButton:disabled {{
                background-color: {cls.COLORS['outline']};
                color: {cls.COLORS['on_surface_variant']};
            """
        else:  # text variant
            color_style = f"""
                background-color: transparent;
                color: {cls.COLORS['primary']};
            }}
            QPushButton:hover {{
                background-color: {cls.COLORS['success_bg']};
            }}
            QPushButton:pressed {{
                background-color: {cls.COLORS['outline']};
            }}
            QPushButton:disabled {{
                color: {cls.COLORS['on_surface_variant']};
            """
        
        return base_style + "QPushButton { " + size_style + color_style
    
    @classmethod
    def get_input_style(cls):
        """Get professional input field styling"""
        return f"""
            QLineEdit, QSpinBox, QDoubleSpinBox, QComboBox {{
                border: 2px solid {cls.COLORS['outline']};
                border-radius: {cls.RADIUS['md']}px;
                padding: {cls.SPACING['sm']}px {cls.SPACING['md']}px;
                font-size: {cls.FONTS['body_large']['size']}px;
                background-color: {cls.COLORS['surface']};
                color: {cls.COLORS['on_surface']};
            }}
            QLineEdit:focus, QSpinBox:focus, QDoubleSpinBox:focus, QComboBox:focus {{
                border-color: {cls.COLORS['primary']};
                outline: none;
            }}
            QLineEdit:disabled, QSpinBox:disabled, QDoubleSpinBox:disabled, QComboBox:disabled {{
                background-color: {cls.COLORS['surface_variant']};
                color: {cls.COLORS['on_surface_variant']};
                border-color: {cls.COLORS['outline']};
            }}
            QLineEdit[error="true"] {{
                border-color: {cls.COLORS['error']};
                background-color: {cls.COLORS['error_bg']};
            }}
        """
    
    @classmethod
    def get_tab_style(cls):
        """Get professional tab widget styling"""
        return f"""
            QTabWidget::pane {{
                border: 1px solid {cls.COLORS['outline']};
                border-radius: {cls.RADIUS['lg']}px;
                background-color: {cls.COLORS['surface']};
                padding: {cls.SPACING['xs']}px;
                margin-top: 2px;
            }}
            QTabBar::tab {{
                background-color: {cls.COLORS['surface_variant']};
                border: 1px solid {cls.COLORS['outline']};
                border-bottom: none;
                border-radius: {cls.RADIUS['md']}px {cls.RADIUS['md']}px 0 0;
                padding: {cls.SPACING['md']}px {cls.SPACING['xl']}px;
                margin-right: 2px;
                font-weight: {cls.FONTS['title_medium']['weight']};
                font-size: {cls.FONTS['title_medium']['size']}px;
                color: {cls.COLORS['on_surface_variant']};
            }}
            QTabBar::tab:selected {{
                background-color: {cls.COLORS['surface']};
                color: {cls.COLORS['primary']};
                border-bottom: 3px solid {cls.COLORS['primary']};
                font-weight: 600;
            }}
            QTabBar::tab:hover {{
                background-color: {cls.COLORS['success_bg']};
                color: {cls.COLORS['primary']};
            }}
        """
    
    @classmethod
    def get_group_box_style(cls):
        """Get professional group box styling"""
        return f"""
            QGroupBox {{
                font-weight: {cls.FONTS['headline_small']['weight']};
                font-size: {cls.FONTS['headline_small']['size']}px;
                color: {cls.COLORS['on_surface']};
                border: 2px solid {cls.COLORS['outline']};
                border-radius: {cls.RADIUS['lg']}px;
                margin-top: {cls.SPACING['md']}px;
                padding-top: {cls.SPACING['sm']}px;
                background-color: {cls.COLORS['surface']};
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 0 {cls.SPACING['sm']}px;
                background-color: {cls.COLORS['surface']};
                color: {cls.COLORS['primary']};
            }}
        """
    
    @classmethod
    def get_status_card_style(cls, status='info'):
        """Get status card styling"""
        status_colors = {
            'success': (cls.COLORS['success'], cls.COLORS['success_bg']),
            'warning': (cls.COLORS['warning'], cls.COLORS['warning_bg']),
            'error': (cls.COLORS['error'], cls.COLORS['error_bg']),
            'info': (cls.COLORS['info'], cls.COLORS['info_bg']),
        }
        
        border_color, bg_color = status_colors.get(status, status_colors['info'])
        
        return f"""
            QFrame {{
                background-color: {bg_color};
                border: 2px solid {border_color};
                border-radius: {cls.RADIUS['lg']}px;
                padding: {cls.SPACING['md']}px;
                margin: {cls.SPACING['xs']}px;
            }}
            QLabel {{
                color: {cls.COLORS['on_surface']};
                font-size: {cls.FONTS['body_large']['size']}px;
                background-color: transparent;
                border: none;
            }}
        """
    
    @classmethod
    def get_progress_style(cls):
        """Get professional progress bar styling"""
        return f"""
            QProgressBar {{
                border: 2px solid {cls.COLORS['outline']};
                border-radius: {cls.RADIUS['lg']}px;
                background-color: {cls.COLORS['surface_variant']};
                text-align: center;
                font-weight: 500;
                font-size: {cls.FONTS['body_medium']['size']}px;
                color: {cls.COLORS['primary']};
                height: 24px;
            }}
            QProgressBar::chunk {{
                background-color: {cls.COLORS['primary']};
                border-radius: {cls.RADIUS['md']}px;
                margin: 2px;
            }}
        """

    @classmethod
    def apply_global_style(cls, app):
        """Apply professional theme to the entire application"""
        style_sheet = f"""
            QApplication {{
                font-family: "Segoe UI", "Roboto", "Helvetica Neue", Arial, sans-serif;
                background-color: {cls.COLORS['background']};
                color: {cls.COLORS['on_surface']};
            }}
            
            QDialog, QMainWindow {{
                background-color: {cls.COLORS['background']};
                color: {cls.COLORS['on_surface']};
            }}
            
            {cls.get_button_style()}
            {cls.get_input_style()}
            {cls.get_tab_style()}
            {cls.get_group_box_style()}
            {cls.get_progress_style()}
            
            QScrollBar:vertical {{
                background-color: {cls.COLORS['surface_variant']};
                width: 12px;
                border-radius: 6px;
            }}
            QScrollBar::handle:vertical {{
                background-color: {cls.COLORS['on_surface_variant']};
                border-radius: 6px;
                min-height: 20px;
            }}
            QScrollBar::handle:vertical:hover {{
                background-color: {cls.COLORS['primary']};
            }}
        """
        
        app.setStyleSheet(style_sheet)

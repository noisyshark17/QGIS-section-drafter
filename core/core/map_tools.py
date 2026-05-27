"""
Map tools for the Integrated Geospatial Tools plugin
"""

from PyQt5.QtCore import pyqtSignal, Qt
from PyQt5.QtGui import QColor
from qgis.core import QgsMessageLog, Qgis, QgsWkbTypes
from qgis.gui import QgsMapToolEmitPoint, QgsRubberBand


class PolylineMapTool(QgsMapToolEmitPoint):
    """Map tool for drawing polylines"""
    
    finished = pyqtSignal(list)
    
    def __init__(self, canvas):
        super().__init__(canvas)
        self.canvas = canvas
        self.points = []
        self.rubber_band = QgsRubberBand(canvas, QgsWkbTypes.LineGeometry)
        self.rubber_band.setColor(QColor(255, 0, 0))
        self.rubber_band.setWidth(2)
        
    def canvasReleaseEvent(self, event):
        """Handle mouse clicks"""
        try:
            if event.button() == Qt.LeftButton:
                point = self.toMapCoordinates(event.pos())
                self.points.append(point)
                self.update_rubber_band()
                
            elif event.button() == Qt.RightButton:
                if len(self.points) >= 2:
                    self.finished.emit(self.points)
                self.reset()
        except Exception as e:
            QgsMessageLog.logMessage(f"Error in map tool: {str(e)}", "IntegratedProfileAnalyzer", Qgis.Warning)
            
    def update_rubber_band(self):
        """Update the rubber band display"""
        try:
            self.rubber_band.reset(QgsWkbTypes.LineGeometry)
            for point in self.points:
                self.rubber_band.addPoint(point)
        except Exception as e:
            QgsMessageLog.logMessage(f"Error updating rubber band: {str(e)}", "IntegratedProfileAnalyzer", Qgis.Warning)
            
    def reset(self):
        """Reset the tool"""
        self.points = []
        self.rubber_band.reset()
        
    def deactivate(self):
        """Clean up when tool is deactivated"""
        self.reset()
        super().deactivate()

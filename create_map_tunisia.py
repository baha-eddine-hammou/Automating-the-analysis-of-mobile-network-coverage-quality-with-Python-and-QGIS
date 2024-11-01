from qgis.core import QgsProject, QgsVectorLayer, QgsSymbol, QgsSingleSymbolRenderer
from PyQt5.QtWidgets import QFileDialog, QApplication
from PyQt5.QtGui import QColor
import os
import sys

# Function to get the map folder path
def get_map_folder_path():
    app = QApplication(sys.argv)
    map_folder_path = QFileDialog.getExistingDirectory(None, "Select Map Folder")
    app.exit()
    return map_folder_path

# Get the map folder path
map_folder_path = get_map_folder_path()

# Ensure a map folder was selected
if not map_folder_path:
    print("No map folder selected. Exiting.")
    sys.exit()

print(f"Selected map folder: {map_folder_path}")

# Define color mapping for specific files
color_mapping = {
    'water': 'cyan',
    'secondary_road': 'yellow',
    'street': 'purple',
    'railway': 'brown',
    'main_road': 'red',
    'highway': 'green'
}

# Define a default color for files that do not match any of the specified categories
default_color = 'gray'

# Iterate over all files in the folder
for filename in os.listdir(map_folder_path):
    # Check if the file is a .shp file
    if filename.endswith('.shp'):
        # Construct the full file path
        file_path = os.path.join(map_folder_path, filename)
        
        # Create a vector layer from the .shp file
        layer = QgsVectorLayer(file_path, filename, 'ogr')
        
        # Check if the layer is valid
        if layer.isValid():
            # Determine the color based on the filename
            color = default_color
            for key, mapped_color in color_mapping.items():
                if key in filename:
                    color = mapped_color
                    break
            
            # Create and apply the symbol
            symbol = QgsSymbol.defaultSymbol(layer.geometryType())
            symbol.setColor(QColor(color))
            
            # Create a single symbol renderer
            renderer = QgsSingleSymbolRenderer(symbol)
            layer.setRenderer(renderer)
            
            # Add the layer to the QGIS project
            QgsProject.instance().addMapLayer(layer)
            print(f'Layer {filename} added and colored {color}.')
        else:
            print(f'Failed to load layer {filename}.')

# Function to save the QGIS project
def save_project():
    app = QApplication(sys.argv)
    options = QFileDialog.Options()
    options |= QFileDialog.DontUseNativeDialog
    file_dialog = QFileDialog()
    file_dialog.setAcceptMode(QFileDialog.AcceptSave)
    file_dialog.setDefaultSuffix('qgz')
    file_dialog.setNameFilters(["QGIS Projects (*.qgz)", "All Files (*)"])
    
    project_path, _ = file_dialog.getSaveFileName(None, "Save QGIS Project", "", "QGIS Projects (*.qgz);;All Files (*)", options=options)
    
    if project_path:
        # Ensure the file path has the correct .qgz extension
        if not project_path.lower().endswith('.qgz'):
            project_path += '.qgz'
        
        # Save the QGIS project
        QgsProject.instance().write(project_path)
        print(f'Project saved to {project_path}.')
    else:
        print('Project save canceled.')
    app.exit()

# Save the project
save_project()

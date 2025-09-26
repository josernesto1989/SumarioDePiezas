# Este módulo actúa como el "pegamento" entre la Vista y el Modelo.
# Maneja la lógica de la UI y los datos de la aplicación.
from PyQt6.QtCore import QObject, QStringListModel, pyqtSignal
from PyQt6.QtWidgets import QFileDialog
from excel_processor import ExcelProcessor


class MainViewModel(QObject):
    """
    El ViewModel de la aplicación.
    Gestiona la lógica de la UI, los modelos de datos y se comunica con el ExcelProcessor.
    """
    resumeCompleted = pyqtSignal()
    
    def __init__(self):
        super().__init__()
        self.file_list = []
        self.file_list_model = QStringListModel()

    def load_files(self, file_paths):
        """
        Añade las rutas de los archivos a la lista y actualiza el modelo.
        """
        # Filtra para aceptar solo archivos .xlsx
        excel_files = [f for f in file_paths if f.endswith('.xlsx')]

        if excel_files:
            self.file_list.extend(excel_files)
            self.file_list_model.setStringList(self.file_list)
        else:
            print("No se seleccionaron archivos .xlsx válidos.")

    def handle_load_button(self):
        """
        Manejador para el botón "Cargar Archivos".
        Abre un diálogo de selección de archivos.
        """
        file_dialog = QFileDialog()
        file_paths, _ = file_dialog.getOpenFileNames(
            None,
            "Seleccionar Archivos de Excel",
            "",
            "Archivos de Excel (*.xlsx);;Todos los archivos (*)"
        )
        if file_paths:
            self.load_files(file_paths)
    
    def handle_resumir(self):
        file_dialog = QFileDialog()
        file_path = file_dialog.getExistingDirectory(None, "Seleccionar Directorio de Salida")
        if file_path:
            excel_processor = ExcelProcessor()
            excel_processor.process_file(self.file_list, file_path)
            self.resumeCompleted.emit()
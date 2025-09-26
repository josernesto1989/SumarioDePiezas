# Este módulo define la ventana principal (la Vista).
# Se encarga de la interfaz de usuario y la comunicación con el ViewModel.
from PyQt6 import QtWidgets, uic
from PyQt6.QtCore import QUrl, Qt, pyqtSignal
from PyQt6.QtGui import QAction, QKeySequence

class MainWindow(QtWidgets.QMainWindow):
    """
    La ventana principal de la aplicación.
    Carga el diseño de Qt Designer y maneja los eventos de la UI.
    """
    def __init__(self, view_model):
        super().__init__()
        self.view_model = view_model

        # Cargar el diseño de la UI desde el archivo .ui generado por Qt Designer.
        try:
            uic.loadUi('main_window.ui', self)
        except FileNotFoundError:
            # En caso de que el archivo .ui no exista, se crea una interfaz básica.
            # Este es un buen fallback, pero el usuario debe crear el archivo .ui
            # para una experiencia completa.
            self.setWindowTitle("Sumario de Piezas")
            self.setGeometry(100, 100, 800, 600)

            # Widgets de la interfaz básica
            self.dropLabel = QtWidgets.QLabel("Arrastra y suelta tus archivos aquí", self)
            self.dropLabel.setAlignment(Qt.AlignmentFlag.AlignCenter)
            self.fileListView = QtWidgets.QListView(self)
            self.loadFilesButton = QtWidgets.QPushButton("Cargar Archivos", self)

            # Diseño de la interfaz básica
            layout = QtWidgets.QVBoxLayout()
            layout.addWidget(self.dropLabel)
            layout.addWidget(self.fileListView)
            layout.addWidget(self.loadFilesButton)

            central_widget = QtWidgets.QWidget()
            central_widget.setLayout(layout)
            self.setCentralWidget(central_widget)

        # Habilitar el drag and drop para toda la ventana
        self.setAcceptDrops(True)

        # Conectar los widgets con los métodos del ViewModel
        self.pushButtonAdicionar.clicked.connect(self.view_model.handle_load_button)
        self.listView.setModel(self.view_model.file_list_model)
        self.pushButtonResumir.clicked.connect(self.view_model.handle_resumir)

        # Configurar atajo de teclado para abrir archivos (Ctrl+O)
        open_action = QAction(self)
        open_action.setShortcut(QKeySequence("Ctrl+O"))
        open_action.triggered.connect(self.view_model.handle_load_button)
        self.addAction(open_action)
        self.view_model.resumeCompleted.connect(self.showFinishMessage)
        
    def dragEnterEvent(self, event):
        """
        Manejador para el evento de entrada de drag and drop.
        Acepta el evento si los datos MIME son URLs (archivos).
        """
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event) -> None:
        """
        Manejador para el evento de soltar.
        Extrae las rutas de los archivos y las envía al ViewModel.
        """
        urls = event.mimeData().urls()
        file_paths = [url.toLocalFile() for url in urls if url.isLocalFile()]
        if file_paths:
            self.view_model.load_files(file_paths)
            event.acceptProposedAction()
        else:
            event.ignore()
    
    def showFinishMessage(self):
        """
        Muestra un mensaje de información al usuario.
        """
        QtWidgets.QMessageBox.information(self, "Completado", "El proceso de resumen ha finalizado.")
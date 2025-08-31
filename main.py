# Este es el punto de entrada principal de la aplicación.
# Crea la instancia de la aplicación, el ViewModel y la Vista.
import sys
from PyQt6 import QtWidgets

from main_window import MainWindow
from main_view_model import MainViewModel

def main():
    """
    Función principal para iniciar la aplicación.
    """
    app = QtWidgets.QApplication(sys.argv)

    # Crear una instancia del ViewModel y la Vista
    view_model = MainViewModel()
    main_window = MainWindow(view_model)

    # Mostrar la ventana y ejecutar el bucle de eventos
    main_window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
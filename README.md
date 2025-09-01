# SumarioDePiezas
extrae los datos de un excel definido y hace un resumen de las piezas vendidas

## Descripción
SumarioDePiezas es una aplicación de escritorio desarrollada en Python y PyQt6 que permite extraer y resumir piezas vendidas desde archivos Excel (.xlsx) con un formato definido.

## Instalación
1. **Clona el repositorio:**
	```bash
	git clone https://github.com/josernesto1989/SumarioDePiezas.git
	cd SumarioDePiezas
	```

2. **Instala Python 3.10+** (si no lo tienes ya instalado).

3. **Instala las dependencias:**
	Se recomienda usar un entorno virtual:
	```bash
	python3 -m venv venv
	source venv/bin/activate
	pip install -r requirements.txt
	```

## Uso
1. Ejecuta la aplicación:
	```bash
	python main.py
	```

2. Arrastra y suelta tus archivos Excel (.xlsx) en la ventana principal o usa el botón para cargarlos.

3. La aplicación procesará los archivos y mostrará un resumen de las piezas vendidas.

## Requisitos
- Python >= 3.10
- PyQt6
- openpyxl

## Estructura principal
- `main.py`: Punto de entrada de la aplicación.
- `main_window.py`: Interfaz gráfica principal.
- `main_view_model.py`: Lógica de la interfaz y conexión con el procesador de Excel.
- `excel_processor.py`: Procesamiento de archivos Excel.
- `main_window.ui`: (Opcional) Archivo de diseño de la interfaz generado por Qt Designer.

## Licencia
Este proyecto está bajo la licencia MIT.

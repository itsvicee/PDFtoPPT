PDF to PPTX Converter
Una aplicación de escritorio sencilla e intuitiva para convertir archivos PDF a presentaciones de PowerPoint (.pptx) con un solo click.

🌟 Descripción
PDF to PPTX Converter es una herramienta desarrollada en Python que ofrece una solución rápida para transformar las páginas de un documento PDF en diapositivas individuales de una presentación de PowerPoint. La aplicación cuenta con una interfaz gráfica amigable construida con tkinter, que permite a cualquier usuario realizar la conversión sin necesidad de conocimientos técnicos.

El proceso se ejecuta en segundo plano para no congelar la aplicación, mostrando el progreso en tiempo real a través de una barra y mensajes de estado.

✨ Características Principales
Interfaz Gráfica Sencilla: Selecciona tu archivo PDF a través de un explorador de archivos nativo.

Conversión Directa: Convierte cada página del PDF en una diapositiva de PowerPoint a pantalla completa (formato 16:9).

Procesamiento Asíncrono: Gracias al uso de hilos (threading), la interfaz de usuario permanece fluida y receptiva durante la conversión.

Feedback en Tiempo Real: Una barra de progreso y una etiqueta de estado te mantienen informado sobre el proceso de conversión.

Gestión Automática de Archivos: El archivo .pptx resultante se guarda automáticamente en la misma carpeta que el PDF original.

Portátil y Ligero: No requiere instalaciones complejas, solo las dependencias de Python.

⚙️ Requisitos
Para poder ejecutar este script, necesitas tener instalado Python 3 y las siguientes bibliotecas:

PyMuPDF: Para leer y procesar los archivos PDF.

python-pptx: Para crear y manipular las presentaciones de PowerPoint.

Tkinter: (Generalmente incluido en las instalaciones estándar de Python).

🚀 Instalación
Clona o descarga este repositorio.

Instala las dependencias necesarias.
Abre tu terminal o línea de comandos y ejecuta el siguiente comando para instalar las bibliotecas requeridas a través de pip:

pip install PyMuPDF python-pptx

O si tienes problemas con el PATH de Windows, usa:

py -m pip install PyMuPDF python-pptx

▶️ ¿Cómo Usarlo?
Asegúrate de haber instalado todas las dependencias.

Ejecuta el script de Python desde tu terminal:

python pdf_to_ppt.py

O si es necesario:

py pdf_to_ppt.py

Se abrirá la ventana de la aplicación.

Haz clic en el botón "Seleccionar archivo PDF".

Elige el documento PDF que deseas convertir.

¡Listo! La conversión comenzará automáticamente. Podrás ver el progreso en la barra y el estado actual del proceso.

Una vez finalizado, encontrarás un archivo .pptx con el mismo nombre que tu PDF en la misma carpeta.

📜 Licencia
Este proyecto está bajo la Licencia MIT. Consulta el archivo LICENSE para más detalles.

# System Info Reporter

Herramienta de escritorio para Windows que recopila información detallada del sistema (hardware y software) y genera reportes en formatos HTML, PDF y Excel.

## Características

- **Interfaz Gráfica Sencilla**: Permite seleccionar fácilmente qué información incluir en el reporte.
- **Recolección de Datos Asíncrona**: La interfaz no se congela mientras se recopila la información.
- **Múltiples Formatos de Salida**: Genera reportes limpios y legibles en:
    - HTML
    - PDF
    - Excel (.xlsx)
- **Información Recopilada**:
    - Información general del sistema (Fabricante, Modelo).
    - Red (Hostname, IP, Dominio).
    - BIOS (Fabricante, Versión, Número de serie).
    - CPU (Modelo, Generación, Núcleos).
    - Memoria RAM (Capacidad, Velocidad).
    - Discos de Almacenamiento (Fabricante, Capacidad, Número de Serie).
    - Sistema Operativo (Nombre, Versión, Arquitectura).
    - Impresoras instaladas.
    - Lista de software instalado (puede ser un proceso lento).

## Estructura del Proyecto

El proyecto está dividido en tres componentes principales para separar la lógica de la presentación:

- **`InfoSystem_GUI.py`**:
    - Es el punto de entrada de la aplicación.
    - Crea la interfaz de usuario utilizando la librería `tkinter`.
    - Gestiona las selecciones del usuario y lanza el proceso de recolección de datos en un hilo secundario para mantener la fluidez de la aplicación.

- **`InfoSystem_backend.py`**:
    - Contiene toda la lógica para obtener la información del sistema.
    - Utiliza la librería `wmi` para interactuar con la API de Windows Management Instrumentation y extraer los datos de hardware y software.

- **`report_generator.py`**:
    - Se encarga de tomar los datos recopilados por el backend.
    - Utiliza las librerías `reportlab` para crear los documentos PDF y `openpyxl` para los archivos de Excel.
    - Genera los archivos de reporte con un nombre estandarizado que incluye el hostname y la fecha/hora.

## Cómo Funciona

1.  El usuario ejecuta la aplicación (`InfoSystem_GUI.py`).
2.  Se presenta una ventana donde puede marcar/desmarcar las categorías de información que desea obtener y el formato de salida.
3.  Al hacer clic en "Generar Reporte", la GUI llama a las funciones del `InfoSystem_backend.py` para recolectar los datos en segundo plano.
4.  Una vez recopilados los datos, se pasan al `report_generator.py`.
5.  El generador crea los archivos correspondientes (HTML, PDF, Excel) en la misma carpeta donde se encuentra el ejecutable.
6.  La aplicación muestra un mensaje de éxito indicando los nombres de los archivos generados.

## Cómo Crear el Ejecutable (.exe)

Para compilar la aplicación en un único archivo `.exe` auto-contenido, se utiliza **PyInstaller**.

### Prerrequisitos

- Python instalado.
- Instalar las dependencias listadas en `requirements.txt`. Abre una terminal y ejecuta:
  ```bash
  pip install -r requirements.txt
  ```
- El ícono `SysInf.ico` debe estar en la misma carpeta que los scripts.

### Compilación

La forma más sencilla de compilar es utilizando el archivo de especificaciones (`.spec`) que ya está configurado.

1.  Abre una terminal (cmd o PowerShell) en la carpeta raíz del proyecto.
2.  Ejecuta el siguiente comando:

    ```bash
    pyinstaller "SystemInfo.spec"
    ```

3.  PyInstaller creará las carpetas `build` y `dist`. Dentro de `dist`, encontrarás el archivo `SystemInfo.exe` listo para ser utilizado.

Alternativamente, si quieres generar el archivo `.spec` desde cero, puedes usar un comando como este:

```bash
pyinstaller --name "SystemInfo" --onefile --windowed --icon="SysInf.ico" InfoSystem_GUI.py
```

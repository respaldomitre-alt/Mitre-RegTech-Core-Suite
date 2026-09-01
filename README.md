# Mitre RegTech Core Suite v2026 🏎️🦉

Sistema Core de cumplimiento automatizado y minería de datos forenses diseñado para la prevención de lavado de dinero (PLD) en el sector Fintech e Instituciones de Financiamiento Tecnológico (IFPE) en México (2026).

## 🚀 Características Clave
- **Menú de Avenidas Maestro:** Separación de flujos para auditorías automatizadas o minería libre.
- **5 Algoritmos de Tipologías CNBV:** Detección en tiempo real de Pitufeo (T1), Picos Históricos (T2), Incongruencia de Objeto Social SAT (T3), Cuentas Puente Internacionales (T4) y Umbrales de Ley ROR (T5).
- **Firmas Geométricas Activas:** Gráficas vivas generadas en Matplotlib para el análisis espectral de transacciones.
- **Estación de Minería de Data Cruda:** Rejilla interactiva estilo Excel equipada con 5 filtros cruzados de Pandas en caliente (Monto, PyME ID, Giro SAT, Tipo SPEI y País).
- **Escribano de Gobierno Corporativo:** Generación automatizada de reportes institucionales en Word/PDF y timbrado automático de bitácoras anuales en Excel.

## 🛠️ ¿Cómo probar el sistema en tu Mac/PC?

Para correr el Ferrari Mitre en tu computadora local, sigue estos sencillos pasos:

1. **Clona o descarga este repositorio** en tu máquina.
2. Asegúrate de tener instalado **Python 3.10 o superior** y las librerías necesarias ejecutando en tu terminal:
   ```bash
   pip install pandas openpyxl python-docx fpdf2 matplotlib
   ```
3. Ejecuta el software central:
   ```bash
   python app_core_felipe.py
   ```
4. **Prueba de Campo:** En la pantalla de login, ingresa tus credenciales de auditor (la fecha se actualizará sola con el reloj de tu sistema). Al avanzar, el sistema te pedirá un archivo CSV; selecciona el archivo de muestra ubicado en la carpeta `Data_Prueba/muestrario_transacciones.csv` y ¡listo! Disfruta de la suite completa operando al 100%.

# 📦 Sistema de Inventario Físico - UPC San Francisco

Esta es una aplicación web progresiva (PWA) diseñada para realizar tomas de inventario físico de manera ágil y profesional desde cualquier dispositivo móvil o computadora. Permite la carga de datos desde Excel, el conteo en tiempo real con cálculos automáticos de diferencias y la exportación de reportes detallados en formatos PDF y Excel.

## 🚀 Características Principales

* **📱 Mobile First:** Optimizado para smartphones con teclado numérico automático.
* **📊 Cálculos en Tiempo Real:** Visualización instantánea de diferencias de cantidad y valor monetario.
* **🎨 Alertas Visuales:** Semáforo de colores (Rojo/Verde) para identificar sobrantes y faltantes.
* **📄 Reportes Profesionales:**
    * **PDF:** Formato contable con totales consolidados y alineación profesional.
    * **Excel (.xlsx):** Reporte con formato de moneda chilena y colores automáticos.
* **💾 Persistencia de Datos:** Los datos se guardan localmente en el navegador para evitar pérdidas por cierres accidentales o falta de internet.
* **📅 Formato Local:** Fechas en formato chileno (`DD - MM - AAAA`) y moneda `CLP`.

## 🛠️ Tecnologías Utilizadas

* **HTML5 / CSS3:** Estructura y diseño responsivo.
* **JavaScript (Vanilla):** Lógica del sistema y cálculos.
* **[SheetJS (XLSX)](https://sheetjs.com/):** Para la lectura y generación de archivos Excel.
* **[jsPDF](https://parall.ax/products/jspdf):** Para la creación de reportes en PDF.
* **[jsPDF-AutoTable](https://github.com/simonbengtsson/jspdf-autotable):** Para el diseño de tablas en los documentos PDF.

## 📋 Requisitos del Archivo Excel (Carga)

Para que el sistema procese la información, el archivo Excel de entrada debe tener el siguiente orden de columnas en la primera hoja:

1.  **A:** Código / ID del Producto
2.  **B:** Nombre del Producto
3.  **C:** Lote
4.  **D:** Sublote
5.  **E:** Unidad de Medida (UN)
6.  **F:** Cantidad Teórica (Sistema)
7.  **G:** Precio Unitario (Costo)

## 💻 Instalación y Uso

1.  Clona este repositorio o descarga los archivos.
2.  Abre el archivo `index.html` en cualquier navegador moderno.
3.  Ingresa los datos de la cabecera (Bodega, Site, Responsable).
4.  Carga tu archivo Excel maestro.
5.  ¡Comienza a contar! Los cambios se guardan automáticamente.
6.  Al finalizar, utiliza los botones de exportación para obtener tus reportes de auditoría.

## 📲 Configuración como PWA (App Móvil)

Para usarlo como una aplicación en el celular:
1.  Sube los archivos a un servidor HTTPS (GitHub Pages, Netlify).
2.  Abre el link en tu celular.
3.  Selecciona **"Agregar a la pantalla de inicio"** en las opciones de tu navegador.

---
Desarrollado para la toma de inventarios en **Casinos**.
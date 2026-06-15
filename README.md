# Sistema de Control de Resguardos y Activos — Hospital Escandón

Este sistema es una aplicación web local desarrollada para el **Área de Sistemas del Hospital Escandón**, diseñada para gestionar de manera integral el control de activos fijos (TI), direccionamiento de red, cuentas institucionales del personal, préstamos temporales y reportes de hojas de servicio de proveedores.

---

## 🚀 Módulos del Sistema

### 1. 📊 Panel Principal (Dashboard)
* Estadísticas clave en tiempo real (Personal activo, préstamos, IPs ocupadas/libres, impresoras y nodos).
* Detección de conflictos de red (direcciones IP duplicadas).
* Plan de mejora tecnológica (alertas de obsolescencia en equipos con RAM ≤ 4GB o discos duros HDD).
* Últimos movimientos y stock libre en almacén.

### 2. 👥 Personal Activo (Expediente Digital)
* Registro de trabajadores con datos de alta, dirección, área y ubicación física.
* Control de cuentas institucionales (`VRTCL`, `Correo`, `SAP B1`, `POS/Caja`, `Portal Requisiciones`) con almacenamiento de contraseñas opcionales.
* Generación automatizada de cartas responsivas de resguardo en formato Word (.docx).
* Carga y digitalización de resguardos firmados en formatos PDF o imágenes.

### 3. ⏱ Préstamos Temporales
* Control de equipos asignados de forma temporal (residentes, guardias nocturnas, etc.).
* Impresión de carta responsiva de préstamo temporal.
* Historial y marcación rápida de devolución con retorno automático de stock al almacén.

### 4. 🚪 Historial de Bajas
* Bitácora de empleados dados de baja en el sistema.
* Motivos de la baja y fecha de salida.
* Devolución automática de los activos asignados al inventario de almacén.
* Respaldo histórico de resguardos y cartas de devolución digitalizadas.

### 5. 🖥 Inventario General
* Catálogo de todos los activos fijos de TI (Computadoras, Laptops, Monitores, No-Breaks, etc.).
* Registro de especificaciones de hardware (procesador, RAM, almacenamiento) y licencias (Windows, Office).
* Control preciso de existencias (stock total y stock libre en almacén).

### 6. 🖨 Catálogo de Impresoras Kyocera Autorizadas
* Directorio restringido y ordenado para las **21 impresoras Kyocera autorizadas** por el Área de Sistemas de acuerdo con el número de serie, modelo y ubicación asignada.
* Gestión de permisos y computadoras conectadas a cada impresora.

### 7. 🛠 Hojas de Servicio y Mantenimiento (Nuevo)
* Módulo especializado para archivar y organizar reportes técnicos en formato PDF (entregados por proveedores de Kyocera).
* Clasificación por tipo de servicio: **Mantenimiento** o **Servicio de revisión/reparación**.
* Permite renombrar archivos para evitar nombres genéricos de escáner.
* Filtrado interactivo instantáneo por **Impresora**, **Fecha** y búsqueda por texto.

### 8. 🌐 Gestión de IPs
* Visualización en cuadrícula de direccionamiento lógico para subredes:
  * **Datos**: `192.168.254.x`
  * **Voz / Telefonía**: `172.16.90.x`
* Identificación por color del estado de cada dirección IP (libre, asignada, reservada, duplicada).

### 9. 📥 Centro de Reportes
* Descarga global o modular de la base de datos en formato Excel (.xlsx) con pestañas auto-ajustadas y diseño legible.

---

## 🛠 Tecnologías Utilizadas

* **Backend**:
  * Python 3.x
  * Flask (Servidor web ligero)
  * SQLite3 (Base de datos relacional integrada)
  * python-docx (Generación de cartas responsivas en Word)
  * openpyxl (Generación de hojas de cálculo de Excel)
* **Frontend**:
  * HTML5 semántico
  * CSS3 (Diseño responsivo, CSS Variables, paleta de colores institucional, animaciones de transición)
  * JavaScript ES6+ (AJAX con Fetch API, manejo asíncrono y renderizado dinámico en tiempo real)
  * JsBarcode (Generación de códigos de barras)

---

## ⚙️ Requisitos e Instalación

### Prerrequisitos
1. Tener instalado **Python 3** en el sistema.

### Instalación rápida (Windows)
1. Descarga o clona el repositorio en tu carpeta local.
2. Da doble clic sobre el archivo **`iniciar.bat`**.
   * El script configurará automáticamente las reglas del firewall (puerto 5000), instalará las dependencias descritas en `requirements.txt` e iniciará la aplicación backend.
3. El sistema abrirá automáticamente el navegador en:
   * **Local**: `http://localhost:5000`
   * **Red Local**: `http://192.168.254.150:5000` *(Comparte esta dirección con tu equipo de trabajo)*

---

## 📁 Estructura del Repositorio (Preparado para Git/GitHub)

El repositorio incluye un archivo **`.gitignore`** configurado para evitar subir archivos locales pesados o sensibles:
* **`database.db`**: Se ignora la base de datos local para evitar sobrescribir registros entre desarrolladores (el esquema se inicializa solo al correr por primera vez).
* **`archivos/`**: Se ignoran las cartas responsivas, hojas de servicio e imágenes subidas en producción, manteniendo únicamente el archivo `.gitkeep` para conservar la estructura del directorio.
* **`__pycache__/`**: Se omiten los archivos compilados de Python.

---
Hospital Escandón — Área de Sistemas.

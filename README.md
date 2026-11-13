# 🎓 Johana Software

<p align="center">
 <img src="CalificacionesDesktop/screenshots/icono-johana-software.png" alt="Johana Software" width="150">
</p>

<p align="center">
  <strong>Sistema de gestión y transferencia de calificaciones académicas</strong><br>
  Una herramienta de escritorio para organizar el trabajo con archivos Excel y facilitar la actualización de calificaciones.
</p>

<p align="center">
  <img src="https://img.shields.io/badge/Plataforma-Windows%2010%20%7C%2011-2563EB?logo=windows&logoColor=white" alt="Windows">
  <img src="https://img.shields.io/badge/.NET-8-512BD4?logo=dotnet&logoColor=white" alt=".NET 8">
  <img src="https://img.shields.io/badge/Interfaz-WPF-0A66C2" alt="WPF">
  <img src="https://img.shields.io/badge/Archivos-Excel-217346?logo=microsoftexcel&logoColor=white" alt="Excel">
</p>

---

## ✨ ¿Qué es Johana Software?

**Johana Software** es una aplicación de escritorio desarrollada para apoyar la gestión de calificaciones académicas mediante archivos de Microsoft Excel. Reúne dos procesos en una interfaz sencilla: la gestión de notas de proyectos y la transferencia de notas generales por asignatura y trimestre.

## 🖥️ Vista general

<p align="center">
  <img src="CalificacionesDesktop/screenshots/menu-principal.png" alt="Menú principal de Johana Software" width="85%">
</p>

Desde el menú principal puedes ingresar al módulo que necesitas:

- **Notas de proyectos:** carga de varias materias, revisión de diferencias y actualización del Excel maestro.
- **Notas generales:** transferencia de calificaciones trimestrales por asignatura.

## 📚 Módulos principales

### 1. Notas de proyectos

<p align="center">
<img src="CalificacionesDesktop/screenshots/modulo-notas-proyectos.png" alt="Módulo de notas de proyectos" width="90%"></p>

Permite seleccionar un Excel maestro y agregar varios archivos de materias. La pantalla reúne los archivos cargados y los resultados del análisis para revisar la información antes de aplicar las notas.

### 2. Notas generales

<p align="center">
  <img src="CalificacionesDesktop/screenshots/modulo-notas-generales.png" alt="Módulo de notas generales" width="90%">
</p>

Permite seleccionar el Excel maestro, elegir el trimestre y agregar uno o varios archivos de los docentes. El módulo está orientado a transferir el **TOTAL TRIMESTRE (100%)** por asignatura.

### 🔎 Revisión de calificaciones

<p align="center">
  <img src="CalificacionesDesktop/screenshots/analisis-calificaciones.png" alt="Análisis de calificaciones" width="90%">
</p>

Antes de aplicar los cambios, puedes analizar los archivos y revisar las coincidencias, las notas detectadas y los resultados del proceso.

## 🚀 Características

- Interfaz de escritorio con menú principal y dos módulos.
- Carga de uno o varios archivos Excel, según el módulo.
- Selección del archivo Excel maestro.
- Selección de trimestre en el módulo de notas generales.
- Análisis y revisión de calificaciones antes de aplicarlas.
- Actualización del archivo maestro mediante automatización de Microsoft Excel, para conservar su estructura y formato.

## 🧰 Tecnologías

- **C#**
- **.NET 8**
- **WPF (Windows Presentation Foundation)**
- **Microsoft Excel Interop / automatización de Excel**
- **Inno Setup** para generar el instalador de Windows.

## 💻 Requisitos

- Windows 10 o Windows 11 de 64 bits.
- Microsoft Excel de escritorio instalado.
- Para ejecutar la versión publicada autocontenida, no es necesario instalar .NET por separado.

> **Importante:** la aplicación utiliza Microsoft Excel de escritorio. Excel Online o aplicaciones alternativas como LibreOffice no sustituyen este requisito.

## 📦 Instalación

1. Descarga el instalador de Johana Software desde la sección **Releases** del repositorio, si hay una versión publicada.
2. Ejecuta `Johana_Software_Setup.exe`.
3. Sigue los pasos del asistente de instalación.
4. Abre Johana Software desde el acceso directo del escritorio.

Si todavía no existe una versión en **Releases**, solicita el instalador al responsable del proyecto.

## 🛠️ Ejecutar desde el código fuente

Requisitos para desarrollo:

- Visual Studio 2022.
- SDK de .NET 8.
- Microsoft Excel de escritorio.

Pasos generales:

1. Clona este repositorio.
2. Abre el archivo `.csproj` en Visual Studio.
3. Restaura los paquetes NuGet.
4. Compila y ejecuta el proyecto.

## 🔐 Uso responsable de la información

Los archivos Excel pueden contener información académica personal. Por seguridad:

- No publiques archivos reales de estudiantes en este repositorio.
- No incluyas calificaciones, nombres, cédulas ni datos institucionales en capturas o ejemplos públicos.
- Trabaja con copias de los archivos originales y conserva un respaldo del Excel maestro antes de realizar cambios.

## 📌 Estado del proyecto

Proyecto de escritorio para Windows. Las funciones disponibles corresponden a los módulos incluidos en el código fuente de este repositorio.

---

<p align="center">
  Hecho para facilitar la gestión académica.<br>
  <strong>Johana Software</strong>
</p>

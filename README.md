# 🧾 Generador de Carpetas y Archivos RTF desde Excel con PowerShell

Este repositorio contiene un script de PowerShell que permite **leer un archivo Excel (.xlsx)** y, a partir de su contenido, **crear una estructura organizada de carpetas y ficheros `.rtf`**. Todo está **comentado paso a paso**, ideal para mantenimiento, aprendizaje y personalización.

---

## 📌 ¿Qué hace este script?

- Lee un fichero Excel con columnas específicas.
- Crea carpetas según el valor de la columna `CODARTICULO`.
- Crea ficheros `.rtf` con:
  - El nombre definido en la columna `NOMBRE`.
  - El contenido en formato RTF contenido en la columna `TEXTO`.
- Organiza el contenido automáticamente según la estructura definida.
- Usa el campo `SECUENCIA` para ordenar ficheros opcionalmente si hay varios textos por artículo.

---

## 📁 Estructura esperada del Excel

Tu archivo `AnexosTxt.xlsx` debe tener una hoja llamada, por ejemplo, `in`, con las siguientes columnas (en cualquier orden, pero con los nombres correctos):

| Columna       | Descripción                                                                 |
|---------------|------------------------------------------------------------------------------|
| `TIPO`        | (Opcional) Tipo de contenido. No se usa por defecto, pero está disponible.  |
| `CODARTICULO` | Código del artículo. **Crea una carpeta** con este nombre.                  |
| `SECUENCIA`   | Número que indica el orden si hay varios textos por artículo.               |
| `USUARIO`     | (Opcional) Nombre del usuario. No utilizado por defecto.                    |
| `FECHA`       | (Opcional) Fecha de creación. No utilizado por defecto.                     |
| `NOMBRE`      | **Nombre del archivo** que se generará (sin extensión).                     |
| `TEXTO`       | **Contenido del archivo**, en formato RTF.                                  |

---

## 🗂 Ejemplo de estructura generada

Dado este contenido en el Excel:

| CODARTICULO | SECUENCIA | NOMBRE        | TEXTO         |
|-------------|-----------|----------------|---------------|
| ABC123      | 1         | Introducción   | {\rtf1\ansi...} |
| ABC123      | 2         | Especificación | {\rtf1\ansi...} |
| XYZ999      |           | Manual         | {\rtf1\ansi...} |

Se creará la siguiente estructura de carpetas y archivos:

/(Ruta donde está el script)
│
├── ABC123/
│ ├── 01_Introducción.rtf
│ └── 02_Especificación.rtf
│
└── XYZ999/
└── Manual.rtf



> **Nota:** Si `SECUENCIA` está presente y no es 0, se antepone al nombre del fichero para mantener el orden.

---

## 🔧 Requisitos

- PowerShell 5.1 o superior.
- Módulo [`ImportExcel`](https://github.com/dfinke/ImportExcel).

### 📦 Instalación del módulo ImportExcel

Si no tienes el módulo instalado, puedes hacerlo fácilmente desde PowerShell:

```powershell
Install-Module -Name ImportExcel -Scope CurrentUser




---------------------------------------------------------------------------------------------------------------------------------------------------------


# 🧾 Folder and RTF File Generator from Excel Using PowerShell

This repository contains a PowerShell script that **reads an Excel (.xlsx) file** and, based on its contents, **generates a structured hierarchy of folders and `.rtf` files**. The entire script is **fully commented step by step**, making it ideal for learning, customization, and long-term maintenance.

---

## 📌 What does this script do?

- Reads an Excel file with specific columns.
- Creates folders based on the values in the `CODARTICULO` column.
- Creates `.rtf` files with:
  - The filename taken from the `NOMBRE` column.
  - The file content in RTF format from the `TEXTO` column.
- Organizes everything into a structured directory tree.
- Uses the `SECUENCIA` field to optionally sort multiple texts under the same article.

---

## 📁 Expected Excel Structure

Your `AnexosTxt.xlsx` file should include a worksheet named, for example, `in`, with the following columns (order doesn't matter, but the column names must be exact):

| Column        | Description                                                                 |
|---------------|------------------------------------------------------------------------------|
| `TIPO`        | (Optional) Type of content. Not used by default, but available for future use. |
| `CODARTICULO` | Article code. **Creates a folder** using this value.                        |
| `SECUENCIA`   | Number indicating order if multiple texts exist for the same article.       |
| `USUARIO`     | (Optional) User name. Not used by default.                                  |
| `FECHA`       | (Optional) Creation date. Not used by default.                              |
| `NOMBRE`      | **Name of the file** to be created (without extension).                     |
| `TEXTO`       | **Content of the file**, in RTF format.                                     |

---

## 🗂 Example of Generated Structure

Given the following Excel content:

| CODARTICULO | SECUENCIA | NOMBRE        | TEXTO         |
|-------------|-----------|----------------|---------------|
| ABC123      | 1         | Introduction   | {\rtf1\ansi...} |
| ABC123      | 2         | Specification  | {\rtf1\ansi...} |
| XYZ999      |           | Manual         | {\rtf1\ansi...} |

The script will generate the following folder and file structure:

/(Directory where the script is located)
│
├── ABC123/
│ ├── 01_Introduction.rtf
│ └── 02_Specification.rtf
│
└── XYZ999/
└── Manual.rtf



> **Note:** If `SECUENCIA` is present and not equal to 0, it will be prepended to the file name (e.g., `01_`, `02_`) to maintain order.

---

## 🔧 Requirements

- PowerShell 5.1 or newer
- [`ImportExcel`](https://github.com/dfinke/ImportExcel) module

### 📦 Installing the ImportExcel Module

If the module is not already installed, you can easily install it from PowerShell:

```powershell
Install-Module -Name ImportExcel -Scope CurrentUser


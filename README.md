# ConvertExcelToCSV

![PowerShell Version](https://img.shields.io/badge/PowerShell-5.1%2B-blue)

Script de **Powershell** que convierte los datos de una hoja de un archivo **Excel** (`.xlsx`) pasada como paramatro a **CSV** (`.csv`) utilizando -ComObject.
Originalmente esta pensada para usarlo en tareas automaticas, pero tambien puede ejecutarse de manera manual.

## 📌 Instalación

### Clonar el repositorio

Ejecuta el siguiente comando en tu terminal para clonar el repositorio:

```sh
git clone https://github.com/pablolgamarra/ConvertExcelToCSV.git
cd ConvertExcelToCSV
```

### Descarga manual

1. Descarga el script `ConvertExcelToCSV.ps1` desde el repositorio.
2. Guarda el archivo en una carpeta de tu elección.

## 🛠 Uso

Ejecuta el script desde PowerShell con los siguientes parámetros:

```powershell
.\ConvertExcelToCSV.ps1 <ruta_excel> <nombre_hoja> [ruta_csv]
```

### 📌 Parámetros

-   **`<ruta_excel>`** → Ruta del archivo Excel de origen.
-   **`<nombre_hoja>`** → Nombre de la hoja que se exportará.
-   **`[ruta_csv]`** _(opcional)_ → Ruta de salida para el archivo CSV. Si no se especifica, se usará la misma ruta que el Excel con extensión `.csv`.

### 📍 Ejemplo

```powershell
.\ConvertExcelToCSV.ps1 "C:\Archivos\datos.xlsx" "Hoja1" "C:\Archivos\salida.csv"
```

## ⚠ Configurar ExecutionPolicy

Si es la primera vez que ejecutas scripts en PowerShell, es posible que necesites cambiar la política de ejecución. Para permitir la ejecución del script, abre PowerShell como Administrador y ejecuta:

```powershell
Set-ExecutionPolicy Bypass -Scope Process -Force
```

Esto permitirá la ejecución del script solo en la sesión actual sin modificar la configuración del sistema.

## ⏲ Uso con el Programador de Tareas de Windows

Para programar la ejecución diaria:

1. **Abre** el Programador de Tareas (`taskschd.msc`).
2. **Crea una nueva tarea básica**.
3. **Configura la frecuencia** (Diaria, Semanal, etc.).
4. **Selecciona "Iniciar un programa"**.
5. En **Programa o script**, escribe:

    ```powershell
    powershell.exe
    ```

6. En **Agregar argumentos**, usa:

    ```powershell
    -ExecutionPolicy Bypass -File "C:\Ruta\ConvertExcelToCSV.ps1" "C:\Ruta\Archivo.xlsx" "Hoja1"
    ```

7. **Guarda y ejecuta la tarea** para verificar que funciona.

## 🛠 Solucion a Problemas

-   **El script no se ejecuta:** Asegúrate de ejecutar PowerShell como Administrador.
-   **Error con la hoja:** Verifica que el nombre de la hoja sea exacto (sin espacios extra).
-   **Excel no instalado:** El script requiere Microsoft Excel para funcionar.

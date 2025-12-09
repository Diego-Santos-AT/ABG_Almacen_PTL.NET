# ABG Almacén PTL - .NET

Este proyecto es una conversión fiel del proyecto original ABG Almacén PTL de Visual Basic 6 a VB.NET (.NET 8.0).

## Estructura del Proyecto

```
ABG_Almacen_PTL.NET/
├── ABG_Almacen_PTL.sln           # Archivo de solución
└── ABG_Almacen_PTL.NET/          # Proyecto principal
    ├── Modules/                   # Módulos convertidos de VB6
    │   ├── GDConstantes.vb        # Constantes de la aplicación
    │   ├── GDGlobal.vb            # Variables globales y tipos de datos
    │   ├── GDFunc01.vb            # Funciones generales (menús, modos)
    │   ├── GDFunc02.vb            # Funciones de relación de datos
    │   ├── GDFunc04.vb            # Funciones de utilidad (SSCC, fechas)
    │   └── Profile.vb             # Lectura/escritura de INI y registro
    ├── Classes/                   # Clases convertidas de VB6
    │   └── clGenericaRecordset.vb # Clase genérica de recordset
    ├── DataAccess/                # Capa de acceso a datos
    │   ├── EntornoDeDatos.vb      # Acceso a BD de gestión de almacén
    │   └── edConfig.vb            # Acceso a BD de configuración
    ├── frmMain.vb                 # Formulario principal (MDI)
    └── Program.vb                 # Punto de entrada de la aplicación
```

## Requisitos

- .NET 8.0 SDK o superior
- SQL Server (para la base de datos)
- Windows (para Windows Forms)

## Compilación

```bash
cd ABG_Almacen_PTL.NET
dotnet restore
dotnet build
```

## Configuración

La aplicación busca un archivo `ABG.INI` en el directorio de ejecución con la siguiente estructura:

```ini
[Conexion]
Servidor=localhost
BaseDatos=Config
Usuario=sa
Password=
Timeout=30

[Archivos]
DSN=ABG.dsn
DLL=
```

## Cambios principales respecto a VB6

1. **ADO a ADO.NET**: Se ha reemplazado el uso de ADO (ADODB.Recordset) por ADO.NET (SqlConnection, SqlCommand, DataTable).

2. **DataEnvironment a Clases**: Los DataEnvironments de VB6 se han convertido a clases de acceso a datos con métodos tipados.

3. **Registro y archivos INI**: Se mantiene compatibilidad con las APIs de Windows para leer archivos INI y el registro.

4. **Formularios**: Los formularios MDI se han convertido a Windows Forms de .NET.

5. **Tipos de datos**: Los tipos de datos de VB6 se han mapeado a sus equivalentes en .NET.

## Mapeo de tipos VB6 a VB.NET

| VB6 | VB.NET |
|-----|--------|
| Integer | Short (o Integer para compatibilidad) |
| Long | Integer |
| String | String |
| Variant | Object |
| Currency | Decimal |
| Date | Date |
| Boolean | Boolean |

## Autor

Informática ATOSA

## Licencia

Propietario - Todos los derechos reservados

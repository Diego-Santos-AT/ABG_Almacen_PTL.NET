# ABG_Almacen_PTL.NET

Sistema de Gestión de Almacén PTL (Pick-to-Light)

## Descripción

ABG Almacén PTL es un sistema de gestión de almacén que proporciona funcionalidades para ubicar, extraer, empaquetar y consultar artículos en el almacén mediante tecnología Pick-to-Light.

## Proyectos en la Solución

### 1. ABG_Almacen_PTL.NET (Aplicación Original)
- **Tecnología**: Windows Forms con VB.NET
- **Framework**: .NET 8.0 (net8.0-windows)
- **Descripción**: Versión original convertida de VB6 a VB.NET
- **Estado**: Funcional para Windows

### 2. ABG_Almacen_PTL.MAUI (Nueva Aplicación Multiplataforma)
- **Tecnología**: .NET MAUI con C#
- **Framework**: .NET 10.0
- **Plataformas Soportadas**: 
  - Windows (net10.0-windows10.0.19041.0)
  - Android (net10.0-android)
- **Descripción**: Versión moderna multiplataforma de la aplicación
- **Estado**: En desarrollo

## Características de la Aplicación MAUI

### Funcionalidades Implementadas
- ✅ Sistema de login con autenticación de usuario
- ✅ Gestión de configuración mediante MAUI Preferences (reemplaza archivos INI)
- ✅ Conexión a base de datos SQL Server
- ✅ Selección de empresa y puesto de trabajo
- ✅ Menú principal con acceso a módulos
- ✅ Arquitectura de datos compartida (GlobalData, Constants)
- ✅ Capa de acceso a datos (ConfigDataAccess)

### Módulos del Sistema (Implementados)
- ✅ Ubicar BAC - Ubicación de BACs en ubicaciones PTL con validación y cambio de estado
- ✅ Extraer BAC - Extracción de BACs de ubicaciones PTL con opciones de estado
- ✅ Empaquetar BAC - Empaquetado rápido de BACs a cajas con generación de SSCC
- ✅ Consulta PTL - Consulta de BACs, ubicaciones y cajas con detalles de artículos
- ✅ Repartir Artículo - Reparto de artículos entre BACs con control de cantidad

## Requisitos

### Para Desarrollo
- .NET 10 SDK
- Workload de .NET MAUI instalado
- Visual Studio 2022 (17.9+) o Visual Studio Code con extensiones de C#/.NET MAUI
- SQL Server (para base de datos de configuración y gestión de almacén)

### Para Ejecución
- Windows 10/11 (versión 19041 o superior) para la aplicación Windows
- Android 5.0 (API 21) o superior para la aplicación Android

## Instalación de Workloads de MAUI

```bash
# Instalar MAUI para Android
dotnet workload install android

# Instalar MAUI para Windows
dotnet workload install maui-windows
```

## Compilación

### Proyecto Windows Forms (VB.NET)
```bash
cd ABG_Almacen_PTL.NET/ABG_Almacen_PTL.NET/ABG_Almacen_PTL.NET
dotnet build
```

### Proyecto MAUI (C#)
```bash
cd ABG_Almacen_PTL.MAUI

# Para Android
dotnet build -f net10.0-android

# Para Windows
dotnet build -f net10.0-windows10.0.19041.0
```

## Configuración

La aplicación MAUI utiliza MAUI Preferences para almacenar la configuración en lugar de archivos INI. La configuración incluye:

### Conexión a Base de Datos
- **BDDServLocal**: Servidor local de SQL Server
- **BDDConfig**: Nombre de la base de datos de configuración
- **BDDTime**: Timeout de conexión (en segundos)

### Valores por Defecto
- **UsrDefault**: Usuario por defecto
- **EmpDefault**: Código de empresa por defecto
- **PueDefault**: Puesto de trabajo por defecto

## Arquitectura

### Módulos Compartidos (MAUI)
- **GlobalData.cs**: Variables globales y estructuras de datos
- **Constants.cs**: Constantes de la aplicación
- **ConfigurationHelper.cs**: Gestión de configuración con Preferences

### Capa de Acceso a Datos
- **ConfigDataAccess.cs**: Acceso a base de datos de configuración
- Métodos para gestión de usuarios, empresas, puestos y permisos

### Páginas MAUI
- **LoginPage**: Página de inicio de sesión
- **MainPage**: Página principal con menú de opciones

## Migración de VB.NET a C#

La conversión a MAUI incluye la migración de VB.NET a C#:
- ✅ Módulos globales convertidos
- ✅ Capa de datos actualizada a C#
- ✅ Uso de características modernas de C# (.NET 10)
- ✅ Reemplazo de archivos INI por MAUI Preferences
- ✅ Interfaz de usuario actualizada a XAML

## Base de Datos

La aplicación requiere acceso a dos bases de datos SQL Server:
1. **Base de datos de configuración** (Config): Contiene usuarios, empresas, puestos y permisos
2. **Base de datos de gestión de almacén**: Datos operativos del almacén

## Licencia

Propietario de ABG

## Historial

- **Original**: Aplicación VB6
- **v1.0**: Conversión a VB.NET Windows Forms (.NET 8.0)
- **v2.0**: Migración a .NET MAUI multiplataforma (.NET 10.0) - En desarrollo

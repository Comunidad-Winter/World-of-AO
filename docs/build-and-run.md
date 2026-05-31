# Instalación, Compilación y Ejecución

> **Atribución y Notas:**
> *Esta documentación fue generada por y para **Comunidad-Winter** con el objetivo de preservar recursos de Argentum Online.*

Dado que **World-of-AO** se originó en la época de desarrollo basada en **Microsoft Visual Basic 6.0**, compilar el entorno exige pasos específicos relacionados a las dependencias de COM/ActiveX de aquel entorno. 
*(Esta guía ha sido inferida directamente desde los archivos y extensiones disponibles en el código fuente).*

## 1. Requisitos Previos

1. **IDE (Entorno de Desarrollo)**: Instalar Microsoft Visual Basic 6.0 (Preferentemente Service Pack 6).
2. **Sistema Operativo**: Recomendado ejecutarse bajo sistemas Windows de 32-bits para máxima compatibilidad (ej. Windows XP / Windows 7), aunque es posible hacerlo correr en Windows 10/11 utilizando herramientas de compatibilidad o máquinas virtuales.

## 2. Registro de Librerías y Dependencias

Antes de abrir el código, es **crítico** que el sistema operativo tenga registradas las dependencias `.ocx` y `.dll`. De lo contrario, Visual Basic emitirá errores de componentes faltantes.

Ubicados en el directorio raíz del cliente (ej. `Cliente 7.7 1024x768`), abre una consola de símbolo de sistema con privilegios de Administrador y utiliza el comando `regsvr32` sobre estos archivos identificados:

```cmd
regsvr32 CSWSK32.OCX
regsvr32 MSCOMCT2.OCX
regsvr32 MSINET.OCX
regsvr32 MSWINSCK.OCX
regsvr32 Mscomctl.ocx
regsvr32 RICHTX32.OCX
regsvr32 comctl32.ocx
regsvr32 msdxm.ocx
```

Adicionalmente, comprueba que las DLL como `DX7VB.DLL` (DirectX 7 for Visual Basic) estén presentes en las carpetas `SysWOW64` o `System32` del sistema, o dentro de la misma ruta de compilación.

## 3. Compilación

### Compilar el Servidor
1. Ingresa a la carpeta `Servidor WOAO\`.
2. Abre el archivo de proyecto `Server.vbp`.
3. Ve a `File -> Make Server.exe` (Archivo -> Generar Server.exe).
4. El ejecutable del servidor quedará generado en la raíz del directorio del servidor.

### Compilar el Cliente
1. Ingresa a la carpeta `Cliente 7.7 1024x768\` o `Cliente 7.7 1280x720\`.
2. Abre el archivo de proyecto de Visual Basic `GMDinamico.vbp` (u otro archivo `.vbp` que posea la terminación principal).
3. Ve a `File -> Make Cliente.exe` (Archivo -> Generar Cliente.exe).

## 4. Configuración y Ejecución Local

### Servidor
Antes de encender el servidor, verifica el archivo `Server.ini` ubicado en la carpeta del servidor. Asegúrate de que las opciones de red estén configuradas para operar localmente:
```ini
[INIT]
ServerIp=127.0.0.1
StartPort=9000
```
Inicia `Server.exe`. Se habilitará la consola y comenzará la carga de mapas e items.

### Cliente
1. Inicia el ejecutable generado del Cliente.
2. Ingresa a la sección "Conectar" y si existe un menú de IPs, o bien modifícalo en el código antes de compilar para que apunte a `127.0.0.1` usando el puerto `9000` (el puerto predeterminado que figura en `Server.ini`).
3. Crea una cuenta/personaje e ingresa al juego.

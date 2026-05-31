# World-of-AO (Mod Aodrag 7.0)

> **Atribución y Notas:**
> *Esta documentación fue generada por y para **Comunidad-Winter** con el objetivo de preservar recursos de Argentum Online.*

Repositorio de preservación histórica del servidor y clientes de **World-of-AO**, un mod avanzado basado en la versión 7.0 de Aodrag (Argentum Online).

## Qué es

World-of-AO es un proyecto histórico basado en el motor de Argentum Online (específicamente la rama Aodrag 7.0). Implementa una gran cantidad de sistemas personalizados para la época, expandiendo mecánicas de juego orientadas al rol, clanes y combate competitivo.

## Estructura del repositorio

El repositorio está compuesto por los siguientes directorios principales:

- `Cliente 7.7 1024x768/`: Código fuente, recursos y dependencias del cliente gráfico adaptado para resolución 1024x768.
- `Cliente 7.7 1280x720/`: Código fuente, recursos y dependencias del cliente gráfico adaptado para resolución panorámica de 1280x720.
- `Servidor WOAO/`: Código fuente, datos, mapas y configuración del servidor del juego.
- `docs/`: Documentación técnica complementaria generada para entender la arquitectura y componentes del repositorio.

## Componentes principales

- **Servidor (`Server.vbp`)**: Escrito en Visual Basic 6. Gestiona la lógica principal del juego (Sistema de combate, NPCs, clanes, bovedas, facciones, monturas) y persiste los datos en formato de texto e `.ini`.
- **Cliente (`GMDinamico.vbp`)**: Escrito en Visual Basic 6. Renderiza el juego usando DirectX 7/8. Maneja la comunicación con el servidor mediante sockets TCP (`CSWSK32.OCX` / `MSWINSCK.OCX`). Se incluyen dos versiones de cliente que difieren en la resolución nativa de sus pantallas y formularios.

## Cómo funciona (Arquitectura Básica)

El sistema utiliza una arquitectura **Cliente-Servidor** clásica mediante el protocolo TCP/IP. El cliente se encarga de recolectar input del usuario, reproducir audio/música e implementar el renderizado gráfico mediante una librería externa (como `DX7VB.DLL` o DirectX nativo). El servidor autoritativo valida el movimiento, resuelve combates, calcula visibilidad e interactúa con el sistema de archivos para almacenar los personajes, clanes y mapas.
Más detalles en [architecture.md](docs/architecture.md).

## Instalación / Compilación / Ejecución

*Nota: Todos los pasos son inferidos basados en la tecnología del proyecto (VB6).*

1. **Requisitos**: Visual Basic 6.0 instalado.
2. **Registro de Dependencias**: Es necesario registrar (usando `regsvr32`) las librerías `.ocx` y `.dll` incluidas en las carpetas de los clientes (ej. `CSWSK32.OCX`, `MSCOMCT2.OCX`, `DX7VB.DLL`).
3. **Compilación**: Abrir `Server.vbp` (Servidor) o `GMDinamico.vbp` (Cliente) con VB6 y compilar a ejecutable (.exe).
4. **Ejecución del Servidor**: Ejecutar el Servidor verificando el archivo `Server.ini` (puerto por defecto 9000, `ServerIp=127.0.0.1`).

Más detalles en [build-and-run.md](docs/build-and-run.md).

## Estado del proyecto

Este repositorio se cataloga bajo el estado de **Preservación Histórica**. Contiene código heredado de Visual Basic 6 y su funcionamiento actual en entornos modernos (Windows 10/11) requiere configuraciones de compatibilidad o máquinas virtuales. No es un proyecto activo bajo desarrollo normal.

## Notas

- **Archivos inferidos**: Hay muchas lógicas en código ("hardcodeadas") debido al estilo de programación en VB6 de la época.
- El proyecto parece estar fuertemente modificado desde su base original (Aodrag), integrando sistemas como monturas, castillos, remort y guardianes.
- Puede haber código en desuso o implementaciones incompletas. Las contradicciones entre nombres de archivos e implementaciones fueron resueltas confiando en el código fuente de los archivos `.vbp` y `.bas`.

## Implementaciones Originales del Mod

- Sistema de monturas.
- Sistema de Castillos.
- Sistema de facciones.
- Sistema de remort.
- Sistema de eventos automáticos.
- Sistema de clanes (Puntos por conquistas, muertes, etc).
- Sistema de Quest y Canjes por puntos quest.
- Sistema de Guardianes (mejoran stats).
- Sistemas de equipos de canjes por clases.
- Sistema de amuleto de salvación (se quema al morir para evitar pérdida de items).
- Motor gráfico mejorado y Pantalla ampliada (versión 1280x720).
- Contadores (invisibilidad, paralizar, muertes).
- Sistema de party e hijos.
- Clases nuevas y Regalo de los dioses.
- Sistema de gemas y retos 1v1 / 2v2 por oro.

## Imágenes

<img width="1022" height="714" alt="image" src="https://github.com/user-attachments/assets/5242bb30-8f78-48a8-9e4c-45cec2b26384" />

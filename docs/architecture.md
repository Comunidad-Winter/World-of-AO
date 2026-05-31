# Arquitectura del Sistema

> **Atribución y Notas:**
> *Esta documentación fue generada por y para **Comunidad-Winter** con el objetivo de preservar recursos de Argentum Online.*

La arquitectura de **World-of-AO** se adhiere al modelo clásico de Argentum Online, estructurado rígidamente sobre el paradigma **Cliente-Servidor Autoritativo** programado en **Visual Basic 6**.

## Modelo Cliente-Servidor

El flujo principal de información opera de la siguiente manera:
1. **Cliente ("Dumb Client")**: El cliente actúa principalmente como un terminal de renderizado gráfico y colector de inputs (teclado/ratón). Envía comandos serializados al servidor.
2. **Servidor (Autoritativo)**: Todas las lógicas de negocio, validaciones de combate, cálculos de ruta de NPCs (AI_NPC), rangos de visión y persistencia de cuentas se ejecutan y validan en este nodo. El servidor responde al cliente con los cambios de estado (ej: actualizar vida, reproducir un sonido, mover un personaje en el grid de la pantalla).

## Red y Protocolo (Capa de Sockets)

La comunicación se basa en sockets TCP bidireccionales de manera asíncrona.
- **Implementación (Cliente)**: Usa el control OCX `CSWSK32.OCX` (Catalyst Socket) o la API de Windows Sockets estándar, manejado en el módulo de red (inferido de la presencia de `WSKSOCK.bas` y `TCP.bas`).
- **Implementación (Servidor)**: El servidor gestiona múltiples conexiones simultáneas mapeándolas a un array global de perfiles de usuario (`UserIndex`). Dispone de diferentes iteraciones o backups de lógica de red (`TCP.bas`, `TCP1.bas`, `TCP2.bas`, `TCP3.bas`). El protocolo encripta paquetes específicos, lo cual se infiere por la existencia de `BlowFish.bas` y módulos criptográficos.

## Renderizado y Gráficos (Cliente)

El pipeline de renderizado del cliente no emplea la interfaz estándar de controles de Visual Basic, sino que dibuja directamente en búferes de video para conseguir mayor rendimiento:
- **DirectX**: Se utiliza DirectX 7/8 (evidenciado por librerías como `DX7VB.DLL` y módulos de `DX_InIt.bas` y `TileEngine.bas`).
- **Tile Engine**: Los mapas son representaciones en un grid bidimensional (grillas). El cliente lee estructuras binarias `.map` y renderiza capas de terrenos, decoraciones, personajes y techos superpuestos.

## Sistema de Persistencia

- **Archivos de Texto / INI**: El servidor de World-of-AO, como muchos proyectos derivados de versiones anteriores a la 12.0 de AO, persiste la base de datos principal en formato `.ini` y `.txt` en disco. Las carpetas `Charfile/`, `Guilds/`, `Accounts/` almacenan la información de estado.
- **MySQL (Probable/Incompleto)**: Se detectó la presencia de un archivo `MySQLDTA.dll` y un módulo `Mysql.bas` en el código del servidor, lo que sugiere que hubo integraciones con bases de datos MySQL, posiblemente para el portal web o registros específicos, aunque la base del funcionamiento del juego sigue recayendo fuertemente en sistemas de archivos locales.

## Motor de IA (Servidor)

La lógica de los enemigos (NPCs) está centralizada. Los scripts como `MODULO_NPCs.bas` y `AI_NPC.bas` iteran sobre un bucle de juego (GameLoop) decidiendo acciones como movimiento y ataques de los NPCs basados en el rango de proximidad y agresividad preconfigurada en los archivos `.dat` del servidor.

# Notas Técnicas y Estado de Preservación

> **Atribución y Notas:**
> *Esta documentación fue generada por y para **Comunidad-Winter** con el objetivo de preservar recursos de Argentum Online.*

Este archivo consolida advertencias, deuda técnica, y suposiciones formuladas a partir del análisis estático del repositorio de **World-of-AO**.

## 1. Contexto de Preservación
Este repositorio **no es un proyecto de software activo**. Es una pieza histórica resguardada para su estudio y conservación. Las métricas convencionales de calidad de software moderno no aplican a esta base de código (Clean Code, Patrones de Diseño modernos, Tests Unitarios no existen en este entorno). 

## 2. Suposiciones e Inferencias
Debido a la carencia de documentación original estricta, ciertos comportamientos han sido **inferidos** mediante la estructura y nomenclatura de archivos:
- **Herencia**: El archivo original `README.md` indica un origen basado en "Aodrag 7.0". Sin embargo, la vasta cantidad de formularios (como `frmMontura`, `frmFamilia`, `frmSubastas`) evidencia que sufrió bifurcaciones extremas desde su punto de partida base, con agregados y características hardcodeadas directamente.
- **Nombres de Proyectos Inconsistentes**: En la carpeta del cliente existen docenas de archivos `.vbp` duplicados (`GMDinamico - copia - copia...vbp`). Esto suele ser una mala práctica común de la época como método de "backup". Es seguro ignorarlos y centrarse en `GMDinamico.vbp`.
- **Integración con Bases de Datos (No verificada totalmente)**: Se visualizó `MySQLDTA.dll` y `Myqsl.bas` (sic). Sin embargo, gran parte del almacenamiento (`Charfile`, `Guilds`) sigue el modelo `INI` característico del motor base de AO. Se sospecha que MySQL se usaba secundariamente para portales web, pero no es excluyente para levantar el servidor localmente.

## 3. Deuda Técnica y Observaciones
- **Código Hardcodeado**: Muchos índices (ej: `ArmaduraImperial1 = 370` en el `Server.ini`) y validaciones están escritos directamente en el código fuente (`.bas`), acoplando los datos del servidor a la lógica del binario.
- **Seguridad**: El cliente incluye archivos nombrados `AntiCheat.bas`, sin embargo, debido a la naturaleza del protocolo en texto plano (en su mayoría) de VB6 en esa época, la seguridad de red de este proyecto no resistiría inyecciones ni ingenierías inversas contemporáneas.
- **Gestión de Memoria y API de Windows**: El cliente utiliza extensas llamadas a la API (API de `user32` y `kernel32`) descritas en `APIdeclaraciones.bas`. Estas llamadas podrían originar problemas de estabilidad al ejecutar el juego bajo sistemas de 64-bits que interceptan punteros largos, por lo que requiere especial atención o parches si se intenta modernizar su binario nativo.
- **Rendimiento Visual**: Se mantienen dependencias a librerías DirectX antiguas y controles Flash (`Flash10a.ocx`), las cuales están completamente obsoletas y son focos de vulnerabilidades en entornos modernos.

## 4. Estado de los Recursos
Los recursos binarios del repositorio (`GRAFICOS/`, `Midi/`, `wav/`, `Mapas/`) aparentan estar intactos, pero los formatos cerrados requerirían de las herramientas originales del juego (Indexadores de recursos, Editores de mapas originados en la era de ORE / AO 0.11.x) si se pretendiese realizar modificaciones masivas a los activos multimedia.

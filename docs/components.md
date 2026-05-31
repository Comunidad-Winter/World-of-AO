# Componentes del Proyecto

> **Atribución y Notas:**
> *Esta documentación fue generada por y para **Comunidad-Winter** con el objetivo de preservar recursos de Argentum Online.*

Este documento desglosa las piezas lógicas principales encontradas en la base del código de **World-of-AO**. Todos los módulos detectados pertenecen al ecosistema de Visual Basic 6.

## 1. Servidor (`Servidor WOAO`)

El directorio `Servidor WOAO` engloba el núcleo funcional y autoritativo.

### Módulos Base y Lógica de Juego (`Codigo/`)
- **GameLogic.bas**: Orquestador principal del ciclo de vida del juego, resolviendo validaciones generales.
- **SistemaCombate.bas**: Algoritmos de cálculo de daño, evasión, modificadores por clases y razas, y resoluciones de golpes físicos.
- **modHechizos.bas**: Manejo de daño mágico, sanaciones, alteraciones de estado (parálisis, invisibilidad) y área de efecto.
- **AI_NPC.bas / MODULO_NPCs.bas**: Máquina de estados básica y rutinas para la inteligencia artificial de las entidades no-jugadoras.

### Sistemas Característicos (Mods)
- **Monturas.bas**: Lógica que amplía el movimiento y atributos del usuario montado.
- **ModClanes.bas / ModBovedaClan.bas**: Administración de creación de clanes, facciones, bóvedas compartidas.
- **ConquistaCiudad.bas / Mod_Guerras.bas / Mod_Arena.bas**: Sistemas competitivos PvP (Player vs Player), arenas 1v1 y eventos automáticos de guerra.
- **Trabajo.bas / InvUsuario.bas**: Sistemas de recolección (talar, minar, pescar), forja y administración de inventarios de los jugadores.

### Datos del Servidor (`Dat/` y `.ini`)
Los datos estáticos se leen de binarios y configuraciones (items, balance, clases):
- Configuración global leída de `Server.ini`.
- La persistencia descansa sobre clases de utilería de acceso a disco como `clsIniReader.cls`.

---

## 2. Cliente (`Cliente 7.7 ...`)

Existen dos proyectos cliente independientes, adaptados para distintas resoluciones, pero que comparten una base de módulos homóloga. Se describirán basándose en el cliente principal.

### Interfaz de Usuario y Controles
- **frmMain.frm**: El formulario central donde se aloja la pantalla renderizada del juego, el chat en vivo, barras de estadísticas y controles de interfaz general (inventario, hechizos).
- **Sub-formularios (`frm*.frm`)**: Pantallas modales específicas del diseño UI. Se incluyen controles de Clanes (`frmGuildAdm.frm`), Banco (`frmBancoObj.frm`), Comercio (`frmComerciar.frm`), Skills (`frmSkills3.frm`), entre muchos otros orientados a brindar soporte visual a los nuevos sistemas.

### Motor Gráfico y Audio
- **TileEngine.bas / DX_InIt.bas**: Integración con las librerías DirectX de Windows. El TileEngine lee los `Mapas/` procesados y se encarga del renderizado de capas 2D.
- **Bass.bas / Mod_Mp3.bas / ClsAudio.cls**: Dependencias para reproducción asíncrona de efectos de sonido ambiente y banda sonora.

### Red y Criptografía
- **TCP.bas**: Envío y recepción de un flujo binario hacia el servidor. Mapea la interfaz con los servidores de Sockets (ej. CSWSK32).
- **BlowFish.bas / Mod_Cripto.bas**: Encripción simétrica básica orientada a oscurecer la información de las contraseñas e intercambios críticos de red, una práctica común y necesaria para evitar bots/exploits sencillos.

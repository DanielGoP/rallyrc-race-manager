# Rally RC Race Manager

Aplicación web para gestionar carreras de Rally RC con **Google Apps Script** y **Google Sheets**. Permite configurar carreras con uno o varios tramos, registrar tiempos desde móvil, tablet u ordenador, calcular la clasificación en tiempo real y publicar el resultado en un campeonato independiente.

No requiere servidor propio, base de datos externa ni framework frontend.

## Funcionalidad actual

- Acceso mediante PIN y registro del juez o responsable de mesa.
- Carreras dinámicas con uno o varios tramos.
- Generación automática de idas y vueltas para cada tramo.
- Estados de carrera `CONFIGURACION` e `INICIADA`.
- Registro de tiempos y penalizaciones por inscripción y pasada.
- Bloqueo de resultados duplicados.
- Clasificación dinámica, descarte global, estado y gap.
- Dashboard de tiempos y constancia.
- Administración responsive para carrera, resultados, inscripciones, pilotos y categorías.
- Corrección administrativa de resultados existentes.
- Backups automáticos durante la migración schema v5 y al crear una carrera.
- Publicación única en un visualizador de campeonato con histórico compatible.

El modelo activo no utiliza dorsal. Una inscripción representa la participación de un piloto en una categoría:

| inscripcionId | pilotoId | categoriaId | piloto | categoria |
|---|---|---|---|---|
| I001 | P001 | C001 | Piloto Demo 1 | Rally 1/10 |

Un piloto puede participar en varias categorías, pero no puede tener dos inscripciones en la misma categoría.

## Modelo de carrera

### Tramos y pasadas

La estructura se configura por tramos. La aplicación asigna automáticamente los nombres `Tramo 1`, `Tramo 2`, etc. y genera sus pasadas.

Cada tramo cumple estas reglas:

- tiene una o más idas;
- tiene cero vueltas o exactamente tantas vueltas como idas;
- con vueltas, las pasadas se generan como `Ida 1`, `Vuelta 1`, `Ida 2`, `Vuelta 2`, etc.;
- sin vueltas, se generan únicamente `Ida 1`, `Ida 2`, etc.;
- los identificadores de las pasadas existentes se conservan cuando se amplía una carrera.

Ejemplo:

```text
Tramo 1: 2 idas y 2 vueltas
  Ida 1, Vuelta 1, Ida 2, Vuelta 2

Tramo 2: 1 ida y 0 vueltas
  Ida 1
```

La configuración admite entre 1 y 100 tramos y entre 1 y 500 pasadas totales.

### Ciclo de vida

Una carrera tiene dos estados:

| Estado | Comportamiento |
|---|---|
| `CONFIGURACION` | Permite editar tramos e inscripciones. No admite resultados. |
| `INICIADA` | Admite resultados. La estructura existente queda bloqueada. |

Iniciar la carrera es una transición explícita de `CONFIGURACION` a `INICIADA`. Después de iniciarla no se pueden quitar tramos ni pasadas, reducir cantidades o cambiar una modalidad. Solo se puede ampliar un tramo existente:

- un tramo con idas y vueltas aumenta por parejas `ida + vuelta`;
- un tramo solo de idas aumenta de una ida en una ida.

Crear una nueva carrera genera un `CARRERA_ID`, deja la prueba en `CONFIGURACION`, conserva las inscripciones seleccionadas, limpia los resultados y crea backups antes de sustituir la carrera activa.

Una carrera publicada en el campeonato queda bloqueada para modificaciones. `PUBLICANDO` impide cambios mientras dura el envío. Si la respuesta remota es ambigua, la carrera pasa a `PUBLICACION_INCIERTA` y permanece bloqueada hasta reintentar exactamente el mismo snapshot.

## Cálculo de clasificación

La clave lógica de un resultado activo es:

```text
carreraId + inscripcionId + pasadaId
```

La ruta normal rechaza duplicados; una corrección administrativa actualiza la fila existente.

Reglas de total y descarte:

- sin resultados, el total queda vacío y el estado es `Sin resultados`;
- mientras falten pasadas, el total es la suma parcial y el estado es `Pendiente`;
- al completar una carrera con más de una pasada, se descarta la peor pasada de toda la carrera, no una por tramo;
- una carrera de una sola pasada tiene cero descartes;
- el descarte se aplica al mayor `total` de pasada, que ya incluye su penalización;
- si dos pasadas empatan como peores, se descarta la de orden posterior;
- las penalizaciones mostradas suman todas las pasadas registradas, incluida la descartada;
- los estados incorporan `con penalización` cuando corresponde.

La clasificación se ordena por total ascendente. Los pilotos sin total quedan al final; los desempates usan categoría y piloto.

El gap se calcula contra la fila anterior únicamente cuando ambos resultados pertenecen a la misma categoría y tienen exactamente el mismo conjunto de pasadas completadas. En otro caso se muestra `-`.

## Hojas activas de cronometraje

`setupSheets()` crea y mantiene estas hojas activas:

| Hoja | Cabecera exacta | Uso |
|---|---|---|
| `Config` | `key \| value` | PIN, identidad, estado, revisión, publicación, entorno y versión de esquema. |
| `PilotosDB` | `pilotoId \| nombre \| alias \| activo \| notas \| createdAt \| updatedAt` | Catálogo estable de pilotos. |
| `CategoriasDB` | `categoriaId \| nombre \| activa \| orden \| notas \| createdAt \| updatedAt` | Catálogo estable de categorías. |
| `Tramos` | `tramoId \| nombre \| orden \| numIdas \| numVueltas` | Definición de tramos de la carrera activa. |
| `Pasadas` | `pasadaId \| tramoId \| label \| tipo \| numero \| orden` | Pasadas generadas para todos los tramos. |
| `Inscripciones` | `inscripcionId \| pilotoId \| categoriaId \| piloto \| categoria` | Participantes de la carrera activa, sin dorsal. |
| `Resultados` | `resultadoId \| timestamp \| carreraId \| inscripcionId \| tramoId \| pasadaId \| piloto \| categoria \| pasadaLabel \| tiempo \| penalizacion \| total \| juez` | Un tiempo por inscripción y pasada. |
| `Clasificacion` | `inscripcionId \| pilotoId \| categoriaId \| piloto \| categoria \| descartada \| penalizaciones \| total \| gap \| completadas \| previstas \| estado` | Resumen calculado; las pasadas se obtienen dinámicamente de `Pasadas` y `Resultados`. |

Valores principales de `Config`:

| key | Ejemplo o valor |
|---|---|
| `PIN` | `CAMBIAR_PIN_ACCESO` |
| `ADMIN_PIN` | `CAMBIAR_PIN_ADMIN` |
| `CARRERA_ACTIVA` | `Carrera Demo` |
| `CARRERA_ID` | UUID generado |
| `CARRERA_ESTADO` | `CONFIGURACION` o `INICIADA` |
| `CONFIG_REVISION` | entero incremental |
| `CAMPEONATO_PUBLICADA` | `NO`, `PUBLICANDO` o `SI` |
| `CAMPEONATO_PUBLICACION_INICIADA` | fecha ISO durante un envío |
| `ENVIRONMENT` | `PRODUCTION` o entorno de prueba |
| `SCHEMA_VERSION` | `5` |

No añadas dorsales ni columnas de pasada fija a las hojas activas. Los nombres `piloto` y `categoria` son copias para lectura; `pilotoId` y `categoriaId` son las identidades estables.

## Administración

Administración usa una interfaz responsive de sala de control:

- navegación lateral en escritorio y horizontal desplazable en pantallas estrechas;
- resumen de estado y publicación;
- configuración, inicio, ampliación y creación de carrera;
- carga y corrección de resultados;
- alta, edición y baja de inscripciones sin resultados;
- creación y edición de pilotos y categorías;
- bloqueo explícito de la sesión administrativa.

Las inscripciones con resultados no pueden editarse ni eliminarse. Tampoco se admiten pilotos o categorías inactivos ni un duplicado de `pilotoId + categoriaId`.

## Migración a schema v5 y backups

`setupSheets()` compara `Config.SCHEMA_VERSION` con la versión soportada. Si la hoja usa una versión futura, se detiene sin intentar degradarla. Si necesita migración:

1. adquiere un `ScriptLock`;
2. crea una única copia de `Config`, `Tramos`, `Pasadas`, `Inscripciones`, `Resultados` y `Clasificacion` con sufijo `_backup_v5_<fecha>`;
3. registra el marcador `MIGRATION_V5_BACKUP`;
4. convierte inscripciones y resultados antiguos al modelo normalizado;
5. crea una definición legacy equivalente a seis pasadas solo cuando los datos antiguos no tenían definición dinámica;
6. recalcula `Clasificacion` y escribe `SCHEMA_VERSION = 5`.

Los lectores de migración aceptan inscripciones históricas de cuatro o cinco columnas y resultados numéricos antiguos. Esa compatibilidad sirve únicamente para migrar o leer datos legacy; no es el contrato de escritura activo.

Al generar una nueva carrera se crean backups independientes con sufijo `_backup_nueva_carrera_<fecha>` de las mismas seis hojas de carrera. `PilotosDB` y `CategoriasDB` no se limpian porque son catálogos maestros.

## Instalación de cronometraje

1. Crea un Google Sheet y abre `Extensiones -> Apps Script`.
2. Copia `codigo.gs` en el archivo de servidor.
3. Crea un archivo HTML llamado exactamente `Index` y copia `index.html`.
4. Guarda y ejecuta `setupSheets()` desde el editor.
5. Revisa las ocho hojas activas y cambia `PIN` y `ADMIN_PIN` en `Config`.
6. Publica como aplicación web, ejecutando como propietario y con el acceso adecuado para el evento.

El archivo del repositorio se llama `index.html`, pero en Apps Script debe ser `Index.html` porque `doGet()` usa `createTemplateFromFile('Index')`.

## Campeonato schema v2

`campeonato-app` es una segunda Web App y debe usar otro Google Sheet. El esquema actual del campeonato es v2 y conserva tanto carreras dinámicas nuevas como el histórico v1.

Hojas del campeonato:

| Hoja | Uso |
|---|---|
| `Config` | `SCHEMA_VERSION = 2`, nombre y descartes del campeonato. |
| `Puntuacion` | Puntos por posición. |
| `Carreras` | Registro de carreras importadas; su fila actúa como marcador de importación completada. |
| `TramosCarrera` | Definición normalizada de tramos y pasadas para publicaciones v2. |
| `ClasificacionesCampeonato` | Resumen sin dorsal de cada inscripción publicada en v2. |
| `TiemposCampeonato` | Tiempos normalizados por `carreraId + inscripcionId + pasadaId`. |
| `ResultadosCampeonato` | Histórico legacy v1 de seis pasadas; puede contener dorsal y no recibe publicaciones v2. |

`setupChampionshipSheets()` añade las hojas normalizadas y eleva `SCHEMA_VERSION` de 1 a 2 sin convertir ni borrar `ResultadosCampeonato`. El visualizador combina mediante adaptadores de lectura:

- carreras v1 desde `ResultadosCampeonato`, con una definición legacy de seis pasadas;
- carreras v2 desde `TramosCarrera`, `ClasificacionesCampeonato` y `TiemposCampeonato`.

El dorsal legacy no forma parte del modelo público canónico ni de ninguna publicación v2.

### Configuración del campeonato

1. Crea otro Google Sheet y abre Apps Script.
2. Copia `campeonato-app/codigo.gs` y crea `Index.html` con `campeonato-app/index.html`.
3. Configura `CHAMPIONSHIP_TOKEN` en las propiedades del script del campeonato.
4. Ejecuta `setupChampionshipSheets()`.
5. Despliega la Web App del campeonato y conserva su URL `/exec`.
6. En las propiedades del script de crono configura `CHAMPIONSHIP_ENDPOINT` con esa URL y `CHAMPIONSHIP_TOKEN` con el mismo secreto.

No guardes el token en el repositorio, en el HTML ni en `Config`.

Cada `CARRERA_ID` se importa una sola vez. Las publicaciones v2 validan identificadores estables, definición, pasadas, penalizaciones, totales, descarte y estado antes de escribir. También incluyen un `payloadHash` SHA-256 que el receptor recalcula: un reintento solo devuelve `ALREADY_EXISTS` cuando el contenido coincide exactamente. La aplicación de crono no sincroniza de nuevo una carrera ya publicada.

## Orden de despliegue

Para actualizar una instalación existente al contrato dinámico:

1. Haz una copia externa de los dos Google Sheets.
2. Actualiza primero `campeonato-app/codigo.gs` e `Index.html` en el proyecto de campeonato.
3. Ejecuta `setupChampionshipSheets()` y comprueba las siete hojas del campeonato y `SCHEMA_VERSION = 2`.
4. Despliega una nueva versión del campeonato y verifica su URL `/exec`.
5. Actualiza después `codigo.gs` e `Index.html` en el proyecto de cronometraje.
6. Ejecuta `setupSheets()` para crear los backups y completar la migración a schema v5.
7. Inspecciona cabeceras, filas migradas, `MIGRATION_V5_BACKUP` y `SCHEMA_VERSION = 5`.
8. Despliega juntos backend y frontend de crono como una nueva versión.
9. Haz una recarga fuerte y prueba el flujo antes de usar datos reales.

Este orden evita que crono publique un payload v2 contra un receptor de campeonato antiguo.

Guardar código no actualiza una URL ya publicada. En cada proyecto usa:

```text
Implementar -> Gestionar implementaciones -> Editar -> Nueva versión -> Implementar
```

## Verificación recomendada

Antes de una carrera real, usa datos genéricos y verifica:

1. Migración v5 y presencia de backups.
2. Carrera de una pasada: cero descartes y total completo correcto.
3. Carrera con un tramo ida/vuelta y otro solo ida: generación y orden correctos.
4. Bloqueo de registro en `CONFIGURACION` e inicio explícito.
5. Ampliación posterior por pareja o por ida sin cambiar IDs existentes.
6. Rechazo del duplicado `carreraId + inscripcionId + pasadaId`.
7. Penalización, corrección, descarte global y desempate del descarte por orden.
8. Gap solo entre pilotos de la misma categoría con la misma firma de pasadas.
9. Inscripciones sin dorsal y bloqueo de edición o borrado cuando tienen resultados.
10. Administración en móvil y escritorio.
11. Publicación v2 y lectura conjunta de una carrera v2 con histórico v1.

## Solución de problemas

| Síntoma | Comprobación |
|---|---|
| Los cambios no aparecen | Despliega una nueva versión y haz recarga fuerte. |
| La carrera no admite tiempos | Comprueba que `CARRERA_ESTADO` sea `INICIADA`. |
| La configuración cambió | Recarga; las escrituras validan `CARRERA_ID` y `CONFIG_REVISION`. |
| La hoja usa un esquema futuro | No fuerces la migración; despliega una versión de código compatible. |
| La publicación devuelve HTTP 401 | El despliegue de campeonato debe ser accesible para `UrlFetchApp` y ejecutar como propietario. |
| La carrera ya existe en campeonato | El mismo `CARRERA_ID` ya fue importado y no se sobrescribe. |

## Alcance

Están implementados los tramos dinámicos y el campeonato con puntuación y descartes de temporada. Continúan como posibles mejoras futuras las exportaciones y una eventual integración con cronometraje externo. No existe integración funcional con ESP32.

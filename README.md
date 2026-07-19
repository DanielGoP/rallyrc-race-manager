# Rally RC Race Manager

Aplicación web para gestionar carreras de Rally RC usando **Google Apps Script** y **Google Sheets**.

El objetivo del proyecto es sustituir el registro manual en papel por una herramienta sencilla, gratuita y accesible desde móvil, tablet u ordenador. La aplicación permite registrar tiempos, evitar duplicados, consultar la clasificación en tiempo real, aplicar correcciones desde administración y analizar la constancia de los pilotos mediante un dashboard.

> Proyecto pensado para carreras de club, entrenamientos, pruebas internas y eventos de Rally RC donde se necesita una solución rápida sin montar servidores ni bases de datos externas.

---

## Índice

- [Características principales](#características-principales)
- [Tecnologías utilizadas](#tecnologías-utilizadas)
- [Cómo funciona la aplicación](#cómo-funciona-la-aplicación)
- [Modelo de carrera](#modelo-de-carrera)
- [Estructura del repositorio](#estructura-del-repositorio)
- [Estructura de Google Sheets](#estructura-de-google-sheets)
- [Instalación paso a paso](#instalación-paso-a-paso)
- [Configuración inicial](#configuración-inicial)
- [Publicación como aplicación web](#publicación-como-aplicación-web)
- [Uso de la aplicación](#uso-de-la-aplicación)
- [Registro de tiempos](#registro-de-tiempos)
- [Clasificación](#clasificación)
- [Dashboard](#dashboard)
- [Administración](#administración)
- [Nueva carrera](#nueva-carrera)
- [Corrección de resultados](#corrección-de-resultados)
- [Bloqueo de duplicados](#bloqueo-de-duplicados)
- [Cálculo de clasificación](#cálculo-de-clasificación)
- [Cálculo de gap](#cálculo-de-gap)
- [Formato de tiempos](#formato-de-tiempos)
- [Pruebas recomendadas](#pruebas-recomendadas)
- [Solución de problemas](#solución-de-problemas)
- [Roadmap](#roadmap)

---

## Características principales

La aplicación incluye:

- Acceso mediante PIN.
- Registro de juez o responsable de mesa.
- Registro de tiempos por categoría, piloto y pasada.
- Gestión de seis pasadas por piloto:
  - Ida 1
  - Vuelta 1
  - Ida 2
  - Vuelta 2
  - Ida 3
  - Vuelta 3
- Entrada de tiempos mediante campos separados:
  - minutos;
  - segundos;
  - décimas.
- Penalizaciones en segundos.
- Bloqueo de resultados duplicados.
- Marcado visual de pasadas ya registradas.
- Colores diferenciados para pasadas de ida y vuelta.
- Clasificación en tiempo real.
- Descarte automático de la peor pasada cuando un piloto tiene seis resultados.
- Visualización de la pasada descartada tachada.
- Alternancia entre visualización en segundos y formato `minutos:segundos,décimas`.
- Cálculo de gap con el piloto anterior comparable.
- Panel de administración protegido con PIN de administración.
- Corrección de resultados existentes.
- Generación de nueva carrera con backup previo.
- Dashboard de constancia con métricas y gráfica.
- Funcionamiento sin servidor propio.

---

## Tecnologías utilizadas

El proyecto utiliza:

- **Google Apps Script** como backend.
- **Google Sheets** como almacenamiento de datos.
- **HTML, CSS y JavaScript** para la interfaz.
- **Canvas HTML** para pintar gráficas en el dashboard.

No requiere:

- servidor propio;
- base de datos externa;
- hosting adicional;
- dependencias npm;
- frameworks frontend.

---

## Cómo funciona la aplicación

La aplicación se ejecuta como una **Web App de Google Apps Script**.

El flujo general es:

```text
Usuario entra en la URL pública
        ↓
Introduce PIN de acceso y nombre de juez
        ↓
Selecciona una pestaña de la aplicación
        ↓
Registra tiempos o consulta resultados
        ↓
Apps Script lee/escribe en Google Sheets
        ↓
La clasificación se recalcula automáticamente
```

Google Sheets actúa como base de datos. Cada pestaña tiene una responsabilidad concreta:

```text
Config          → configuración general
Inscripciones   → pilotos inscritos
Resultados      → tiempos registrados
Clasificacion   → clasificación calculada
```

---

## Modelo de carrera

El modelo actual está pensado para una **carrera individual**.

Cada piloto puede participar en una o varias categorías. Para evitar mezclar tiempos, la aplicación trabaja con el concepto de **inscripción**.

Una inscripción representa la participación de un piloto en una categoría concreta.

Ejemplo genérico:

| inscripcionId | dorsal | piloto        | categoria |
|---------------|--------|---------------|-----------|
| I001          | 1      | Piloto Demo 1 | Rally 1/10 |
| I002          | 2      | Piloto Demo 2 | Rally 1/10 |
| I003          | 1      | Piloto Demo 1 | Clásicos |

En este ejemplo, `Piloto Demo 1` participa en dos categorías diferentes, por lo que tiene dos inscripciones.

---

## Estructura del repositorio

El repositorio contiene principalmente:

```text
rallyrc-race-manager/
├── codigo.gs
├── index.html
└── README.md
```

### `codigo.gs`

Contiene la lógica de servidor:

- creación de pestañas;
- configuración inicial;
- validación de PIN;
- validación de PIN de administración;
- lectura de categorías;
- lectura de inscripciones;
- registro de resultados;
- bloqueo de duplicados;
- cálculo de clasificación;
- descarte de peor pasada;
- cálculo de gap;
- generación de nueva carrera;
- corrección de resultados;
- generación de datos para el dashboard.

### `index.html`

Contiene la aplicación web:

- pantalla de acceso;
- navegación;
- registro de tiempos;
- clasificación;
- dashboard;
- administración;
- estilos CSS;
- JavaScript cliente.

---

## Estructura de Google Sheets

La aplicación crea y utiliza las siguientes pestañas.

---

### `Config`

Pestaña de configuración general.

Ejemplo:

| key | value |
|-----|-------|
| PIN | CAMBIAR_PIN_ACCESO |
| ADMIN_PIN | CAMBIAR_PIN_ADMIN |
| CARRERA_ACTIVA | Rally RC |

Campos:

| Campo | Descripción |
|------|-------------|
| `PIN` | PIN usado para entrar en la aplicación. |
| `ADMIN_PIN` | PIN usado para acceder a administración. |
| `CARRERA_ACTIVA` | Nombre de la carrera actual. |

> Recomendación: cambiar siempre los valores de ejemplo antes de publicar la aplicación.

---

### `Inscripciones`

Pestaña donde se registran los pilotos inscritos.

Ejemplo:

| inscripcionId | dorsal | piloto | categoria |
|---------------|--------|--------|-----------|
| I001 | 1 | Piloto Demo 1 | Rally 1/10 |
| I002 | 2 | Piloto Demo 2 | Rally 1/10 |
| I003 | 3 | Piloto Demo 3 | Rally 1/10 |
| I004 | 1 | Piloto Demo 1 | Clásicos |

Campos:

| Campo | Descripción |
|------|-------------|
| `inscripcionId` | Identificador único de la inscripción. |
| `dorsal` | Número de dorsal del piloto. |
| `piloto` | Nombre del piloto. |
| `categoria` | Categoría en la que participa. |

Recomendaciones:

- `inscripcionId` debe ser único.
- Un piloto que participe en dos categorías debe tener dos inscripciones.
- Las categorías deben escribirse siempre igual.
- Se pueden usar dorsales repetidos si pertenecen a categorías distintas.

---

### `Resultados`

Pestaña donde se guarda cada tiempo registrado.

Ejemplo:

| resultadoId | timestamp | inscripcionId | dorsal | piloto | categoria | pasada | pasadaLabel | tiempo | penalizacion | total | juez |
|------------|-----------|---------------|--------|--------|-----------|--------|-------------|--------|--------------|-------|------|
| uuid | fecha/hora | I001 | 1 | Piloto Demo 1 | Rally 1/10 | 1 | Ida 1 | 70.5 | 0 | 70.5 | Mesa |

Campos:

| Campo | Descripción |
|------|-------------|
| `resultadoId` | Identificador único del resultado. |
| `timestamp` | Fecha y hora del registro. |
| `inscripcionId` | Inscripción asociada. |
| `dorsal` | Dorsal del piloto. |
| `piloto` | Nombre del piloto. |
| `categoria` | Categoría. |
| `pasada` | Número de pasada. |
| `pasadaLabel` | Nombre visible de la pasada. |
| `tiempo` | Tiempo base en segundos. |
| `penalizacion` | Penalización en segundos. |
| `total` | Tiempo base + penalización. |
| `juez` | Persona que registró el tiempo. |

---

### `Clasificacion`

Pestaña con la clasificación calculada.

Ejemplo:

| categoria | dorsal | piloto | ida1 | vuelta1 | ida2 | vuelta2 | ida3 | vuelta3 | descartada | penalizaciones | total | gap | estado |
|----------|--------|--------|------|---------|------|---------|------|---------|------------|----------------|-------|-----|--------|
| Rally 1/10 | 1 | Piloto Demo 1 | 70.5 | 72.3 | 69.8 | 75 | 68.4 | 80.2 | Vuelta 3 - 80.2 | 0 | 356 | - | Completo |

Campos:

| Campo | Descripción |
|------|-------------|
| `categoria` | Categoría. |
| `dorsal` | Dorsal. |
| `piloto` | Piloto. |
| `ida1` | Tiempo total de Ida 1. |
| `vuelta1` | Tiempo total de Vuelta 1. |
| `ida2` | Tiempo total de Ida 2. |
| `vuelta2` | Tiempo total de Vuelta 2. |
| `ida3` | Tiempo total de Ida 3. |
| `vuelta3` | Tiempo total de Vuelta 3. |
| `descartada` | Pasada descartada si hay seis tiempos. |
| `penalizaciones` | Suma de penalizaciones. |
| `total` | Total de clasificación. |
| `gap` | Diferencia con el piloto anterior comparable. |
| `estado` | Estado del piloto. |

---

## Instalación paso a paso

### 1. Crear una hoja de cálculo

Crea un nuevo Google Sheet en Google Drive.

Nombre sugerido:

```text
Rally RC Race Manager
```

---

### 2. Abrir Apps Script

Desde el Google Sheet:

```text
Extensiones → Apps Script
```

---

### 3. Crear el archivo `codigo.gs`

En Apps Script, crea o renombra el archivo principal como:

```text
codigo.gs
```

Pega el contenido del archivo `codigo.gs` del repositorio.

---

### 4. Crear el archivo `index.html`

En Apps Script:

```text
+ → HTML
```

Crea un archivo llamado:

```text
Index
```

Pega el contenido del archivo `index.html` del repositorio.

> Nota: el backend utiliza `HtmlService.createTemplateFromFile('Index')`, por lo que el archivo HTML debe llamarse `Index.html` dentro de Apps Script.

---

### 5. Guardar el proyecto

Guarda los cambios con:

```text
Ctrl + S
```

o con el botón de guardar.

---

### 6. Ejecutar la configuración inicial

En el desplegable de funciones del editor de Apps Script, selecciona:

```text
setupSheets
```

Pulsa:

```text
Ejecutar
```

La primera vez Google pedirá autorización.

Acepta los permisos para que el script pueda leer y escribir en la hoja de cálculo.

---

### 7. Comprobar las pestañas

Vuelve al Google Sheet y comprueba que existen:

```text
Config
Inscripciones
Resultados
Clasificacion
```

---

## Configuración inicial

### Cambiar PIN de acceso

En `Config`, cambia el valor de:

```text
PIN
```

por un valor propio.

Ejemplo recomendado:

| key | value |
|-----|-------|
| PIN | TU_PIN_DE_ACCESO |

---

### Cambiar PIN de administración

En `Config`, cambia el valor de:

```text
ADMIN_PIN
```

por un valor propio.

Ejemplo recomendado:

| key | value |
|-----|-------|
| ADMIN_PIN | TU_PIN_DE_ADMINISTRACION |

---

### Cambiar nombre de carrera

En `Config`:

| key | value |
|-----|-------|
| CARRERA_ACTIVA | Nombre de la carrera |

Ejemplo:

```text
Rally RC Club - Carrera 1
```

---

### Añadir pilotos

En `Inscripciones`, añade los pilotos que participarán en la carrera.

Ejemplo:

| inscripcionId | dorsal | piloto | categoria |
|---------------|--------|--------|-----------|
| I001 | 1 | Piloto Demo 1 | Rally 1/10 |
| I002 | 2 | Piloto Demo 2 | Rally 1/10 |
| I003 | 3 | Piloto Demo 3 | Rally 1/10 |
| I004 | 1 | Piloto Demo 1 | Clásicos |

---

## Publicación como aplicación web

En Apps Script:

1. Pulsa:

```text
Implementar → Nueva implementación
```

2. Selecciona el tipo:

```text
Aplicación web
```

3. Configura:

```text
Ejecutar como: Yo
Quién tiene acceso: Cualquier usuario con el enlace
```

4. Pulsa:

```text
Implementar
```

5. Copia la URL generada.

Esa URL será la aplicación web.

---

### Actualizar una aplicación ya publicada

Cuando hagas cambios en el código, no basta con guardar.

Debes publicar una nueva versión:

```text
Implementar → Gestionar implementaciones → Editar → Nueva versión → Implementar
```

Después, recarga la aplicación web.

---

## Uso de la aplicación

Al abrir la URL de la aplicación aparece la pantalla de acceso.

El usuario debe introducir:

- PIN de acceso;
- nombre del juez o responsable de mesa.

Después se muestran las secciones principales:

```text
Registrar pasada
Clasificación
Dashboard
Administración
```

---

## Registro de tiempos

La pestaña **Registrar pasada** permite introducir tiempos.

Flujo habitual:

1. Seleccionar categoría.
2. Seleccionar piloto.
3. Seleccionar pasada.
4. Introducir minutos.
5. Introducir segundos.
6. Introducir décimas.
7. Introducir penalización si aplica.
8. Guardar resultado.

---

### Pasadas disponibles

| Número | Pasada |
|--------|--------|
| 1 | Ida 1 |
| 2 | Vuelta 1 |
| 3 | Ida 2 |
| 4 | Vuelta 2 |
| 5 | Ida 3 |
| 6 | Vuelta 3 |

Las pasadas de ida y vuelta se muestran con colores diferenciados.

Cuando una pasada ya está registrada, aparece marcada visualmente para evitar errores.

---

### Entrada de tiempo

Ejemplo:

```text
1 minuto, 10 segundos y 5 décimas
```

Se introduce como:

| Campo | Valor |
|------|-------|
| Minutos | 1 |
| Segundos | 10 |
| Décimas | 5 |

La aplicación lo convierte internamente a:

```text
70.5 segundos
```

---

### Penalizaciones

La penalización se introduce en segundos.

Ejemplo:

```text
Tiempo base: 60.0
Penalización: 5.0
Total: 65.0
```

La clasificación utiliza el campo `total`.

---

### Validaciones

La aplicación valida:

- que haya una categoría seleccionada;
- que haya un piloto seleccionado;
- que haya una pasada seleccionada;
- que los minutos no sean negativos;
- que los segundos estén entre 0 y 59;
- que las décimas estén entre 0 y 9;
- que la penalización no sea negativa;
- que no exista ya un resultado para esa inscripción y pasada.

---

## Clasificación

La pestaña **Clasificación** muestra los resultados en tiempo real.

La tabla incluye:

- categoría;
- dorsal;
- piloto;
- tiempos de cada pasada;
- total;
- gap;
- estado.

---

### Cambio de formato

La clasificación puede alternar entre:

```text
Segundos
```

y:

```text
Minutos:segundos,décimas
```

Ejemplo:

| Segundos | Formato visible |
|---------|-----------------|
| 70.5 | 1:10,5 |
| 356 | 5:56,0 |

---

### Pasada descartada

Cuando un piloto tiene seis resultados, la aplicación descarta automáticamente la peor pasada.

Ejemplo:

```text
Ida 1: 70.5
Vuelta 1: 72.3
Ida 2: 69.8
Vuelta 2: 75.0
Ida 3: 68.4
Vuelta 3: 80.2
```

Cálculo:

```text
Suma de todas las pasadas = 436.2
Peor pasada = 80.2
Total final = 436.2 - 80.2 = 356.0
```

La pasada descartada se muestra tachada en la tabla.

---

## Dashboard

La pestaña **Dashboard** permite analizar la regularidad de los pilotos.

Está enfocada en rendimiento y constancia, no solo en clasificación.

---

### Filtros

El dashboard permite filtrar por:

- categoría;
- piloto;
- tipo de pasada:
  - todas;
  - solo idas;
  - solo vueltas;
- orden:
  - constancia;
  - media;
  - mejor tiempo;
  - pasadas completadas;
  - dorsal.

---

### Métricas disponibles

| Métrica | Descripción |
|--------|-------------|
| Pilotos visibles | Número de pilotos que coinciden con el filtro. |
| Más constante | Piloto con menor diferencia entre peor y mejor tiempo. |
| Mejor tiempo | Mejor tiempo individual registrado. |
| Mejor media | Mejor media de tiempos. |

---

### Constancia

La constancia se calcula así:

```text
Constancia = peor tiempo - mejor tiempo
```

Ejemplo:

```text
Tiempos: 70.5, 72.3, 69.8
Mejor: 69.8
Peor: 72.3
Constancia: 2.5
```

Cuanto menor es la diferencia, más constante ha sido el piloto.

---

### Gráfica

La gráfica del dashboard muestra los tiempos individuales por pasada.

No es acumulada.

Ejemplo:

```text
Ida 1    → 70.5
Vuelta 1 → 72.3
Ida 2    → 69.8
Vuelta 2 → 75.0
Ida 3    → 68.4
Vuelta 3 → 80.2
```

---

## Administración

La pestaña **Administración** está protegida con PIN de administración.

Permite:

- generar una nueva carrera;
- corregir resultados existentes.

---

## Nueva carrera

La opción **Nueva carrera** prepara la aplicación para una nueva prueba.

Antes de limpiar los datos, la app crea backups de:

```text
Resultados
Clasificacion
```

Ejemplo de pestañas de backup:

```text
Resultados_backup_YYYY-MM-DD_HH-mm-ss
Clasificacion_backup_YYYY-MM-DD_HH-mm-ss
```

Después:

- limpia la pestaña `Resultados`;
- limpia la pestaña `Clasificacion`;
- actualiza `CARRERA_ACTIVA`;
- recalcula la clasificación vacía.

Para evitar borrados accidentales, se solicita una confirmación manual.

---

## Corrección de resultados

Desde administración se puede modificar un resultado ya registrado.

Flujo:

1. Entrar en administración.
2. Introducir PIN de administración.
3. Seleccionar categoría.
4. Seleccionar piloto.
5. Seleccionar pasada.
6. Cargar resultado.
7. Modificar tiempo o penalización.
8. Confirmar la corrección.
9. Guardar.

La corrección:

- modifica la fila existente;
- no crea duplicados;
- recalcula la clasificación.

---

## Bloqueo de duplicados

La app impide registrar dos veces la misma pasada para la misma inscripción.

La clave lógica es:

```text
inscripcionId + pasada
```

Ejemplo:

```text
I001 + Ida 1
```

Si ya existe un resultado, la aplicación no guarda el nuevo registro y muestra un aviso.

Mensaje esperado:

```text
Ya existe un resultado para esta pasada. No se ha guardado ningún cambio.
```

---

## Cálculo de clasificación

### Sin resultados

Si un piloto no tiene tiempos:

```text
total = vacío
estado = Sin resultados
```

---

### Entre 1 y 5 resultados

Si un piloto tiene entre una y cinco pasadas:

```text
total = suma de pasadas registradas
estado = Pendiente
```

Si tiene penalización:

```text
estado = Pendiente con penalización
```

---

### Con 6 resultados

Si un piloto tiene las seis pasadas:

```text
se descarta la peor pasada
total = suma de las cinco mejores
estado = Completo
```

Si tiene penalización:

```text
estado = Completo con penalización
```

---

### Ordenación

La clasificación se ordena por:

1. pilotos con total visible;
2. total de menor a mayor;
3. categoría;
4. dorsal.

Esto permite que la clasificación se actualice durante toda la carrera sin esperar a que todos los pilotos tengan las seis pasadas.

---

## Cálculo de gap

La columna `Gap` muestra la diferencia con el piloto inmediatamente anterior, pero solo si ambos pilotos son comparables.

La regla es:

```text
El gap solo se calcula si ambos pilotos tienen el mismo número de pasadas completadas.
```

Ejemplo:

| Posición | Piloto | Pasadas | Total | Gap |
|----------|--------|---------|-------|-----|
| 1 | Piloto Demo 1 | 2/6 | 126 | - |
| 2 | Piloto Demo 2 | 6/6 | 351 | - |
| 3 | Piloto Demo 3 | 6/6 | 356 | 5 |

En este ejemplo:

- `Piloto Demo 2` no tiene gap contra `Piloto Demo 1` porque no llevan las mismas pasadas.
- `Piloto Demo 3` sí tiene gap contra `Piloto Demo 2` porque ambos llevan 6/6.

---

## Formato de tiempos

Internamente todos los tiempos se guardan en segundos.

Ejemplos:

| Entrada | Valor interno |
|--------|---------------|
| 1:10,5 | 70.5 |
| 1:12,3 | 72.3 |
| 5:56,0 | 356 |

La app puede mostrar estos valores como:

```text
70.5
```

o como:

```text
1:10,5
```

---

## Pruebas recomendadas

Antes de usar la app en una carrera real, se recomienda hacer una prueba completa.

---

### Prueba de acceso

Comprobar:

- acceso con PIN correcto;
- rechazo con PIN incorrecto;
- visualización de pestañas.

---

### Prueba de registro

Registrar seis pasadas para un piloto de prueba:

| Pasada | Tiempo |
|--------|--------|
| Ida 1 | 1:10,5 |
| Vuelta 1 | 1:12,3 |
| Ida 2 | 1:09,8 |
| Vuelta 2 | 1:15,0 |
| Ida 3 | 1:08,4 |
| Vuelta 3 | 1:20,2 |

Validar:

- los seis tiempos se guardan;
- la clasificación se recalcula;
- la peor pasada se descarta;
- la pasada descartada aparece tachada;
- el total es correcto.

---

### Prueba de duplicados

Intentar registrar dos veces la misma pasada.

Resultado esperado:

```text
La app bloquea el segundo registro.
```

---

### Prueba de penalización

Registrar un tiempo con penalización.

Ejemplo:

```text
Tiempo: 1:00,0
Penalización: 5
Total esperado: 65.0
```

---

### Prueba de corrección

Desde administración:

- cargar un resultado existente;
- modificarlo;
- guardar;
- comprobar que la clasificación se recalcula.

---

### Prueba de dashboard

Comprobar:

- carga de datos;
- filtro por categoría;
- filtro por piloto;
- filtro por idas/vueltas;
- cálculo de constancia;
- gráfica;
- tabla.

---

## Solución de problemas

### Los cambios no aparecen en la app

Guardar el código no siempre actualiza la web app publicada.

Hay que generar una nueva versión:

```text
Implementar → Gestionar implementaciones → Editar → Nueva versión → Implementar
```

---

### La pantalla queda en blanco

Abrir la consola del navegador.

Causa habitual:

```text
ReferenceError: nombreFuncion is not defined
```

Suele ocurrir si se ha pegado una parte del código pero falta alguna función auxiliar.

---

### El selector de pilotos se queda cargando

Revisar:

- que la aplicación esté desplegada con la última versión;
- que existan pilotos en la categoría;
- que `getInscripcionesByCategoria` no esté fallando;
- los registros de ejecución de Apps Script.

---

### No se guarda un resultado

Comprobar:

- PIN correcto;
- categoría seleccionada;
- piloto seleccionado;
- pasada seleccionada;
- tiempo válido;
- penalización válida;
- que no exista duplicado.

---

### Error de permisos

La primera vez que se ejecuta el proyecto, Google puede pedir autorización.

Hay que aceptar los permisos para que Apps Script pueda leer y escribir en el Google Sheet.

---

## Roadmap

Ideas futuras posibles:

### Tramos

Permitir varias secciones o tramos dentro de una carrera.

```text
Tramo 1
Tramo 2
Tramo 3
```

---

### Campeonato

Permitir agrupar varias carreras dentro de un campeonato.

```text
Campeonato 2026
├── Carrera 1
├── Carrera 2
└── Carrera 3
```

---

### Puntuación por carrera

Asignar puntos por posición.

Ejemplo:

| Posición | Puntos |
|----------|--------|
| 1º | 25 |
| 2º | 18 |
| 3º | 15 |
| 4º | 12 |
| 5º | 10 |

---

### Descartes de campeonato

Permitir descartar el peor resultado de una temporada.

---

### Cronometraje externo

Posible integración futura con sistemas externos de cronometraje.

No forma parte de la funcionalidad implementada actualmente.

---

### Exportación

Posibles exportaciones futuras:

- CSV;
- Excel;
- PDF;
- imagen para compartir resultados.

---

## Estado actual

La aplicación permite gestionar una carrera individual de Rally RC con:

- registro de tiempos;
- clasificación;
- descarte de peor pasada;
- gap;
- administración;
- correcciones;
- dashboard de constancia.

No incluye actualmente integración real con ESP32 ni sistemas de cronometraje externos.
